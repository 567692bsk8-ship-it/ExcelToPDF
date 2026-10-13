"""
Hardware and OS specification detector for Windows and macOS.
Zero external dependencies (uses standard library and OS APIs).
"""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from excel_to_pdf.core.engine_detector import EngineDetector


@dataclass
class SystemSpec:
    os_name: str
    os_version: str
    architecture: str
    cpu_model: str
    cpu_cores_logical: int
    cpu_cores_physical: int
    ram_total_gb: float
    ram_available_gb: float | None
    disk_total_gb: float
    disk_free_gb: float
    python_version: str
    engine_name: str
    engine_path: str | None
    recommended_workers: int
    recommendation_comment: str


class SystemSpecDetector:
    """Collects hardware and OS specifications for Windows and macOS."""

    @classmethod
    def get_specs(cls) -> SystemSpec:
        os_name, os_version = cls._get_os_info()
        arch = platform.machine() or platform.architecture()[0]

        cpu_model = cls._get_cpu_model()
        logical_cores = os.cpu_count() or 1
        physical_cores = cls._get_physical_cores(logical_cores)

        ram_total, ram_avail = cls._get_memory_info()
        disk_total, disk_free = cls._get_disk_info()

        # Engine detection
        best_engine = EngineDetector.get_best_engine()
        engine_path = None
        if best_engine == "windows_excel":
            engine_name = "Microsoft Excel COM (超高速・Windowsネイティブ)"
            engine_path = "Excel.Application"
        elif best_engine == "libreoffice":
            engine_name = "LibreOffice ヘッドレス (クロスプラットフォーム)"
            engine_path = EngineDetector.get_libreoffice_path()
        else:
            engine_name = "未検出 (Excel または LibreOffice が必要です)"

        # Calculate recommended workers based on CPU cores & RAM
        recommended_workers, comment = cls._calc_recommended_workers(logical_cores, ram_total)

        py_ver = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"

        return SystemSpec(
            os_name=os_name,
            os_version=os_version,
            architecture=arch,
            cpu_model=cpu_model,
            cpu_cores_logical=logical_cores,
            cpu_cores_physical=physical_cores,
            ram_total_gb=round(ram_total, 1),
            ram_available_gb=round(ram_avail, 1) if ram_avail is not None else None,
            disk_total_gb=round(disk_total, 1),
            disk_free_gb=round(disk_free, 1),
            python_version=py_ver,
            engine_name=engine_name,
            engine_path=engine_path,
            recommended_workers=recommended_workers,
            recommendation_comment=comment,
        )

    @classmethod
    def _get_os_info(cls) -> tuple[str, str]:
        if sys.platform == "darwin":
            mac_ver = platform.mac_ver()[0] or platform.release()
            return "macOS", mac_ver
        elif sys.platform == "win32":
            win_ver = platform.version()
            release = platform.release()
            return f"Windows {release}", win_ver
        else:
            return "Linux", platform.release()

    @classmethod
    def _get_cpu_model(cls) -> str:
        if sys.platform == "darwin":
            try:
                out = subprocess.check_output(
                    ["sysctl", "-n", "machdep.cpu.brand_string"],
                    stderr=subprocess.DEVNULL,
                    text=True,
                ).strip()
                if out:
                    return out
            except Exception:
                pass
            return f"Apple Silicon ({platform.processor()})" if "arm" in platform.machine().lower() else platform.processor()

        elif sys.platform == "win32":
            try:
                import winreg

                key = winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE,
                    r"HARDWARE\DESCRIPTION\System\CentralProcessor\0",
                )
                val, _ = winreg.QueryValueEx(key, "ProcessorNameString")
                winreg.CloseKey(key)
                if val:
                    return str(val).strip()
            except Exception:
                pass
            return platform.processor() or "Generic x86_64 CPU"

        return platform.processor() or "Unknown CPU"

    @classmethod
    def _get_physical_cores(cls, logical_cores: int) -> int:
        if sys.platform == "darwin":
            try:
                out = subprocess.check_output(
                    ["sysctl", "-n", "hw.physicalcpu"],
                    stderr=subprocess.DEVNULL,
                    text=True,
                ).strip()
                if out.isdigit():
                    return int(out)
            except Exception:
                pass
        # Fallback approximation for hyperthreaded CPUs
        return max(1, logical_cores // 2 if logical_cores > 2 else logical_cores)

    @classmethod
    def _get_memory_info(cls) -> tuple[float, float | None]:
        total_gb = 8.0
        avail_gb = None

        if sys.platform == "darwin":
            try:
                out = subprocess.check_output(
                    ["sysctl", "-n", "hw.memsize"],
                    stderr=subprocess.DEVNULL,
                    text=True,
                ).strip()
                if out.isdigit():
                    total_gb = int(out) / (1024**3)
            except Exception:
                pass
            # macOS memory pressure doesn't easily expose free memory without vm_stat parsing
            return total_gb, None

        elif sys.platform == "win32":
            try:
                import ctypes

                class MEMORYSTATUSEX(ctypes.Structure):
                    _fields_ = [
                        ("dwLength", ctypes.c_ulong),
                        ("dwMemoryLoad", ctypes.c_ulong),
                        ("ullTotalPhys", ctypes.c_ulonglong),
                        ("ullAvailPhys", ctypes.c_ulonglong),
                        ("ullTotalPageFile", ctypes.c_ulonglong),
                        ("ullAvailPageFile", ctypes.c_ulonglong),
                        ("ullTotalVirtual", ctypes.c_ulonglong),
                        ("ullAvailVirtual", ctypes.c_ulonglong),
                        ("ullAvailExtendedVirtual", ctypes.c_ulonglong),
                    ]

                stat = MEMORYSTATUSEX()
                stat.dwLength = ctypes.sizeof(MEMORYSTATUSEX)
                if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(stat)):
                    total_gb = stat.ullTotalPhys / (1024**3)
                    avail_gb = stat.ullAvailPhys / (1024**3)
                    return total_gb, avail_gb
            except Exception:
                pass

        return total_gb, avail_gb

    @classmethod
    def _get_disk_info(cls) -> tuple[float, float]:
        path = "C:\\" if sys.platform == "win32" else "/"
        try:
            total, _, free = shutil.disk_usage(path)
            return total / (1024**3), free / (1024**3)
        except Exception:
            return 0.0, 0.0

    @classmethod
    def _calc_recommended_workers(cls, logical_cores: int, ram_gb: float) -> tuple[int, str]:
        # Each office process consumes ~150-250MB RAM
        # Balance between CPU concurrency and memory headroom
        if logical_cores <= 2 or ram_gb < 4:
            return 2, "省電力・低負荷設定（2並列）。他の作業への影響を最小限に抑えます。"
        elif logical_cores <= 4 or ram_gb <= 8:
            return 4, "標準バランス設定（4並列）。普段のPC利用と高速変換を両立します。"
        elif logical_cores <= 8 or ram_gb <= 16:
            rec = min(logical_cores, 8)
            return rec, f"高速推奨設定（{rec}並列）。CPU性能を活かして大量ファイルを短時間で一括変換します。"
        else:
            rec = min(logical_cores, 12)
            return rec, f"超高速ハイエンド設定（{rec}並列）。高スペックPCの性能をフル活用して最速処理します。"
