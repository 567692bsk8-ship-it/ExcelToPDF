"""
Detection of available conversion engines on the host system.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path


class EngineDetector:
    """Detects available conversion engines on Windows, macOS, and Linux."""

    @staticmethod
    def get_libreoffice_path() -> str | None:
        """Find the executable path for LibreOffice (soffice)."""
        # 1. Check PATH
        cmd = shutil.which("soffice") or shutil.which("libreoffice")
        if cmd:
            return cmd

        # 2. Check platform-specific common paths
        if sys.platform == "darwin":
            mac_paths = [
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
            ]
            for p in mac_paths:
                if os.path.exists(p) and os.access(p, os.X_OK):
                    return p

        elif sys.platform == "win32":
            win_paths = [
                r"C:\Program Files\LibreOffice\program\soffice.exe",
                r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            ]
            for p in win_paths:
                if os.path.exists(p):
                    return p

        elif sys.platform.startswith("linux"):
            linux_paths = [
                "/usr/bin/soffice",
                "/usr/bin/libreoffice",
                "/usr/local/bin/soffice",
                "/snap/bin/libreoffice",
            ]
            for p in linux_paths:
                if os.path.exists(p) and os.access(p, os.X_OK):
                    return p

        return None

    @staticmethod
    def is_windows_excel_available() -> bool:
        """Check if Microsoft Excel COM is available (Windows only)."""
        if sys.platform != "win32":
            return False

        try:
            import win32com.client  # type: ignore
            # Test creating an application object
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Quit()
            return True
        except Exception:
            return False

    @classmethod
    def get_available_engines(cls) -> list[str]:
        """Return a list of available engine identifiers."""
        engines: list[str] = []
        if cls.is_windows_excel_available():
            engines.append("windows_excel")
        if cls.get_libreoffice_path() is not None:
            engines.append("libreoffice")
        return engines

    @classmethod
    def get_best_engine(cls) -> str | None:
        """Select the best available engine. Prioritizes MS Excel on Windows, LibreOffice elsewhere."""
        if cls.is_windows_excel_available():
            return "windows_excel"
        if cls.get_libreoffice_path() is not None:
            return "libreoffice"
        return None
