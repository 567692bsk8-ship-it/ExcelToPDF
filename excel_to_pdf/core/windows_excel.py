"""
High-performance Microsoft Excel COM converter for Windows.
Optimized for multi-process parallel execution and minimal latency.
"""

from __future__ import annotations

import logging
import os
import sys
from pathlib import Path
from typing import Any

from excel_to_pdf.core.converter_base import BaseConverter

logger = logging.getLogger(__name__)

# xlFixedFormatType constants
XL_TYPE_PDF = 0
XL_CALCULATION_MANUAL = -4135


class WindowsExcelConverter(BaseConverter):
    """
    Windows-native Microsoft Excel COM automation converter.
    Disables screen updates, alerts, events, automatic calculation, and link prompts
    to achieve sub-second conversion speeds per document.
    """

    def __init__(self, excel_app: Any = None) -> None:
        self._external_app = excel_app
        self._owned_app: Any = None

    @classmethod
    def is_available(cls) -> bool:
        if sys.platform != "win32":
            return False
        try:
            import win32com.client  # type: ignore
            return True
        except ImportError:
            return False

    @classmethod
    def engine_name(cls) -> str:
        return "Microsoft Excel (Windows COM)"

    def _get_app(self) -> Any:
        if self._external_app is not None:
            return self._external_app

        if self._owned_app is None:
            import pythoncom  # type: ignore
            import win32com.client  # type: ignore

            pythoncom.CoInitialize()
            app = win32com.client.DispatchEx("Excel.Application")
            # Apply maximum performance optimizations
            app.Visible = False
            app.ScreenUpdating = False
            app.DisplayAlerts = False
            app.EnableEvents = False
            try:
                app.AskToUpdateLinks = False
                app.Calculation = XL_CALCULATION_MANUAL
            except Exception:
                pass
            self._owned_app = app

        return self._owned_app

    def close(self) -> None:
        """Safely release the owned Excel COM application."""
        if self._owned_app is not None:
            try:
                self._owned_app.Quit()
            except Exception:
                pass
            self._owned_app = None
            try:
                import pythoncom  # type: ignore
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def __enter__(self) -> WindowsExcelConverter:
        return self

    def __exit__(self, exc_type: Any, exc_val: Any, exc_tb: Any) -> None:
        self.close()

    def convert_file(self, excel_path: Path, output_pdf_path: Path) -> tuple[bool, str | None]:
        """Convert a single Excel file to PDF using COM automation."""
        if sys.platform != "win32":
            return False, "Windows COM engine is only available on Windows."

        excel_abs = str(excel_path.resolve())
        pdf_abs = str(output_pdf_path.resolve())

        # Ensure parent output directory exists
        output_pdf_path.parent.mkdir(parents=True, exist_ok=True)

        app = None
        workbook = None
        try:
            app = self._get_app()
            # Open with read-only and no link update to bypass prompts & delay
            workbook = app.Workbooks.Open(
                excel_abs,
                UpdateLinks=0,
                ReadOnly=True,
                IgnoreReadOnlyRecommended=True,
                Notify=False,
            )

            # Export to PDF (Type 0 = xlTypePDF)
            workbook.ExportAsFixedFormat(
                Type=XL_TYPE_PDF,
                Filename=pdf_abs,
                Quality=0,  # xlQualityStandard
                IncludeDocProperties=True,
                IgnorePrintAreas=False,
                OpenAfterPublish=False,
            )
            return True, None
        except Exception as e:
            return False, f"Excel COM Error: {str(e)}"
        finally:
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass


# --- Multiprocessing Worker Entry Point ---

def run_windows_worker_batch(
    file_pairs: list[tuple[str, str]],
    progress_queue: Any = None,
) -> list[tuple[str, bool, str | None]]:
    """
    Worker process target:
    Initializes a single Excel instance for this process, converts multiple files sequentially,
    and cleanly releases resources upon completion.
    Reports progress per file via progress_queue if provided.
    """
    results: list[tuple[str, bool, str | None]] = []
    if sys.platform != "win32":
        for src, _ in file_pairs:
            err = "Not running on Windows"
            results.append((src, False, err))
            if progress_queue is not None:
                try:
                    progress_queue.put((src, False, err))
                except Exception:
                    pass
        return results

    converter = WindowsExcelConverter()
    try:
        for src_path_str, dst_path_str in file_pairs:
            src = Path(src_path_str)
            dst = Path(dst_path_str)
            success, err = converter.convert_file(src, dst)
            results.append((src_path_str, success, err))
            if progress_queue is not None:
                try:
                    progress_queue.put((src_path_str, success, err))
                except Exception:
                    pass
    finally:
        converter.close()

    return results
