"""
Cross-platform LibreOffice (soffice) converter.
Uses isolated UserInstallation profiles to enable 100% concurrent multi-process conversions.
"""

from __future__ import annotations

import logging
import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import Any

from excel_to_pdf.core.converter_base import BaseConverter
from excel_to_pdf.core.engine_detector import EngineDetector

logger = logging.getLogger(__name__)


class LibreOfficeConverter(BaseConverter):
    """
    Converts Excel documents to PDF via LibreOffice in headless mode.
    Creates an isolated temporary UserInstallation per process/session to prevent profile lock collisions.
    """

    def __init__(self, soffice_binary: str | None = None) -> None:
        self.soffice_binary = soffice_binary or EngineDetector.get_libreoffice_path()

    @classmethod
    def is_available(cls) -> bool:
        return EngineDetector.get_libreoffice_path() is not None

    @classmethod
    def engine_name(cls) -> str:
        return "LibreOffice (Headless Multi-Profile)"

    def convert_file(
        self,
        excel_path: Path,
        output_pdf_path: Path,
        custom_profile_dir: Path | None = None,
    ) -> tuple[bool, str | None]:
        """Convert a single Excel file to PDF using an isolated LibreOffice profile."""
        if not self.soffice_binary or not os.path.exists(self.soffice_binary):
            return False, f"LibreOffice binary not found at '{self.soffice_binary}'."

        excel_path = excel_path.resolve()
        output_pdf_path = output_pdf_path.resolve()
        output_pdf_path.parent.mkdir(parents=True, exist_ok=True)

        # Create isolated profile directory if not provided
        created_temp_profile = False
        if custom_profile_dir is None:
            profile_dir = Path(tempfile.mkdtemp(prefix="soffice_prof_"))
            created_temp_profile = True
        else:
            profile_dir = custom_profile_dir

        try:
            # Format URI for LibreOffice UserInstallation
            profile_uri = profile_dir.as_uri()

            # Execute headless conversion into the target's parent directory
            target_dir = output_pdf_path.parent
            cmd = [
                self.soffice_binary,
                f"-env:UserInstallation={profile_uri}",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(target_dir),
                str(excel_path),
            ]

            process = subprocess.run(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=60,  # 60s safety timeout
            )

            if process.returncode != 0:
                err_msg = process.stderr.strip() or process.stdout.strip() or f"Exit code {process.returncode}"
                return False, f"LibreOffice conversion failed: {err_msg}"

            # LibreOffice names the output file as {stem}.pdf in target_dir
            default_generated_pdf = target_dir / f"{excel_path.stem}.pdf"
            if not default_generated_pdf.exists():
                return False, "Generated PDF was not found in destination directory."

            # If the user requested a specific filename different from {stem}.pdf, rename it
            if default_generated_pdf != output_pdf_path:
                if output_pdf_path.exists():
                    output_pdf_path.unlink()
                default_generated_pdf.rename(output_pdf_path)

            return True, None

        except subprocess.TimeoutExpired:
            return False, "LibreOffice conversion timed out after 60 seconds."
        except Exception as e:
            return False, f"Unexpected error during LibreOffice conversion: {str(e)}"
        finally:
            if created_temp_profile and profile_dir.exists():
                try:
                    shutil.rmtree(profile_dir, ignore_errors=True)
                except Exception:
                    pass


# --- Multiprocessing Worker Entry Point ---

def run_libreoffice_worker_batch(
    file_pairs: list[tuple[str, str]],
    progress_queue: Any = None,
) -> list[tuple[str, bool, str | None]]:
    """
    Worker process target for LibreOffice:
    Reuses a single isolated profile across multiple files in the same worker batch.
    Reports progress per file via progress_queue if provided.
    """
    results: list[tuple[str, bool, str | None]] = []
    converter = LibreOfficeConverter()
    if not converter.is_available():
        for src, _ in file_pairs:
            err = "LibreOffice is not installed."
            results.append((src, False, err))
            if progress_queue is not None:
                try:
                    progress_queue.put((src, False, err))
                except Exception:
                    pass
        return results

    profile_dir = Path(tempfile.mkdtemp(prefix="soffice_worker_prof_"))
    try:
        for src_path_str, dst_path_str in file_pairs:
            src = Path(src_path_str)
            dst = Path(dst_path_str)
            success, err = converter.convert_file(src, dst, custom_profile_dir=profile_dir)
            results.append((src_path_str, success, err))
            if progress_queue is not None:
                try:
                    progress_queue.put((src_path_str, success, err))
                except Exception:
                    pass
    finally:
        if profile_dir.exists():
            shutil.rmtree(profile_dir, ignore_errors=True)

    return results
