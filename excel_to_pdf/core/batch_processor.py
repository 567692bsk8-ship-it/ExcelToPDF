"""
Parallel batch processor for Excel to PDF conversion.
Distributes workloads across CPU cores and reports real-time progress.
"""

from __future__ import annotations

import concurrent.futures
import math
import multiprocessing
import os
import queue
import sys
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from excel_to_pdf.core.engine_detector import EngineDetector
from excel_to_pdf.core.libreoffice import run_libreoffice_worker_batch
from excel_to_pdf.core.windows_excel import run_windows_worker_batch


@dataclass
class ConversionItem:
    source_path: Path
    output_path: Path
    status: str = "pending"  # "pending", "success", "failed", "skipped"
    error_message: str | None = None


@dataclass
class BatchResult:
    total_files: int = 0
    success_count: int = 0
    failed_count: int = 0
    skipped_count: int = 0
    elapsed_seconds: float = 0.0
    engine_used: str = ""
    items: list[ConversionItem] = field(default_factory=list)
    cancelled: bool = False


# Type definition for progress callback:
# (completed_count, total_count, current_filename, status_text)
ProgressCallback = Callable[[int, int, str, str], None]


class BatchProcessor:
    """Manages file scanning, path resolution, and multi-process parallel execution."""

    def __init__(
        self,
        workers: int = 4,
        engine_type: str | None = None,
        overwrite: bool = False,
        supported_extensions: list[str] | None = None,
    ) -> None:
        self.workers = max(1, workers)
        self.overwrite = overwrite
        self.supported_extensions = [
            ext.lower() for ext in (supported_extensions or [".xlsx", ".xls", ".xlsm", ".xlsb"])
        ]

        if engine_type is None:
            self.engine_type = EngineDetector.get_best_engine() or "none"
        else:
            self.engine_type = engine_type

        self._cancel_requested = False

    def cancel(self) -> None:
        """Signal to cancel ongoing batch execution."""
        self._cancel_requested = True

    def scan_files(
        self,
        input_dir: Path | str,
        recursive: bool = True,
    ) -> list[Path]:
        """Scan input directory for valid, non-temporary Excel files."""
        input_path = Path(input_dir).resolve()
        if not input_path.exists() or not input_path.is_dir():
            return []

        matched_files: list[Path] = []
        pattern = "**/*" if recursive else "*"

        for p in input_path.glob(pattern):
            if not p.is_file():
                continue
            # Skip temporary/lock files created by Excel or OS
            if p.name.startswith("~$") or p.name.startswith("."):
                continue
            if p.suffix.lower() in self.supported_extensions:
                matched_files.append(p)

        return sorted(matched_files)

    def prepare_items(
        self,
        files: list[Path],
        input_dir: Path | str,
        output_dir: Path | str | None = None,
    ) -> list[ConversionItem]:
        """Prepare conversion source and destination paths."""
        input_base = Path(input_dir).resolve()
        custom_out = Path(output_dir).resolve() if output_dir else None

        items: list[ConversionItem] = []
        for src in files:
            if custom_out is None:
                # Output alongside source file
                dest = src.with_suffix(".pdf")
            else:
                # Maintain relative subdirectory structure inside custom output folder
                try:
                    rel_path = src.relative_to(input_base)
                    dest = (custom_out / rel_path).with_suffix(".pdf")
                except ValueError:
                    dest = (custom_out / src.name).with_suffix(".pdf")

            items.append(ConversionItem(source_path=src, output_path=dest))

        return items

    def process_batch(
        self,
        items: list[ConversionItem],
        progress_cb: ProgressCallback | None = None,
    ) -> BatchResult:
        """
        Execute parallel conversion on prepared items.
        Groups items into batches per worker to reuse engine instances.
        """
        self._cancel_requested = False
        start_time = time.time()
        result = BatchResult(
            total_files=len(items),
            engine_used=self.engine_type,
            items=items,
        )

        if not items:
            result.elapsed_seconds = round(time.time() - start_time, 2)
            return result

        if self.engine_type == "none":
            for it in items:
                it.status = "failed"
                it.error_message = "No available conversion engine found (neither MS Excel nor LibreOffice)."
            result.failed_count = len(items)
            result.elapsed_seconds = round(time.time() - start_time, 2)
            return result

        # Separate items into items to convert vs items already existing (if overwrite=False)
        to_convert: list[ConversionItem] = []
        for item in items:
            if not self.overwrite and item.output_path.exists():
                item.status = "skipped"
                item.error_message = "File already exists (overwrite=False)"
                result.skipped_count += 1
            else:
                to_convert.append(item)

        if not to_convert:
            result.elapsed_seconds = round(time.time() - start_time, 2)
            return result

        # Determine batch chunks for worker pool
        num_workers = min(self.workers, len(to_convert))
        chunk_size = math.ceil(len(to_convert) / num_workers)
        chunks: list[list[tuple[str, str]]] = []

        item_lookup: dict[str, ConversionItem] = {
            str(it.source_path.resolve()): it for it in to_convert
        }

        for i in range(0, len(to_convert), chunk_size):
            chunk_items = to_convert[i : i + chunk_size]
            chunks.append(
                [
                    (str(it.source_path.resolve()), str(it.output_path.resolve()))
                    for it in chunk_items
                ]
            )

        worker_fn = (
            run_windows_worker_batch
            if self.engine_type == "windows_excel"
            else run_libreoffice_worker_batch
        )

        completed_count = result.skipped_count
        total_target = len(to_convert) + result.skipped_count

        manager = multiprocessing.Manager()
        progress_q = manager.Queue()

        try:
            with concurrent.futures.ProcessPoolExecutor(max_workers=num_workers) as executor:
                futures = [
                    executor.submit(worker_fn, chunk, progress_q) for chunk in chunks
                ]

                # Stream and report progress in real time as each file completes
                while completed_count < total_target:
                    if self._cancel_requested:
                        result.cancelled = True
                        break

                    try:
                        src_str, success, err = progress_q.get(timeout=0.03)
                        item = item_lookup.get(src_str)
                        if item and item.status == "pending":
                            if success:
                                item.status = "success"
                                item.error_message = None
                                result.success_count += 1
                            else:
                                item.status = "failed"
                                item.error_message = err
                                result.failed_count += 1

                            completed_count += 1
                            if progress_cb:
                                status_msg = "成功" if success else f"失敗: {err}"
                                progress_cb(
                                    completed_count,
                                    result.total_files,
                                    item.source_path.name,
                                    status_msg,
                                )
                    except queue.Empty:
                        if all(f.done() for f in futures) and progress_q.empty():
                            break

                # Wait for all workers to finish and check results
                for f in futures:
                    try:
                        chunk_res = f.result(timeout=5)
                        for src_str, success, err in chunk_res:
                            item = item_lookup.get(src_str)
                            if item and item.status == "pending":
                                if success:
                                    item.status = "success"
                                    result.success_count += 1
                                else:
                                    item.status = "failed"
                                    item.error_message = err
                                    result.failed_count += 1
                                completed_count += 1
                                if progress_cb:
                                    status_msg = "成功" if success else f"失敗: {err}"
                                    progress_cb(
                                        completed_count,
                                        result.total_files,
                                        item.source_path.name,
                                        status_msg,
                                    )
                    except Exception as e:
                        for it in to_convert:
                            if it.status == "pending":
                                it.status = "failed"
                                it.error_message = f"Worker error: {str(e)}"
                                result.failed_count += 1
                                completed_count += 1
                                if progress_cb:
                                    progress_cb(
                                        completed_count,
                                        result.total_files,
                                        it.source_path.name,
                                        "ワーカー異常終了",
                                    )
        finally:
            try:
                manager.shutdown()
            except Exception:
                pass

        # Check remaining pending items if cancelled
        for it in items:
            if it.status == "pending":
                it.status = "skipped"
                it.error_message = "ユーザーにより中断されました"
                result.skipped_count += 1

        result.elapsed_seconds = round(time.time() - start_time, 2)
        return result
