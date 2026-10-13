"""
Conversion history management with complete CRUD capabilities.
"""

from __future__ import annotations

import json
import os
import uuid
from dataclasses import asdict, dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Any


@dataclass
class HistoryRecord:
    id: str = field(default_factory=lambda: "hist_" + str(uuid.uuid4())[:8])
    timestamp: str = field(default_factory=lambda: datetime.now().isoformat())
    source_folder: str = ""
    output_folder: str = ""
    total_files: int = 0
    success_count: int = 0
    failed_count: int = 0
    elapsed_seconds: float = 0.0
    workers_used: int = 4
    engine_used: str = "Unknown"
    memo: str = ""
    failed_files: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> HistoryRecord:
        return cls(
            id=data.get("id", "hist_" + str(uuid.uuid4())[:8]),
            timestamp=data.get("timestamp", datetime.now().isoformat()),
            source_folder=data.get("source_folder", ""),
            output_folder=data.get("output_folder", ""),
            total_files=int(data.get("total_files", 0)),
            success_count=int(data.get("success_count", 0)),
            failed_count=int(data.get("failed_count", 0)),
            elapsed_seconds=float(data.get("elapsed_seconds", 0.0)),
            workers_used=int(data.get("workers_used", 4)),
            engine_used=data.get("engine_used", "Unknown"),
            memo=data.get("memo", ""),
            failed_files=list(data.get("failed_files", [])),
        )


class HistoryManager:
    """Manages CRUD operations for conversion execution history."""

    def __init__(self, storage_path: str | Path | None = None) -> None:
        if storage_path is None:
            base_dir = Path(__file__).resolve().parent.parent.parent / "data"
            base_dir.mkdir(parents=True, exist_ok=True)
            self.storage_file = base_dir / "history.json"
        else:
            self.storage_file = Path(storage_path)
            self.storage_file.parent.mkdir(parents=True, exist_ok=True)

        self._ensure_initialized()

    def _ensure_initialized(self) -> None:
        if not self.storage_file.exists():
            self._save_all([])

    def _load_all(self) -> list[HistoryRecord]:
        try:
            with open(self.storage_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                records_data = data.get("records", [])
                return [HistoryRecord.from_dict(item) for item in records_data]
        except (json.JSONDecodeError, OSError):
            return []

    def _save_all(self, records: list[HistoryRecord]) -> None:
        temp_file = self.storage_file.with_suffix(".tmp")
        payload = {"records": [r.to_dict() for r in records]}
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        os.replace(temp_file, self.storage_file)

    # --- CRUD Operations ---

    # 1. Create
    def add_record(self, record: HistoryRecord) -> HistoryRecord:
        """Create and store a new history record."""
        records = self._load_all()
        # Add to beginning of list (newest first)
        records.insert(0, record)
        # Keep maximum 100 recent records to prevent file bloat
        records = records[:100]
        self._save_all(records)
        return record

    # 2. Read
    def get_all_records(self) -> list[HistoryRecord]:
        """Read all history records, newest first."""
        return self._load_all()

    def get_record_by_id(self, record_id: str) -> HistoryRecord | None:
        """Read a single record by ID."""
        records = self._load_all()
        for r in records:
            if r.id == record_id:
                return r
        return None

    # 3. Update
    def update_record_memo(self, record_id: str, new_memo: str) -> bool:
        """Update the memo of an existing record."""
        records = self._load_all()
        found = False
        for r in records:
            if r.id == record_id:
                r.memo = new_memo
                found = True
                break
        if found:
            self._save_all(records)
        return found

    # 4. Delete
    def delete_record(self, record_id: str) -> bool:
        """Delete a single history record by ID."""
        records = self._load_all()
        new_records = [r for r in records if r.id != record_id]
        if len(new_records) == len(records):
            return False
        self._save_all(new_records)
        return True

    def clear_all_records(self) -> None:
        """Delete all history records."""
        self._save_all([])
