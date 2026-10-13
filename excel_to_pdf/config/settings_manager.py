"""
Conversion profiles management with complete CRUD capabilities.
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
class ConversionProfile:
    id: str = field(default_factory=lambda: str(uuid.uuid4())[:8])
    name: str = "標準プロファイル"
    description: str = "標準的な高速並列変換設定"
    workers: int = 4
    recursive: bool = False
    output_mode: str = "pdf_subfolder"  # "pdf_subfolder", "same_directory", or "custom_directory"
    custom_output_dir: str = ""
    overwrite: bool = False
    supported_extensions: list[str] = field(
        default_factory=lambda: [".xlsx", ".xls", ".xlsm", ".xlsb"]
    )
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.now().isoformat())

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> ConversionProfile:
        return cls(
            id=data.get("id", str(uuid.uuid4())[:8]),
            name=data.get("name", "名称未設定"),
            description=data.get("description", ""),
            workers=int(data.get("workers", 4)),
            recursive=bool(data.get("recursive", False)),
            output_mode=data.get("output_mode", "pdf_subfolder"),
            custom_output_dir=data.get("custom_output_dir", ""),
            overwrite=bool(data.get("overwrite", False)),
            supported_extensions=list(
                data.get("supported_extensions", [".xlsx", ".xls", ".xlsm", ".xlsb"])
            ),
            created_at=data.get("created_at", datetime.now().isoformat()),
            updated_at=data.get("updated_at", datetime.now().isoformat()),
        )


class SettingsManager:
    """Manages CRUD operations for conversion profiles."""

    DEFAULT_PROFILES = [
        ConversionProfile(
            id="default_standard",
            name="標準（バランス型）",
            description="並列4プロセス。指定フォルダ直下のExcelのみを「PDF」フォルダに一括出力。",
            workers=4,
            recursive=False,
            output_mode="pdf_subfolder",
            overwrite=False,
        ),
        ConversionProfile(
            id="high_speed_turbo",
            name="超高速ターボ（最大並列）",
            description="並列8プロセスで最速処理。指定フォルダ直下を「PDF」フォルダに高速出力。",
            workers=8,
            recursive=False,
            output_mode="pdf_subfolder",
            overwrite=True,
        ),
        ConversionProfile(
            id="safe_overwrite_off",
            name="安全モード（上書きなし）",
            description="並列2プロセス。既存PDFはスキップし「PDF」フォルダに安全出力。",
            workers=2,
            recursive=False,
            output_mode="pdf_subfolder",
            overwrite=False,
        ),
        ConversionProfile(
            id="recursive_all_subfolders",
            name="サブフォルダ一括探索",
            description="配下のサブフォルダ内も含めて全階層のExcelを再帰的に検出・変換。",
            workers=4,
            recursive=True,
            output_mode="pdf_subfolder",
            overwrite=False,
        ),
    ]

    def __init__(self, storage_path: str | Path | None = None) -> None:
        if storage_path is None:
            base_dir = Path(__file__).resolve().parent.parent.parent / "data"
            base_dir.mkdir(parents=True, exist_ok=True)
            self.storage_file = base_dir / "profiles.json"
        else:
            self.storage_file = Path(storage_path)
            self.storage_file.parent.mkdir(parents=True, exist_ok=True)

        self._ensure_initialized()

    def _ensure_initialized(self) -> None:
        if not self.storage_file.exists():
            self._save_all(self.DEFAULT_PROFILES)

    def _load_all(self) -> list[ConversionProfile]:
        try:
            with open(self.storage_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                profiles_data = data.get("profiles", [])
                return [ConversionProfile.from_dict(item) for item in profiles_data]
        except (json.JSONDecodeError, OSError):
            return list(self.DEFAULT_PROFILES)

    def _save_all(self, profiles: list[ConversionProfile]) -> None:
        temp_file = self.storage_file.with_suffix(".tmp")
        payload = {"profiles": [p.to_dict() for p in profiles]}
        with open(temp_file, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        os.replace(temp_file, self.storage_file)

    # --- CRUD Operations ---

    # 1. Create
    def create_profile(self, profile: ConversionProfile) -> ConversionProfile:
        """Create a new profile."""
        profiles = self._load_all()
        # Ensure unique ID
        if any(p.id == profile.id for p in profiles):
            profile.id = str(uuid.uuid4())[:8]
        profile.created_at = datetime.now().isoformat()
        profile.updated_at = profile.created_at
        profiles.append(profile)
        self._save_all(profiles)
        return profile

    # 2. Read
    def get_all_profiles(self) -> list[ConversionProfile]:
        """Read all profiles."""
        return self._load_all()

    def get_profile_by_id(self, profile_id: str) -> ConversionProfile | None:
        """Read a single profile by ID."""
        profiles = self._load_all()
        for p in profiles:
            if p.id == profile_id:
                return p
        return None

    # 3. Update
    def update_profile(self, profile: ConversionProfile) -> bool:
        """Update an existing profile."""
        profiles = self._load_all()
        found = False
        for idx, p in enumerate(profiles):
            if p.id == profile.id:
                profile.updated_at = datetime.now().isoformat()
                profiles[idx] = profile
                found = True
                break
        if found:
            self._save_all(profiles)
        return found

    # 4. Delete
    def delete_profile(self, profile_id: str) -> bool:
        """Delete a profile by ID."""
        profiles = self._load_all()
        new_profiles = [p for p in profiles if p.id != profile_id]
        if len(new_profiles) == len(profiles):
            return False
        # Do not allow deleting all profiles; keep at least one default
        if not new_profiles:
            new_profiles = [self.DEFAULT_PROFILES[0]]
        self._save_all(new_profiles)
        return True
