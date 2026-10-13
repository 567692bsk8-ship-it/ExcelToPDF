"""
Base abstract class for Excel to PDF converters.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from pathlib import Path


class BaseConverter(ABC):
    """Abstract base class for conversion engines."""

    @classmethod
    @abstractmethod
    def is_available(cls) -> bool:
        """Check if this converter is available on the current system."""
        pass

    @classmethod
    @abstractmethod
    def engine_name(cls) -> str:
        """Return the user-friendly name of this conversion engine."""
        pass

    @abstractmethod
    def convert_file(self, excel_path: Path, output_pdf_path: Path) -> tuple[bool, str | None]:
        """
        Convert a single Excel file to PDF.
        Returns: (success: bool, error_message: str | None)
        """
        pass
