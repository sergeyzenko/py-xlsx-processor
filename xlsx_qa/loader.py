"""Workbook loading services."""

from __future__ import annotations

import os
from typing import Optional

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook import Workbook


class WorkbookLoader:
    """Load workbooks from disk with basic validation."""

    def __init__(self, data_only: bool = True, read_only: bool = False) -> None:
        self._data_only = data_only
        self._read_only = read_only

    def load(self, filepath: str) -> Workbook:
        """Load an XLSX workbook or raise if the file is not accessible."""
        if not os.path.isfile(filepath):
            raise FileNotFoundError(f"Workbook not found: {filepath}")

        try:
            return load_workbook(filename=filepath, data_only=self._data_only, read_only=self._read_only)
        except InvalidFileException:
            raise
        except Exception as exc:  # pragma: no cover - defensive path
            raise InvalidFileException(str(exc)) from exc

    def reload(self, filepath: str) -> Workbook:
        """Convenience method kept for symmetry with higher-level workflows."""
        return self.load(filepath)
