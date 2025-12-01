"""Workbook catalog extraction logic."""

from __future__ import annotations

from typing import Dict, List

from openpyxl.cell.cell import Cell, MergedCell
from openpyxl.utils import get_column_letter
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .domain import CatalogRecord


class WorkbookCatalogExtractor:
    """Scan a workbook and capture all textual entries."""

    def extract(self, workbook: Workbook) -> List[CatalogRecord]:
        records: List[CatalogRecord] = []

        for sheet in workbook.worksheets:
            merged_lookup = _build_merged_lookup(sheet)

            for row in sheet.iter_rows():
                for cell in row:
                    if isinstance(cell, MergedCell):
                        continue

                    if cell.coordinate in merged_lookup.skip_coordinates:
                        continue

                    text_value = _coerce_text(cell.value)
                    if text_value is None:
                        continue

                    records.append(
                        CatalogRecord(
                            tab_name=sheet.title,
                            text_location=cell.coordinate,
                            text_value=text_value,
                        )
                    )

        return records


def _coerce_text(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        return value
    return None


class _MergedLookup:
    __slots__ = ("rightmost_column", "skip_coordinates")

    def __init__(self, rightmost_column: Dict[str, int], skip_coordinates: set[str]) -> None:
        self.rightmost_column = rightmost_column
        self.skip_coordinates = skip_coordinates


def _build_merged_lookup(sheet: Worksheet) -> _MergedLookup:
    rightmost_column: Dict[str, int] = {}
    skip_coordinates: set[str] = set()

    for merged_range in sheet.merged_cells.ranges:
        min_row, max_row = merged_range.min_row, merged_range.max_row
        min_col, max_col = merged_range.min_col, merged_range.max_col
        top_left = sheet.cell(row=min_row, column=min_col)
        rightmost_column[top_left.coordinate] = max_col

        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                coord = f"{get_column_letter(col)}{row}"
                if coord != top_left.coordinate:
                    skip_coordinates.add(coord)

    return _MergedLookup(rightmost_column, skip_coordinates)
