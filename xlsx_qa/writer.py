"""Answer writing services."""

from __future__ import annotations

import os
from typing import Iterable

from openpyxl.workbook import Workbook

from .domain import CatalogRecord
from .loader import WorkbookLoader


class AnswerWriter:
    """Write collected answers back to a workbook and persist to disk."""

    def __init__(self, loader: WorkbookLoader | None = None) -> None:
        self._loader = loader or WorkbookLoader(data_only=False, read_only=False)

    def write_answers(
        self,
        filepath: str,
        records: Iterable[CatalogRecord],
        output_path: str | None = None,
    ) -> str:
        workbook = self._loader.reload(filepath)
        self._apply_answers(workbook, records)

        target_path = output_path or _default_output_path(filepath)
        workbook.save(target_path)
        return target_path

    def _apply_answers(self, workbook: Workbook, records: Iterable[CatalogRecord]) -> None:
        for record in records:
            if not record.text_response or not record.text_response_location:
                continue
            if record.tab_name not in workbook.sheetnames:
                continue
            sheet = workbook[record.tab_name]
            sheet[record.text_response_location].value = record.text_response


def _default_output_path(source_path: str) -> str:
    root, ext = os.path.splitext(source_path)
    suffix = "_answered"
    return f"{root}{suffix}{ext or '.xlsx'}"
