"""Answer writing services."""

from __future__ import annotations

import os
from typing import Iterable

from openpyxl.workbook import Workbook

from .domain import Question
from .loader import WorkbookLoader


class AnswerWriter:
    """Write collected answers back to a workbook and persist to disk."""

    def __init__(self, loader: WorkbookLoader | None = None) -> None:
        self._loader = loader or WorkbookLoader(data_only=False, read_only=False)

    def write_answers(self, filepath: str, questions: Iterable[Question], output_path: str | None = None) -> str:
        workbook = self._loader.reload(filepath)
        self._apply_answers(workbook, questions)

        target_path = output_path or _default_output_path(filepath)
        workbook.save(target_path)
        return target_path

    def _apply_answers(self, workbook: Workbook, questions: Iterable[Question]) -> None:
        for question in questions:
            if not question.answer:
                continue
            sheet = workbook[question.sheet]
            sheet[question.answer_coord].value = question.answer


def _default_output_path(source_path: str) -> str:
    root, ext = os.path.splitext(source_path)
    suffix = "_answered"
    return f"{root}{suffix}{ext or '.xlsx'}"
