"""Application orchestration."""

from __future__ import annotations

import os
from typing import Iterable, List

from openpyxl.utils.exceptions import InvalidFileException

from .cli import AppConfig
from .domain import Question
from .errors import NoQuestionsFoundError, ResumeFileMismatchError
from .extractor import QuestionExtractor
from .loader import WorkbookLoader
from .persistence import ProgressRepository
from .session import InteractiveSession, SessionResult
from .writer import AnswerWriter


class XLSXQuestionnaireApp:
    """High-level coordinator for the questionnaire workflow."""

    def __init__(
        self,
        loader: WorkbookLoader | None = None,
        extractor: QuestionExtractor | None = None,
        session: InteractiveSession | None = None,
        repository: ProgressRepository | None = None,
        writer: AnswerWriter | None = None,
    ) -> None:
        self._loader = loader or WorkbookLoader()
        self._extractor = extractor or QuestionExtractor()
        self._session = session or InteractiveSession()
        self._repository = repository or ProgressRepository()
        self._writer = writer or AnswerWriter(self._loader)

    def run(self, config: AppConfig) -> int:
        try:
            workbook = self._loader.load(config.source_path)
        except FileNotFoundError as error:
            print(error)
            return 1
        except InvalidFileException as error:
            print(f"Unable to open workbook: {error}")
            return 1

        questions = self._extractor.extract(workbook)
        if not questions:
            raise NoQuestionsFoundError("No questions found in workbook.")

        start_index = 0
        if config.resume_path:
            try:
                questions, start_index = self._apply_resume(
                    config.resume_path, config.source_path, questions
                )
            except ResumeFileMismatchError as error:
                print(error)
                return 1

        print(f"Found {len(questions)} questions...")
        result = self._session.run(questions, start_index=start_index)

        if result.aborted:
            self._persist_progress(config, result, config.source_path)
            print("Session aborted by user. Progress saved.")
            return 0

        output_path = self._writer.write_answers(config.source_path, result.questions, config.output_path)
        print(f"Answers written to: {output_path}")
        return 0

    def _apply_resume(
        self, resume_path: str, source_path: str, questions: List[Question]
    ) -> tuple[List[Question], int]:
        snapshot = self._repository.load(resume_path)
        source_basename = os.path.basename(source_path)
        snapshot_basename = os.path.basename(snapshot.source_file)

        if source_basename != snapshot_basename:
            raise ResumeFileMismatchError(
                f"Resume file created for {snapshot.source_file}, not {source_basename}."
            )

        question_map = {(q.sheet, q.coord): q for q in questions}
        for stored in snapshot.questions:
            key = (stored.sheet, stored.coord)
            if key in question_map:
                question_map[key] = question_map[key].with_answer(stored.answer)

        ordered_questions = [question_map[(q.sheet, q.coord)] for q in questions]
        start_index = min(snapshot.current_index, len(ordered_questions))
        return ordered_questions, start_index

    def _persist_progress(self, config: AppConfig, result: SessionResult, source_path: str) -> None:
        resume_path = config.resume_path or _default_progress_path(source_path)
        self._repository.save(resume_path, result.questions, result.current_index, os.path.basename(source_path))
        print(f"Progress saved to {resume_path}")


def _default_progress_path(source_path: str) -> str:
    root, _ = os.path.splitext(source_path)
    return f"{root}_progress.json"
