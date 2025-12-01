"""Application orchestration."""

from __future__ import annotations

import os
from typing import Dict, Iterable, List

from openpyxl.utils.exceptions import InvalidFileException

from .cli import AppConfig
from .domain import CatalogRecord, catalog_key
from .errors import NoQuestionsFoundError
from .extractor import WorkbookCatalogExtractor
from .loader import WorkbookLoader
from .persistence import CatalogRepository
from .session import InteractiveSession
from .writer import AnswerWriter


class XLSXQuestionnaireApp:
    """High-level coordinator for the questionnaire workflow."""

    def __init__(
        self,
        loader: WorkbookLoader | None = None,
        extractor: WorkbookCatalogExtractor | None = None,
        session: InteractiveSession | None = None,
        repository: CatalogRepository | None = None,
        writer: AnswerWriter | None = None,
    ) -> None:
        self._loader = loader or WorkbookLoader()
        self._extractor = extractor or WorkbookCatalogExtractor()
        self._session = session or InteractiveSession()
        self._repository = repository or CatalogRepository()
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

        base_records = self._extractor.extract(workbook)
        if not base_records:
            raise NoQuestionsFoundError("No textual entries found in workbook.")

        catalog_path = config.catalog_path or _default_catalog_path(config.source_path)
        existing_records = self._repository.load(catalog_path)
        merged_records = self._merge_records(base_records, existing_records)

        initial_catalog_path = self._repository.save(catalog_path, merged_records)
        print(f"Catalog saved to: {initial_catalog_path}")

        question_candidates = [
            record
            for record in merged_records
            if record.is_question is True and not record.text_response
        ]
        default_answers = self._build_default_answer_map(merged_records)

        if not question_candidates:
            answered_records = [
                record
                for record in merged_records
                if record.text_response and record.text_response_location
            ]
            if answered_records:
                output_path = self._writer.write_answers(
                    config.source_path, merged_records, config.output_path
                )
                print(f"Answers written to: {output_path}")
            else:
                print("No unanswered questions flagged in catalog. Update the CSV and rerun when ready.")
            return 0

        print(f"Found {len(question_candidates)} unanswered questions marked in catalog...")
        result = self._session.run(question_candidates, default_answers)

        updated_catalog_path = self._repository.save(catalog_path, merged_records)

        if result.aborted:
            print(f"Session aborted by user. Catalog saved to {updated_catalog_path}.")
            return 0

        print(f"Catalog saved to: {updated_catalog_path}")
        output_path = self._writer.write_answers(config.source_path, merged_records, config.output_path)
        print(f"Answers written to: {output_path}")
        return 0

    def _merge_records(
        self,
        base_records: Iterable[CatalogRecord],
        existing_records: Dict[tuple[str, str], CatalogRecord],
    ) -> List[CatalogRecord]:
        merged: Dict[tuple[str, str], CatalogRecord] = {}
        ordered_records: List[CatalogRecord] = []

        for record in base_records:
            key = catalog_key(record)
            if key in existing_records:
                record.update_from_existing(existing_records[key])
            merged[key] = record
            ordered_records.append(record)

        for key, existing in existing_records.items():
            if key in merged:
                continue
            merged[key] = existing
            ordered_records.append(existing)

        return ordered_records

    def _build_default_answer_map(self, records: Iterable[CatalogRecord]) -> Dict[str, str]:
        lookup: Dict[str, str] = {}
        for record in records:
            if record.is_default_answer and record.default_answer_question_location and record.text_value:
                lookup[record.default_answer_question_location] = record.text_value
        return lookup


def _default_catalog_path(source_path: str) -> str:
    root, _ = os.path.splitext(source_path)
    return f"{root}_catalog.csv"
