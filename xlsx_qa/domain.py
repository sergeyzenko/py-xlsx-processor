"""Domain models for the questionnaire processor."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional


BOOLEAN_TRUE = {"true", "yes", "y", "1"}
BOOLEAN_FALSE = {"false", "no", "n", "0"}


@dataclass(slots=True)
class CatalogRecord:
    """Representation of a single catalog entry sourced from the workbook."""

    tab_name: str
    text_location: str
    text_value: str
    is_question: Optional[bool] = None
    text_response: Optional[str] = None
    text_response_location: Optional[str] = None
    is_default_answer: Optional[bool] = None
    default_answer_question_location: Optional[str] = None
    is_instruction: Optional[bool] = None

    def as_csv_row(self) -> List[str]:
        return [
            self.tab_name,
            self.text_location,
            self.text_value,
            _bool_to_field(self.is_question),
            self.text_response or "",
            self.text_response_location or "",
            _bool_to_field(self.is_default_answer),
            self.default_answer_question_location or "",
            _bool_to_field(self.is_instruction),
        ]

    def update_from_existing(self, other: "CatalogRecord") -> None:
        self.is_question = other.is_question
        self.text_response = other.text_response
        self.text_response_location = other.text_response_location
        self.is_default_answer = other.is_default_answer
        self.default_answer_question_location = other.default_answer_question_location
        self.is_instruction = other.is_instruction


def _bool_to_field(value: Optional[bool]) -> str:
    if value is None:
        return ""
    return "true" if value else "false"


def parse_bool(value: str | None) -> Optional[bool]:
    if value is None:
        return None
    lowered = value.strip().lower()
    if lowered in BOOLEAN_TRUE:
        return True
    if lowered in BOOLEAN_FALSE:
        return False
    return None


def build_catalog_record(row: Dict[str, str]) -> CatalogRecord:
    return CatalogRecord(
        tab_name=row["tabName"],
        text_location=row["textLocation"],
        text_value=row["textValue"],
        is_question=parse_bool(row.get("isQuestion")),
        text_response=row.get("textResponse") or None,
        text_response_location=row.get("textResponseLocation") or None,
        is_default_answer=parse_bool(row.get("isDefaultAnswer")),
        default_answer_question_location=row.get("defaultAnswerQuestionLocation") or None,
        is_instruction=parse_bool(row.get("isInstruction")),
    )


CATALOG_HEADERS: List[str] = [
    "tabName",
    "textLocation",
    "textValue",
    "isQuestion",
    "textResponse",
    "textResponseLocation",
    "isDefaultAnswer",
    "defaultAnswerQuestionLocation",
    "isInstruction",
]


def catalog_key(record: CatalogRecord) -> tuple[str, str]:
    return (record.tab_name, record.text_location)


def order_catalog(records: Iterable[CatalogRecord]) -> List[CatalogRecord]:
    return list(records)
