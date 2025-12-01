"""Domain models for the questionnaire processor."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from typing import Any, Dict, List, Optional


@dataclass(slots=True)
class Question:
    """Immutable representation of a question discovered in the workbook."""

    id: int
    sheet: str
    coord: str
    row: int
    col: int
    text: str
    answer_coord: str
    answer: Optional[str] = None
    preexisting_answer: Optional[str] = None

    def with_answer(self, answer: Optional[str]) -> "Question":
        """Return a new Question instance with the updated answer."""
        return Question(
            id=self.id,
            sheet=self.sheet,
            coord=self.coord,
            row=self.row,
            col=self.col,
            text=self.text,
            answer_coord=self.answer_coord,
            answer=answer,
            preexisting_answer=self.preexisting_answer,
        )


@dataclass(slots=True)
class ProgressSnapshot:
    """Serializable representation of session progress."""

    source_file: str
    timestamp: datetime
    current_index: int
    questions: List[Question] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        """Convert the snapshot into a JSON-serializable dictionary."""
        return {
            "source_file": self.source_file,
            "timestamp": self.timestamp.isoformat(timespec="seconds"),
            "current_index": self.current_index,
            "questions": [
                {
                    "id": q.id,
                    "sheet": q.sheet,
                    "coord": q.coord,
                    "row": q.row,
                    "col": q.col,
                    "text": q.text,
                    "answer_coord": q.answer_coord,
                    "answer": q.answer,
                }
                for q in self.questions
            ],
        }

    @staticmethod
    def from_dict(payload: Dict[str, Any]) -> "ProgressSnapshot":
        """Hydrate a snapshot from a dictionary."""
        timestamp_value = payload.get("timestamp")
        timestamp = datetime.fromisoformat(timestamp_value) if timestamp_value else datetime.utcnow()

        questions: List[Question] = []
        for entry in payload.get("questions", []):
            questions.append(
                Question(
                    id=int(entry["id"]),
                    sheet=str(entry["sheet"]),
                    coord=str(entry["coord"]),
                    row=int(entry["row"]),
                    col=int(entry["col"]),
                    text=str(entry["text"]),
                    answer_coord=str(entry["answer_coord"]),
                    answer=entry.get("answer"),
                )
            )

        return ProgressSnapshot(
            source_file=str(payload["source_file"]),
            timestamp=timestamp,
            current_index=int(payload.get("current_index", 0)),
            questions=questions,
        )
