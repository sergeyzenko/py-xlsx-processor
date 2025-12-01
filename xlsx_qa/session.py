"""Interactive questionnaire session handling."""

from __future__ import annotations

import shutil
import textwrap
from dataclasses import dataclass
from typing import Dict, Iterable, List, Optional, Protocol

from .domain import CatalogRecord


class IO(Protocol):
    """Minimal IO boundary to allow console replacement in tests."""

    def write(self, message: str = "") -> None:
        ...

    def read(self, prompt: str = "") -> str:
        ...


class ConsoleIO:
    """Default IO implementation that proxies to built-in input/print."""

    def write(self, message: str = "") -> None:
        print(message)

    def read(self, prompt: str = "") -> str:
        return input(prompt)


@dataclass
class SessionResult:
    """Result of a Q&A session."""

    records: List[CatalogRecord]
    current_index: int
    aborted: bool = False


class InteractiveSession:
    """Runs the questionnaire loop and tracks user progress."""

    def __init__(self, io: IO | None = None, wrap_width: int | None = None) -> None:
        self._io = io or ConsoleIO()
        self._wrap_width = wrap_width or _detect_terminal_width()

    def run(
        self,
        question_records: Iterable[CatalogRecord],
        default_answers: Dict[str, str],
        start_index: int = 0,
    ) -> SessionResult:
        records = list(question_records)
        index = max(0, start_index)

        while 0 <= index < len(records):
            record = records[index]
            self._display_header(index + 1, len(records), record)

            default_answer = record.text_response or default_answers.get(record.text_location)
            prompt_result = self._prompt_answer(default_answer)

            if prompt_result.command == "quit":
                return SessionResult(records, index, aborted=True)

            if prompt_result.command == "back":
                index = max(0, index - 1)
                continue

            if prompt_result.command == "skip":
                index += 1
                continue

            answer = prompt_result.answer
            location = self._prompt_location(
                existing=record.text_response_location,
                question_location=record.text_location,
            )

            if location.command == "quit":
                return SessionResult(records, index, aborted=True)

            if location.command == "back":
                index = max(0, index - 1)
                continue

            record.text_response = answer
            record.text_response_location = location.value
            index += 1

        return SessionResult(records, len(records), aborted=False)

    def _display_header(self, number: int, total: int, record: CatalogRecord) -> None:
        divider = "=" * min(self._wrap_width, 70)
        header = f"Question {number}/{total}  |  Sheet: {record.tab_name}  |  Cell: {record.text_location}"
        wrapped_text = textwrap.fill(record.text_value, width=self._wrap_width)
        self._io.write(divider)
        self._io.write(header)
        self._io.write(divider)
        self._io.write("")
        self._io.write(wrapped_text)
        self._io.write("")

    def _prompt_answer(self, default_answer: Optional[str]) -> "AnswerPromptResult":
        if default_answer:
            self._io.write("Proposed default answer:")
            self._io.write(default_answer)
            self._io.write("")

        self._io.write(
            "Answer (enter to accept default/leave blank, 'q' to quit, 'b' to go back, 's' to skip):"
        )
        self._io.write("Use \\n for explicit line breaks or press enter three times to finish multi-line input.")
        lines: List[str] = []
        empty_streak = 0

        while True:
            raw_input = self._io.read("> " if not lines else "... ")
            command = raw_input.strip().lower()

            if not lines and command in {"q", "quit"}:
                return AnswerPromptResult("quit", None)
            if not lines and command in {"b", "back"}:
                return AnswerPromptResult("back", None)
            if not lines and command in {"s", "skip"}:
                return AnswerPromptResult("skip", None)

            if not lines and raw_input == "":
                if default_answer:
                    return AnswerPromptResult("answer", default_answer)
                return AnswerPromptResult("skip", None)

            if raw_input == "":
                empty_streak += 1
                if empty_streak >= 3:
                    break
                lines.append("")
                continue

            empty_streak = 0
            lines.append(raw_input.replace("\\n", "\n"))

        answer = "\n".join(lines).strip()
        return AnswerPromptResult("answer", answer if answer else default_answer)

    def _prompt_location(
        self, existing: Optional[str], question_location: str
    ) -> "LocationPromptResult":
        suggestion = existing or _suggest_response_location(question_location)
        prompt = f"Response location [{suggestion}]: " if suggestion else "Response location: "

        while True:
            raw_input = self._io.read(prompt).strip()
            lowered = raw_input.lower()

            if lowered in {"q", "quit"}:
                return LocationPromptResult("quit", None)
            if lowered in {"b", "back"}:
                return LocationPromptResult("back", None)

            if not raw_input and suggestion:
                return LocationPromptResult("accept", suggestion)

            if raw_input:
                return LocationPromptResult("accept", raw_input)

            self._io.write("A response location is required. Enter 'b' to go back if needed.")


@dataclass
class AnswerPromptResult:
    command: str
    answer: Optional[str]


@dataclass
class LocationPromptResult:
    command: str
    value: Optional[str]


def _detect_terminal_width(default: int = 80) -> int:
    try:
        columns = shutil.get_terminal_size().columns
    except OSError:  # pragma: no cover - depends on runtime
        return default
    return max(40, min(columns, 120))


def _suggest_response_location(question_location: str) -> str:
    column_part = ""
    row_part = ""
    for char in question_location:
        if char.isalpha():
            column_part += char
        elif char.isdigit():
            row_part += char

    if not column_part or not row_part:
        return ""

    column_index = 0
    for char in column_part.upper():
        column_index = column_index * 26 + (ord(char) - ord("A") + 1)

    column_index += 1

    letters: List[str] = []
    while column_index > 0:
        column_index, remainder = divmod(column_index - 1, 26)
        letters.append(chr(ord("A") + remainder))

    return f"{''.join(reversed(letters))}{row_part}"
