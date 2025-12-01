"""Interactive questionnaire session handling."""

from __future__ import annotations

import os
import shutil
import textwrap
from dataclasses import dataclass
from typing import Iterable, List, Optional, Protocol

from .domain import Question


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

    questions: List[Question]
    current_index: int
    aborted: bool = False


class InteractiveSession:
    """Runs the questionnaire loop and tracks user progress."""

    def __init__(self, io: IO | None = None, wrap_width: int | None = None) -> None:
        self._io = io or ConsoleIO()
        self._wrap_width = wrap_width or _detect_terminal_width()

    def run(self, questions: Iterable[Question], start_index: int = 0) -> SessionResult:
        question_list = list(questions)
        index = max(0, start_index)

        while 0 <= index < len(question_list):
            question = question_list[index]
            self._display_header(index + 1, len(question_list), question)

            if question.preexisting_answer and not question.answer:
                if not self._confirm_overwrite(question.preexisting_answer):
                    index += 1
                    continue

            command, answer = self._prompt_answer()

            if command == "quit":
                return SessionResult(question_list, index, aborted=True)

            if command == "back":
                index = max(0, index - 1)
                continue

            if command == "skip":
                index += 1
                continue

            question_list[index] = question.with_answer(answer)
            index += 1

        return SessionResult(question_list, len(question_list), aborted=False)

    def _display_header(self, number: int, total: int, question: Question) -> None:
        divider = "=" * min(self._wrap_width, 70)
        header = f"Question {number}/{total}  |  Sheet: {question.sheet}  |  Cell: {question.coord}"
        wrapped_text = textwrap.fill(question.text, width=self._wrap_width)
        self._io.write(divider)
        self._io.write(header)
        self._io.write(divider)
        self._io.write("")
        self._io.write(wrapped_text)
        self._io.write("")

    def _confirm_overwrite(self, existing_value: str) -> bool:
        self._io.write(f"Existing answer detected: {existing_value}")
        while True:
            response = self._io.read("Overwrite? [y/N]: ").strip().lower()
            if response in {"", "n", "no"}:
                return False
            if response in {"y", "yes"}:
                return True
            self._io.write("Please respond with 'y' or 'n'.")

    def _prompt_answer(self) -> tuple[str, Optional[str]]:
        self._io.write("Answer (enter to skip, 'q' to quit, 'b' to go back, 's' to skip):")
        self._io.write("Use \\n for explicit line breaks or press enter three times to finish multi-line input.")
        lines: List[str] = []
        empty_streak = 0

        while True:
            raw_input = self._io.read("> " if not lines else "... ")
            command = raw_input.strip().lower()

            if not lines and command in {"q", "quit"}:
                return "quit", None
            if not lines and command in {"b", "back"}:
                return "back", None
            if not lines and command in {"s", "skip"}:
                return "skip", None

            if not lines and raw_input == "":
                return "skip", None

            if raw_input == "":
                empty_streak += 1
                if empty_streak >= 3:
                    break
                lines.append("")
                continue

            empty_streak = 0
            lines.append(raw_input.replace("\\n", "\n"))

        answer = "\n".join(lines).strip()
        return "answer", answer if answer else None


def _detect_terminal_width(default: int = 80) -> int:
    try:
        columns = shutil.get_terminal_size().columns
    except OSError:  # pragma: no cover - depends on runtime
        return default
    return max(40, min(columns, 120))
