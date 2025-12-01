"""Progress persistence utilities."""

from __future__ import annotations

import json
import os
from datetime import datetime, timezone
from typing import List

from .domain import ProgressSnapshot, Question


class ProgressRepository:
    """Serialize and deserialize progress snapshots."""

    def __init__(self, directory: str | None = None) -> None:
        self._directory = directory

    def save(self, filepath: str, questions: List[Question], current_index: int, source_file: str) -> str:
        snapshot = ProgressSnapshot(
            source_file=source_file,
            timestamp=datetime.now(tz=timezone.utc),
            current_index=current_index,
            questions=questions,
        )

        target_path = self._resolve_path(filepath)
        directory = os.path.dirname(target_path)
        if directory:
            os.makedirs(directory, exist_ok=True)
        with open(target_path, "w", encoding="utf-8") as stream:
            json.dump(snapshot.to_dict(), stream, indent=2)

        return target_path

    def load(self, filepath: str) -> ProgressSnapshot:
        target_path = self._resolve_path(filepath)
        if not os.path.isfile(target_path):
            raise FileNotFoundError(f"Progress file not found: {target_path}")

        with open(target_path, "r", encoding="utf-8") as stream:
            payload = json.load(stream)

        return ProgressSnapshot.from_dict(payload)

    def _resolve_path(self, filepath: str) -> str:
        if os.path.isabs(filepath):
            return filepath
        base_dir = self._directory or os.getcwd()
        return os.path.join(base_dir, filepath)
