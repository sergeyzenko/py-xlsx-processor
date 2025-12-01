"""Catalog persistence utilities."""

from __future__ import annotations

import csv
import os
from typing import Dict, Iterable, List, Tuple

from .domain import (
    CATALOG_HEADERS,
    CatalogRecord,
    build_catalog_record,
    catalog_key,
    order_catalog,
)


class CatalogRepository:
    """Load and persist catalog CSV files."""

    def __init__(self, directory: str | None = None) -> None:
        self._directory = directory

    def load(self, filepath: str) -> Dict[Tuple[str, str], CatalogRecord]:
        target_path = self._resolve_path(filepath)
        if not os.path.isfile(target_path):
            return {}

        with open(target_path, "r", encoding="utf-8", newline="") as stream:
            reader = csv.DictReader(stream)
            fieldnames = reader.fieldnames or []
            missing_headers = [header for header in CATALOG_HEADERS if header not in fieldnames]
            if missing_headers:
                raise ValueError(
                    "Catalog file is missing required columns: " + ", ".join(missing_headers)
                )

            records: Dict[Tuple[str, str], CatalogRecord] = {}
            for row in reader:
                record = build_catalog_record(row)
                records[catalog_key(record)] = record

        return records

    def save(self, filepath: str, records: Iterable[CatalogRecord]) -> str:
        target_path = self._resolve_path(filepath)
        directory = os.path.dirname(target_path)
        if directory:
            os.makedirs(directory, exist_ok=True)

        ordered_records = order_catalog(records)

        with open(target_path, "w", encoding="utf-8", newline="") as stream:
            writer = csv.writer(stream, quoting=csv.QUOTE_ALL, lineterminator="\n")
            writer.writerow(CATALOG_HEADERS)
            for record in ordered_records:
                writer.writerow(record.as_csv_row())

        return target_path

    def _resolve_path(self, filepath: str) -> str:
        if os.path.isabs(filepath):
            return filepath
        base_dir = self._directory or os.getcwd()
        return os.path.join(base_dir, filepath)
