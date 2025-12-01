"""Command-line interface parsing."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
from typing import Optional


@dataclass(slots=True)
class AppConfig:
    """Configuration values derived from command-line options."""

    source_path: str
    output_path: Optional[str] = None
    catalog_path: Optional[str] = None


def parse_args(argv: list[str] | None = None) -> AppConfig:
    parser = argparse.ArgumentParser(description="Interactive XLSX questionnaire processor")
    parser.add_argument("source", help="Path to the questionnaire XLSX file")
    parser.add_argument(
        "-o",
        "--output",
        dest="output",
        metavar="FILE",
        help="Where to save the answered workbook",
    )
    parser.add_argument(
        "--catalog",
        dest="catalog",
        metavar="CSV",
        help="Override the default catalog CSV path",
    )

    args = parser.parse_args(argv)
    return AppConfig(source_path=args.source, output_path=args.output, catalog_path=args.catalog)
