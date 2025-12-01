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
    resume_path: Optional[str] = None


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
        "--resume",
        dest="resume",
        metavar="JSON",
        help="Resume an interrupted session from the provided JSON file",
    )

    args = parser.parse_args(argv)
    return AppConfig(source_path=args.source, output_path=args.output, resume_path=args.resume)
