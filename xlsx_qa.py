#!/usr/bin/env python3
"""CLI entry point for the XLSX questionnaire processor."""

from __future__ import annotations

import sys

from xlsx_qa.bootstrap import bootstrap


def main(argv: list[str] | None = None) -> int:
    bootstrap(argv)

    from xlsx_qa.app import XLSXQuestionnaireApp
    from xlsx_qa.cli import parse_args
    from xlsx_qa.errors import NoQuestionsFoundError

    config = parse_args(argv)
    app = XLSXQuestionnaireApp()

    try:
        return app.run(config)
    except NoQuestionsFoundError as error:
        print(error)
        return 1
    except KeyboardInterrupt:
        print("\nSession interrupted.")
        return 1


if __name__ == "__main__":
    sys.exit(main())
