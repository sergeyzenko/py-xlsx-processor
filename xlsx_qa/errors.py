"""Application-specific exception hierarchy."""

from __future__ import annotations


class AppError(Exception):
    """Base class for domain-specific errors."""


class NoQuestionsFoundError(AppError):
    """Raised when no questions can be extracted from the workbook."""
