"""Application-specific exception hierarchy."""

from __future__ import annotations


class AppError(Exception):
    """Base class for domain-specific errors."""


class ResumeFileMismatchError(AppError):
    """Raised when the resume file does not correspond to the target workbook."""


class NoQuestionsFoundError(AppError):
    """Raised when no questions can be extracted from the workbook."""


class SessionAbort(AppError):
    """Raised when the user chooses to abort the session."""
