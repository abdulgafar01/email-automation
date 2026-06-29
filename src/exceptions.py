"""Custom exception hierarchy for the email-automation project.

Having typed exceptions lets callers distinguish between recoverable
per-row failures (e.g. invalid email) and fatal setup failures
(e.g. Outlook not running).
"""


class EmailAutomationError(Exception):
    """Base class for all application-specific errors."""


# --- Fatal / setup errors -------------------------------------------------
class OutlookConnectionError(EmailAutomationError):
    """Raised when Outlook cannot be reached or started."""


class TemplateNotFoundError(EmailAutomationError):
    """Raised when the master template draft cannot be located."""


class WorkbookError(EmailAutomationError):
    """Raised for workbook-level problems (missing file, empty, bad headers)."""


class MissingColumnsError(WorkbookError):
    """Raised when required columns are absent from the spreadsheet."""


class EmptyWorkbookError(WorkbookError):
    """Raised when the spreadsheet contains no data rows."""


# --- Per-row / recoverable errors ----------------------------------------
class RowProcessingError(EmailAutomationError):
    """Base class for errors that affect a single row only."""


class InvalidEmailError(RowProcessingError):
    """Raised when a contact's email address is missing or malformed."""


class PlaceholderError(RowProcessingError):
    """Raised when placeholder replacement fails for a row."""


__all__ = [
    "EmailAutomationError",
    "OutlookConnectionError",
    "TemplateNotFoundError",
    "WorkbookError",
    "MissingColumnsError",
    "EmptyWorkbookError",
    "RowProcessingError",
    "InvalidEmailError",
    "PlaceholderError",
]
