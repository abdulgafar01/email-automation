"""Domain models for the email-automation project.

A row is treated positionally: the recipient comes from the first column,
the subject is built from the first N columns, and the HTML table uses every
column. No email validation is performed because the recipient may be an
account/reference value (e.g. ``80409``) that Outlook resolves.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import List


@dataclass
class Contact:
    """A single spreadsheet row.

    Attributes:
        values:     All cell values for the row (already trimmed strings).
        row_number: 1-based row index in the sheet (for logging).
        headers:    Optional column headers (used by the table generator).
    """

    values: List[str] = field(default_factory=list)
    row_number: int = 0
    headers: List[str] = field(default_factory=list)

    @property
    def recipient(self) -> str:
        """Recipient address/value taken from the first column (Column A)."""
        return self.values[0] if self.values else ""

    def subject(self, columns: int = 10) -> str:
        """Build the subject by joining the first *columns* values with tabs."""
        return "\t".join(str(v).strip() for v in self.values[:columns])

    def is_empty(self) -> bool:
        """Return True if every cell in the row is blank."""
        return all(str(v).strip() == "" for v in self.values)


__all__ = ["Contact"]
