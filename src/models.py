"""Domain models for the email-automation project.

A row is treated positionally: the first column is used as the lookup key,
while the final recipient is resolved from a lookup workbook when available.
The subject is built from the first N columns, and the HTML table uses every
column. No email validation is performed because the recipient may be an
account/reference value (e.g. ``80409``) that Outlook resolves.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List


@dataclass
class Contact:
    """A single spreadsheet row.

    Attributes:
        values:     All cell values for the row (already trimmed strings).
        row_number: 1-based row index in the sheet (for logging).
        headers:    Optional column headers (used by the table generator).
        lookup_map: Mapping from the source ID to the recipient email.
    """

    values: List[str] = field(default_factory=list)
    row_number: int = 0
    headers: List[str] = field(default_factory=list)
    lookup_map: Dict[str, str] = field(default_factory=dict)

    @property
    def recipient(self) -> str:
        """Return the resolved recipient email when a lookup mapping exists."""
        source_value = self.values[0] if self.values else ""
        lookup_key = str(source_value).strip()
        if lookup_key and lookup_key in self.lookup_map:
            return self.lookup_map[lookup_key]
        return source_value

    def subject(self, columns: int = 10) -> str:
        """Build the subject by joining the first *columns* values with tabs."""
        return "\t".join(str(v).strip() for v in self.values[:columns])

    def is_empty(self) -> bool:
        """Return True if every cell in the row is blank."""
        return all(str(v).strip() == "" for v in self.values)


__all__ = ["Contact"]
