# `src/models.py` — The `Contact` data model

## What it does
Defines `Contact`, a small dataclass that represents **one row** of the
spreadsheet, plus helper methods to get the recipient, build the subject, and
check if the row is empty.

## Why it exists
The rest of the app shouldn't deal in raw lists of cells (`row[0]`, `row[6]`…).
A `Contact` object gives those values **meaning** (`contact.recipient`,
`contact.subject(7)`), so the Outlook code reads like the business rules and not
like spreadsheet plumbing.

## Why it is built this way

### Positional, not header-based
```python
@dataclass
class Contact:
    values: List[str]
    row_number: int = 0
    headers: List[str] = ...
```
The requirements are **positional**: recipient = column A, subject = first 7
columns, table = every column. So `Contact` stores `values` as an ordered list
rather than a `{header: value}` dict. This matches the spec exactly and works
even when column headers are inconsistent or missing.

### `recipient` as a property
```python
@property
def recipient(self) -> str:
    return self.values[0] if self.values else ""
```
Column A is the recipient. Exposing it as `contact.recipient` documents that
rule in one place and guards against an empty row (no `IndexError`).

### `subject(columns=7)` as a method
```python
def subject(self, columns: int = 7) -> str:
    return " ".join(str(v).strip() for v in self.values[:columns])
```
Joins the first N columns with spaces, exactly as the spec's example shows. It's
a method (not a fixed string) so the count is configurable via
`SUBJECT_COLUMNS`.

### No email validation
The recipient here can be a reference/account value like `514` or `80409`, not a
classic email address — Outlook resolves it. So the model deliberately does
**not** reject non-email recipients, which an earlier version did.

### `row_number` for logging
Carrying the original 1-based row number means every log line and error can say
"row 5", making it trivial to find the offending spreadsheet row.

## Used by
- `excel_reader.py` builds `Contact` objects.
- `outlook_service.py` reads `contact.recipient`, `contact.subject(...)`,
  `contact.values`.
