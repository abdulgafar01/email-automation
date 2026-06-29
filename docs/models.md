---
title: models.py
---

# `src/models.py` — The `Contact` data model

!!! abstract "At a glance"
    **Responsibility:** represent one spreadsheet row with meaningful accessors.
    **Depends on:** stdlib. **Pure:** yes — trivially unit-testable.

## Why it exists

The rest of the app shouldn't juggle raw cells (`row[0]`, `row[6]`). A `Contact`
gives those values **meaning** (`contact.recipient`, `contact.subject(7)`), so the
Outlook code reads like the business rules, not spreadsheet plumbing.

## Reference

### `class Contact`

```python
@dataclass
class Contact:
    values: list[str]            # all cells of the row, trimmed
    row_number: int = 0          # 1-based row index, for logging
    headers: list[str] = ...     # column headers (optional)
```

#### `recipient: str` (property)

```python
@property
def recipient(self) -> str:
    return self.values[0] if self.values else ""
```

Column A is the recipient. Exposed as a property to document the rule once and to
guard against an empty row (no `IndexError`).

#### `subject(columns: int = 7) -> str`

```python
def subject(self, columns: int = 7) -> str:
    return " ".join(str(v).strip() for v in self.values[:columns])
```

Joins the first *N* columns with spaces — exactly the spec example. A method (not
a stored string) so the count stays configurable via `SUBJECT_COLUMNS`.

#### `is_empty() -> bool`

True when every cell is blank; used by the reader to skip empty rows.

## Example

```python
c = Contact(values=["514", "1526", "KOC-1", "Name", "ID", "SUSPENDED", "110", "n"],
            row_number=2)
c.recipient      # "514"
c.subject(7)     # "514 1526 KOC-1 Name ID SUSPENDED 110"
c.is_empty()     # False
```

## Design decisions

??? note "Why positional (`values`) instead of a `{header: value}` dict?"
    The requirements are positional: recipient = column A, subject = first 7,
    table = every column. A list matches that exactly and works even when headers
    are inconsistent or missing.

??? note "Why no email validation?"
    The recipient can be a reference/account value like `514` or `80409`, which
    Outlook resolves. Validating it as an email would wrongly reject valid rows.

??? note "Why carry `row_number`?"
    Every log line and error can then say “row 5”, making the offending
    spreadsheet row trivial to find.

## See also

- [`excel_reader.py`](excel_reader.md) — builds `Contact`s
- [`outlook_service.py`](outlook_service.md) — consumes `recipient`, `subject`,
  `values`
- [Data contract](reference/data-contract.md)
