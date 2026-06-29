# `src/exceptions.py` — Custom error types

## What it does
Defines the project's own exception classes in a small hierarchy, all inheriting
from a common base `EmailAutomationError`.

## Why it exists
Different problems need different responses:
- "Outlook isn't running" is **fatal** — stop the whole run.
- "This one row has a bad value" is **recoverable** — skip the row, keep going.

Custom exception types let the code tell these apart cleanly with `try/except`,
instead of guessing from error message strings.

## Why it is built this way

### A common base class
```python
class EmailAutomationError(Exception): ...
```
Callers can catch *all* app errors with one `except EmailAutomationError`, while
still being able to catch specific ones when needed. It also separates "our"
errors from unexpected Python/COM errors.

### Two families: fatal vs per-row
```
EmailAutomationError
├── OutlookConnectionError     (fatal: can't reach Outlook)
├── TemplateNotFoundError      (fatal: no master draft)
├── WorkbookError              (fatal: workbook problems)
│   ├── MissingColumnsError
│   └── EmptyWorkbookError
└── RowProcessingError         (recoverable: affects one row)
    ├── InvalidEmailError
    └── PlaceholderError
```
`main.py` treats the fatal ones as "stop", and the per-row ones as "log and
continue". Grouping them under `WorkbookError` / `RowProcessingError` lets the
code catch a whole family at once.

### Clear, specific names
A raised `TemplateNotFoundError` is self-explanatory in a log file — far better
than a generic `Exception`. The names *are* the documentation.

## Note
Some classes (e.g. `InvalidEmailError`) are kept for future use and the typed
hierarchy even though the current positional-row design doesn't validate emails.
They cost nothing and keep the door open for features like preview/validation.
