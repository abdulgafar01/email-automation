# `src/main.py` — The orchestrator

## What it does
The entry point. Run with `python -m src.main`. It wires everything together:
load config → set up logging → read Excel → connect Outlook → find template →
create a draft per row → print a summary. Returns an exit code.

## Why it exists
Each module does one job; **something** has to call them in the right order and
handle the overall flow and errors. That's `main.py`. Keeping the orchestration
in one short, readable function makes the whole program easy to follow top to
bottom.

## Why it is built this way

### Strict order of operations
```python
cfg = load_config()
setup_logging(cfg.log_dir)
contacts = ExcelReader(cfg.excel_path).read()
service.connect(); service.locate_template()
for contact in contacts: service.create_draft_for(contact)
```
Config and logging come first so everything afterwards is configured and logged.
Excel is read before touching Outlook, so a bad spreadsheet fails fast and
cheap, before opening Outlook.

### Fatal vs recoverable error handling
```python
except WorkbookError:           # fatal → stop, return 1
except (OutlookConnectionError, TemplateNotFoundError):  # fatal → stop, return 1
...
for contact in contacts:
    try: service.create_draft_for(contact)
    except EmailAutomationError: failed += 1   # recoverable → log, continue
    except Exception: failed += 1              # unexpected → log, continue
```
This is the heart of the "continue processing remaining rows even if one fails"
requirement. A single broken row increments `failed` and the loop moves on; the
run still finishes and reports.

### Catching unexpected errors too
The extra `except Exception` around each row catches surprise COM errors (like
the inline-response one) so one weird row can never abort the entire batch.

### A summary line and meaningful exit codes
```python
log.info("=== Summary: %d created, %d failed, %d total ===", ...)
return 0 if failed == 0 else 2
```
- The summary tells you at a glance what happened.
- Exit codes let **Task Scheduler** know the outcome:
  `0` = all good, `1` = fatal setup error, `2` = finished with some row failures.

### Helpful "where to put the file" hint
On a workbook error it logs the `data/` folder path, so the most common mistake
(no spreadsheet present) is obvious.

## How to read a run's output
```
=== Email automation started (DRAFTS ONLY — never sends) ===
Using workbook: ...\data\KOTC JUNE.xlsx
Read 4 contact(s) ...
Outlook connected
Template located by subject: 'MASTER TEMPLATE'
Draft created for '514' (row 2) -> EntryID ...
...
=== Summary: 4 draft(s) created, 0 failed, 4 total ===
```
Each line maps directly to a step above.
