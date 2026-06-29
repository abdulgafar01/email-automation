# Email Automation — Outlook Draft Generator

Creates personalized Outlook **drafts** from an existing master template draft,
one per row in an Excel spreadsheet. **It never sends email.**

## How it works

```
Excel (openpyxl) ─▶ Connect Outlook (pywin32) ─▶ Locate template draft
        ─▶ Copy draft ─▶ Set recipient (Col A) ─▶ Set subject (Cols A–G)
        ─▶ Generate HTML table (all columns) ─▶ Insert into body ─▶ Save draft
```

For every row only the **To**, **Subject** and a generated **HTML table** are
changed. Everything else in the template copy — fonts, colours, images,
hyperlinks, signature, spacing and embedded objects — is preserved.

## Requirements

- Python 3.12+
- Windows with Outlook (desktop) installed and configured
- Dependencies: `pip install -r requirements.txt`

## Setup

1. In Outlook, create a draft that serves as your master template. Give it a
   recognizable subject, e.g. `MASTER TEMPLATE`. Optionally place a
   `{{TABLE}}` placeholder where you want the generated table to appear; if it
   is absent the table is appended to the end of the body.

2. Put your spreadsheet at `data/contacts.xlsx`. The first row is the header
   and is skipped. For each data row:
   - **Recipient** = column A
   - **Subject** = first 7 columns joined by spaces
   - **Table** = every column rendered as one `<table>` row

3. (Optional) Override defaults with environment variables:

   | Variable           | Default              |
   |--------------------|----------------------|
   | `TEMPLATE_SUBJECT` | `MASTER TEMPLATE`    |
   | `TEMPLATE_ENTRYID` | *(unset)*            |
   | `EXCEL_PATH`       | `data/contacts.xlsx` |
   | `LOG_DIR`          | `logs`               |
   | `SUBJECT_COLUMNS`  | `7`                  |
   | `TABLE_PLACEHOLDER`| `{{TABLE}}`          |

## Run

```powershell
pip install -r requirements.txt
python -m src.main
```

Generate a sample workbook for testing:

```powershell
python scripts/make_sample_xlsx.py
```

## Project layout

```
data/            contacts.xlsx (input)
logs/            timestamped run logs
scripts/         make_sample_xlsx.py (test helper)
src/
  config.py              settings + env overrides
  models.py              Contact dataclass (positional row)
  excel_reader.py        openpyxl reader (skips empty rows)
  html_table_generator.py  reusable HTML table builder
  outlook_service.py     pywin32 COM service (copy/subject/table/save)
  logger.py              file + console logging
  exceptions.py          typed error hierarchy
  main.py                orchestration + summary
```

## Safety

- The original template draft is copied via `MailItem.Copy`; it is never modified.
- Only `To`, `Subject` and the generated table change — all other formatting is preserved.
- Table cell values are HTML-escaped to avoid breaking the body or injection.
- `MailItem.Save()` is used — `.Send()` is never called.
- One failing row does not stop the run; a summary is logged at the end.
