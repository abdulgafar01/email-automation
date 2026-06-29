# Project Documentation

This folder explains **every file** in the email-automation project: what it
does, why it exists, and why it is built the way it is. Start here, then read
the per-file docs.

## What the project does (in one paragraph)

It reads rows from an Excel file, and for each row it duplicates an existing
**Outlook draft** (your master template), fills in the recipient, subject and a
generated HTML table, and saves the result as a **new draft**. It never sends
email. The original template is never changed.

## Why this design

- **Use a real Outlook draft as the template** instead of building HTML in code.
  That way your fonts, colours, images, hyperlinks and signature come straight
  from Outlook and always look right — we only change three things (To, Subject,
  table).
- **Modular files with one job each.** Each module can be read, tested and
  changed on its own. This makes bugs easier to find and future features (Cc,
  attachments, scheduling) easy to add without touching everything.
- **Safety first.** The code only ever calls `Save()`, never `Send()`, and works
  on a *copy* of the template.

## High-level flow

```
load_config()  ──▶  ExcelReader.read()  ──▶  list[Contact]
                                               │
                                               ▼
                              OutlookService.connect()
                              OutlookService.locate_template()
                                               │
                            for each Contact ──┤
                                               ▼
                       OutlookService.create_draft_for(contact)
                          ├─ duplicate template (.oft / Copy)
                          ├─ set To  (column A)
                          ├─ set Cc  (fixed, optional)
                          ├─ set Subject (first 7 columns)
                          ├─ generate_html_table(all columns)
                          ├─ insert table ({{TABLE}} or append)
                          └─ Save()  ← never Send()
```

## Per-file guides

| File | Doc |
|------|-----|
| `src/config.py` | [config.md](config.md) |
| `src/logger.py` | [logger.md](logger.md) |
| `src/exceptions.py` | [exceptions.md](exceptions.md) |
| `src/models.py` | [models.md](models.md) |
| `src/excel_reader.py` | [excel_reader.md](excel_reader.md) |
| `src/html_table_generator.py` | [html_table_generator.md](html_table_generator.md) |
| `src/outlook_service.py` | [outlook_service.md](outlook_service.md) |
| `src/main.py` | [main.md](main.md) |
| `scripts/make_sample_xlsx.py` | [scripts.md](scripts.md) |
| `run.ps1`, Task Scheduler | [running-and-automation.md](running-and-automation.md) |
| Project files (`requirements.txt`, `.gitignore`, folders) | [project-structure.md](project-structure.md) |
