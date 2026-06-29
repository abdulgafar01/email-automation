# `scripts/make_sample_xlsx.py` — Test data generator

## What it does
Creates a sample `data/contacts.xlsx` you can use to test the whole pipeline
without your real data. Run it with `python scripts/make_sample_xlsx.py`.

## Why it exists
You shouldn't need the real, sensitive spreadsheet just to check that the code
works. This script produces a realistic file — including tricky cases — so you
(or anyone new to the project) can run an end-to-end test in seconds.

## Why it is built this way

### Mirrors the real format
The sample uses the same shape as production data: column A is the recipient,
the first seven columns form the subject, and there are extra columns to prove
the table handles "every column".

### Deliberately includes edge cases
- A **blank row** — to prove empty rows are ignored.
- A row with **leading/trailing spaces** and `<b>html</b>` — to prove trimming
  and HTML-escaping work.

These are the same cases the reader and table generator are designed to handle,
so running the script is also a quick regression check.

### Kept in `scripts/`, not `src/`
It's a developer helper, not part of the application. Separating it keeps the
`src/` package focused on the actual product. It's also excluded from the data
that ships, and the generated `.xlsx` is git-ignored.

## Note
This file is for testing only. In real use you simply drop your own `.xlsx` into
`data/` (see [config.md](config.md) for auto-discovery).
