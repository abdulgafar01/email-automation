# Project structure — folders & support files

This explains the non-code files and the folder layout, and why each is here.

## Folder layout
```
email-automation/
├── data/          your input .xlsx (auto-discovered)
├── docs/          this documentation
├── logs/          one timestamped log file per run
├── output/        reserved for future "preview" HTML output
├── scripts/       developer helpers (sample data generator)
├── src/           the application package (the actual product)
├── templates/     reserved for future template assets (e.g. .oft files)
├── .gitignore
├── README.md      quick-start usage
├── requirements.txt
└── run.ps1        launcher for manual/scheduled runs
```

### Why separate `data/`, `logs/`, `output/`, `templates/`
Each kind of file has a predictable home. Inputs go in `data/`, run records in
`logs/`, and the app never has to guess where things are. Empty folders are kept
in git with a `.gitkeep` placeholder so the structure exists on a fresh clone.

### Why `src/` is a package
Keeping all application code under one `src/` package (with `__init__.py`) means
clean imports (`from src.config import ...`) and a clear line between "the
product" and helper scripts, docs and data.

## `requirements.txt`
```
openpyxl==3.1.5
pywin32==311
```
- **Pins exact versions** so the remote PC installs the same, tested libraries —
  no surprise breakage from a newer release.
- `openpyxl` reads the Excel file; `pywin32` drives Outlook via COM.
- These are the **only** two third-party dependencies; everything else
  (pathlib, logging, dataclasses, html, tempfile) is in the Python standard
  library, keeping the install small and locked-down-PC friendly.
- Note: the file was rewritten as clean UTF-8. PowerShell's
  `pip freeze > requirements.txt` writes UTF-16, which `pip install` can't read.

## `.gitignore`
Keeps the repository clean and safe by **not committing**:
- Python caches and build artifacts (`__pycache__/`, `*.pyc`, `*.egg-info/`).
- The virtual environment (`.venv/`).
- Runtime output: `logs/*.log`, `output/`, and **`data/*.xlsx`**.
- IDE/OS junk and Outlook temp files (`~$*`).

> Important consequence: because `data/*.xlsx` is ignored, your real spreadsheet
> is **not** copied when you clone the repo. On a new machine you must place the
> `.xlsx` into `data/` manually. The `.gitkeep` files keep the empty folders.

## `README.md`
The short, task-focused quick-start (install, set up the template, run). This
`docs/` folder is the deeper "why" companion to that "how".

## `src/__init__.py`
Marks `src/` as a Python package so `python -m src.main` and
`from src.x import y` work. It's intentionally tiny — its job is structural, not
behavioural.
