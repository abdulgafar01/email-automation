# `src/config.py` — Central configuration

## What it does
Defines a single `AppConfig` dataclass and a `load_config()` function that
gathers every setting the app needs (which template, which Excel file, where
logs go, subject column count, Cc address, etc.). It also auto-discovers the
Excel file and resolves all paths.

## Why it exists
Every other module needs settings. Instead of scattering hard-coded values
("MASTER TEMPLATE", "data/contacts.xlsx", 7 columns…) across the codebase, they
live in **one place**. Change a setting here and the whole app follows.

## Why it is built this way

### A frozen dataclass (`@dataclass(frozen=True)`)
`AppConfig` is immutable. Once loaded, no module can accidentally change a
setting at runtime, which prevents a whole class of confusing bugs. Dataclasses
also give clean attribute access (`cfg.template_subject`) and free `repr`.

### Environment variables with sensible defaults
```python
DEFAULT_TEMPLATE_SUBJECT = os.getenv("TEMPLATE_SUBJECT", "MASTER TEMPLATE")
```
Each setting can be overridden with an environment variable, but works out of
the box without one. This means:
- Non-technical use: just run it, defaults apply.
- Different machine/run: set an env var, no code edit, no redeploy.

### Paths anchored to the project root
```python
PROJECT_ROOT = Path(__file__).resolve().parent.parent
def _resolve(path): return path if path.is_absolute() else PROJECT_ROOT / path
```
A relative path like `data/contacts.xlsx` would otherwise depend on the folder
you happened to launch the command from — a common source of "file not found".
Anchoring to the project root makes the app work no matter the current
directory.

### Auto-discovering the workbook (`discover_excel`)
```python
candidates = sorted(p for p in data_dir.glob("*.xlsx") if not p.name.startswith("~$"))
return candidates[0] if candidates else (data_dir / "contacts.xlsx")
```
- You can drop **any** `.xlsx` into `data/` — no need to rename it.
- `~$` lock files (created while Excel has the file open) are skipped.
- If none exist, it returns a clear default path so the error message points you
  at the `data/` folder.
- An explicit `EXCEL_PATH` always wins, for when you need a specific file.

### `never_send=True`
A hard-coded safety flag documenting the project's golden rule: this tool
creates drafts only.

## Key settings
| Setting | Env var | Default | Meaning |
|---------|---------|---------|---------|
| `template_subject` | `TEMPLATE_SUBJECT` | `MASTER TEMPLATE` | Subject used to find the template draft |
| `excel_path` | `EXCEL_PATH` | auto-discovered | The workbook to read |
| `data_dir` | `DATA_DIR` | `data/` | Folder scanned for an `.xlsx` |
| `log_dir` | `LOG_DIR` | `logs/` | Where log files are written |
| `outlook_entryid` | `TEMPLATE_ENTRYID` | unset | Pin the template by exact ID instead of subject |
| `subject_columns` | `SUBJECT_COLUMNS` | `7` | How many leading columns form the subject |
| `table_placeholder` | `TABLE_PLACEHOLDER` | `{{TABLE}}` | Token replaced by the generated table |
| `cc_address` | `CC_ADDRESS` | empty | Fixed Cc on every draft (`;`-separated) |
