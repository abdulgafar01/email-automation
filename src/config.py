from dataclasses import dataclass
from pathlib import Path
from typing import Optional
import os

# Project root = the folder that contains this `src` package. Used to resolve
# relative paths so the app works regardless of the current working directory.
PROJECT_ROOT = Path(__file__).resolve().parent.parent


def _resolve(path: Path) -> Path:
    """Anchor a relative path to the project root; leave absolute paths as-is."""
    return path if path.is_absolute() else (PROJECT_ROOT / path)


@dataclass(frozen=True)
class AppConfig:
    template_subject: str
    excel_path: Path
    data_dir: Path
    log_dir: Path
    drafts_folder: str
    outlook_entryid: Optional[str]
    subject_columns: int
    table_placeholder: str
    cc_address: str
    never_send: bool
    max_rows: Optional[int] = None


# --- Defaults ---
# Environment variables override these.
DEFAULT_TEMPLATE_SUBJECT = os.getenv("TEMPLATE_SUBJECT", "No subject")
EXPLICIT_EXCEL_PATH = os.getenv("EXCEL_PATH", "")
DEFAULT_DATA_DIR = Path(__file__).resolve().parent
DEFAULT_LOG_DIR = Path(__file__).resolve().parent
DEFAULT_DRAFTS_FOLDER = "drafts"
DEFAULT_OUTLOOK_ENTRYID = os.getenv("TEMPLATE_ENTRYID", "")
DEFAULT_SUBJECT_COLUMNS = 1
DEFAULT_TABLE_PLACEHOLDER = os.getenv("TABLE_PLACEHOLDER", "{{TABLE}}")
# Fixed Cc applied to every draft. Use ';' to separate multiple addresses.
DEFAULT_CC_ADDRESS = os.getenv("CC_ADDRESS", "")
DEFAULT_MAX_ROWS = int(val) if (val := os.getenv("MAX_ROWS")) else None


def discover_excel(data_dir: Path, explicit: Optional[str] = None) -> Path:
    """Resolve which workbook to read.

    If *explicit* is given (EXCEL_PATH), it is used as-is. Otherwise the first
    ``.xlsx`` file in *data_dir* (alphabetical, ignoring Excel lock files
    ``~$``) is chosen. When none exist, returns ``data_dir/contacts.xlsx`` so
    callers raise a clear "not found" error pointing at the data folder.
    """
    if explicit:
        return _resolve(Path(explicit))
    candidates = sorted(
        p for p in data_dir.glob("*.xlsx") if not p.name.startswith("~$")
    )
    return candidates[0] if candidates else (data_dir / "contacts.xlsx")


def load_config() -> AppConfig:
    """Load application configuration from environment and defaults.

    Keeps configuration centralized so other modules can import a single
    configuration object. Environment variables (optional):
      - TEMPLATE_SUBJECT
      - EXCEL_PATH      (explicit workbook; overrides auto-discovery)
      - DATA_DIR        (folder scanned for an .xlsx when EXCEL_PATH is unset)
      - LOG_DIR
      - DRAFTS_FOLDER
      - TEMPLATE_ENTRYID
      - SUBJECT_COLUMNS
      - TABLE_PLACEHOLDER
      - CC_ADDRESS
      - MAX_ROWS
    """
    cfg = AppConfig(
        template_subject=DEFAULT_TEMPLATE_SUBJECT,
        excel_path=discover_excel(DEFAULT_DATA_DIR, EXPLICIT_EXCEL_PATH),
        data_dir=DEFAULT_DATA_DIR,
        log_dir=DEFAULT_LOG_DIR,
        drafts_folder=DEFAULT_DRAFTS_FOLDER,
        outlook_entryid=DEFAULT_OUTLOOK_ENTRYID,
        subject_columns=DEFAULT_SUBJECT_COLUMNS,
        table_placeholder=DEFAULT_TABLE_PLACEHOLDER,
        cc_address=DEFAULT_CC_ADDRESS,
        never_send=True,  # safety: never send emails from this tool
        max_rows=DEFAULT_MAX_ROWS,
    )
    return cfg


__all__ = ["AppConfig", "load_config", "DEFAULT_TEMPLATE_SUBJECT"]
