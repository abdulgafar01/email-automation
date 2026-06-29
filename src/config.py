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


# Defaults — override with environment variables if needed
DEFAULT_TEMPLATE_SUBJECT = os.getenv("TEMPLATE_SUBJECT", "MASTER TEMPLATE")
DEFAULT_DATA_DIR = _resolve(Path(os.getenv("DATA_DIR", "data")))
# Explicit workbook path (optional). When unset, any .xlsx in DATA_DIR is used.
EXPLICIT_EXCEL_PATH = os.getenv("EXCEL_PATH")
DEFAULT_LOG_DIR = _resolve(Path(os.getenv("LOG_DIR", "logs")))
DEFAULT_DRAFTS_FOLDER = os.getenv("DRAFTS_FOLDER", "Drafts")
DEFAULT_OUTLOOK_ENTRYID = os.getenv("TEMPLATE_ENTRYID")
DEFAULT_SUBJECT_COLUMNS = int(os.getenv("SUBJECT_COLUMNS", "7"))
DEFAULT_TABLE_PLACEHOLDER = os.getenv("TABLE_PLACEHOLDER", "{{TABLE}}")
# Fixed Cc applied to every draft. Use ';' to separate multiple addresses.
DEFAULT_CC_ADDRESS = os.getenv("CC_ADDRESS", "")


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
    )
    return cfg


__all__ = ["AppConfig", "load_config", "DEFAULT_TEMPLATE_SUBJECT"]
