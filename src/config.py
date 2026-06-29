from dataclasses import dataclass
from pathlib import Path
from typing import Optional
import os


@dataclass(frozen=True)
class AppConfig:
    template_subject: str
    excel_path: Path
    log_dir: Path
    drafts_folder: str
    outlook_entryid: Optional[str]
    subject_columns: int
    table_placeholder: str
    cc_address: str
    never_send: bool


# Defaults — override with environment variables if needed
DEFAULT_TEMPLATE_SUBJECT = os.getenv("TEMPLATE_SUBJECT", "MASTER TEMPLATE")
DEFAULT_EXCEL = Path(os.getenv("EXCEL_PATH", "data/contacts.xlsx"))
DEFAULT_LOG_DIR = Path(os.getenv("LOG_DIR", "logs"))
DEFAULT_DRAFTS_FOLDER = os.getenv("DRAFTS_FOLDER", "Drafts")
DEFAULT_OUTLOOK_ENTRYID = os.getenv("TEMPLATE_ENTRYID")
DEFAULT_SUBJECT_COLUMNS = int(os.getenv("SUBJECT_COLUMNS", "7"))
DEFAULT_TABLE_PLACEHOLDER = os.getenv("TABLE_PLACEHOLDER", "{{TABLE}}")
# Fixed Cc applied to every draft. Use ';' to separate multiple addresses.
DEFAULT_CC_ADDRESS = os.getenv("CC_ADDRESS", "")


def load_config() -> AppConfig:
    """Load application configuration from environment and defaults.

    Keeps configuration centralized so other modules can import a single
    configuration object. Environment variables (optional):
      - TEMPLATE_SUBJECT
      - EXCEL_PATH
      - LOG_DIR
      - DRAFTS_FOLDER
      - TEMPLATE_ENTRYID
      - SUBJECT_COLUMNS
      - TABLE_PLACEHOLDER
      - CC_ADDRESS
    """
    cfg = AppConfig(
        template_subject=DEFAULT_TEMPLATE_SUBJECT,
        excel_path=DEFAULT_EXCEL,
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
