"""Application entry point: orchestrates Excel -> Outlook draft creation.

Run from the project root:

    python -m src.main

The program reads contacts from the configured workbook, locates the master
template draft in Outlook, and saves one personalized draft per contact.
It never sends email.
"""

from __future__ import annotations

import sys

from src.config import load_config
from src.excel_reader import ExcelReader
from src.exceptions import (
    EmailAutomationError,
    OutlookConnectionError,
    TemplateNotFoundError,
    WorkbookError,
)
from src.logger import get_logger, setup_logging
from src.outlook_service import OutlookService


def run() -> int:
    cfg = load_config()
    setup_logging(cfg.log_dir)
    log = get_logger("main")

    log.info("=== Email automation started (DRAFTS ONLY — never sends) ===")

    # 1. Read contacts ----------------------------------------------------
    log.info("Using workbook: %s", cfg.excel_path)
    try:
        reader = ExcelReader(cfg.excel_path, lookup_path=cfg.lookup_excel_path)
        contacts = reader.read()
        if cfg.max_rows and cfg.max_rows > 0:
            contacts = contacts[: cfg.max_rows]
            log.info("Limiting to first %d rows (MAX_ROWS)", cfg.max_rows)
    except WorkbookError as exc:
        log.error("Workbook error: %s", exc)
        log.error("Place an .xlsx file in: %s", cfg.data_dir)
        return 1

    # 2. Connect to Outlook & locate template -----------------------------
    service = OutlookService(
        template_subject=cfg.template_subject,
        template_entryid=cfg.outlook_entryid,
        subject_columns=cfg.subject_columns,
        table_placeholder=cfg.table_placeholder,
        cc_address=cfg.cc_address,
    )
    try:
        service.connect()
        service.locate_template()
    except (OutlookConnectionError, TemplateNotFoundError) as exc:
        log.error("Setup error: %s", exc)
        return 1

    # 3. Process each contact --------------------------------------------
    created, failed = 0, 0
    for contact in contacts:
        try:
            service.create_draft_for(contact)
            created += 1
        except EmailAutomationError as exc:
            failed += 1
            log.error("Failed row %d (%s): %s", contact.row_number, contact.recipient, exc)
        except Exception as exc:  # keep going on unexpected COM errors
            failed += 1
            log.exception("Unexpected error on row %d: %s", contact.row_number, exc)

    # 4. Summary ----------------------------------------------------------
    log.info(
        "=== Summary: %d draft(s) created, %d failed, %d total ===",
        created,
        failed,
        len(contacts),
    )
    return 0 if failed == 0 else 2


if __name__ == "__main__":
    sys.exit(run())
