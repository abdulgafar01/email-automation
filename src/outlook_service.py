"""Outlook automation via pywin32 (COM).

Locates a master template draft, copies it, then for each row sets the
recipient (column A), builds the subject (first N columns) and inserts a
dynamically generated HTML table into the body — either replacing a
``{{TABLE}}`` placeholder or appending it while preserving all existing
formatting, images, hyperlinks and the Outlook signature. The original
template is never modified and no email is ever sent.
"""

from __future__ import annotations

import os
import tempfile
from typing import Optional

from src.exceptions import OutlookConnectionError, TemplateNotFoundError
from src.html_table_generator import generate_html_table
from src.logger import get_logger
from src.models import Contact

log = get_logger(__name__)

# Outlook OlDefaultFolders constant for the Drafts folder.
OL_FOLDER_DRAFTS = 16
OL_FORMAT_HTML = 2
# OlSaveAsType.olTemplate — saves a mail item as a reusable .oft template.
OL_SAVE_AS_TEMPLATE = 2


class OutlookService:
    """Thin wrapper around the Outlook COM application.

    Args:
        template_subject:  Subject used to locate the template draft.
        template_entryid:  Optional EntryID; takes precedence over subject.
        subject_columns:   Number of leading columns joined to form the subject.
        table_placeholder: Token replaced with the generated table; appended if
                           the token is not present in the body.
        cc_address:        Fixed Cc applied to every draft (';'-separated for
                           multiple addresses). Empty string leaves Cc as-is.
    """

    def __init__(
        self,
        template_subject: str,
        template_entryid: Optional[str] = None,
        subject_columns: int = 7,
        table_placeholder: str = "{{TABLE}}",
        cc_address: str = "",
    ) -> None:
        self.template_subject = template_subject
        self.template_entryid = template_entryid
        self.subject_columns = subject_columns
        self.table_placeholder = table_placeholder
        self.cc_address = cc_address
        self._app = None
        self._namespace = None
        self._template = None
        self._oft_path: Optional[str] = None

    # --- connection -------------------------------------------------------
    def connect(self) -> None:
        """Connect to a running/!startable Outlook instance."""
        try:
            import win32com.client  # imported lazily so non-Windows can import module
        except ImportError as exc:  # pragma: no cover
            raise OutlookConnectionError(
                "pywin32 is not installed. Run: pip install pywin32"
            ) from exc

        try:
            self._app = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._app.GetNamespace("MAPI")
        except Exception as exc:  # pragma: no cover - COM errors
            raise OutlookConnectionError(
                f"Could not connect to Outlook: {exc}"
            ) from exc

        log.info("Outlook connected")

    # --- template ---------------------------------------------------------
    def locate_template(self):
        """Find and cache the template draft. Returns the COM mail item."""
        if self._namespace is None:
            raise OutlookConnectionError("connect() must be called first")

        if self.template_entryid:
            try:
                self._template = self._namespace.GetItemFromID(self.template_entryid)
                log.info("Template located by EntryID")
                self._cache_template_as_oft()
                return self._template
            except Exception as exc:
                raise TemplateNotFoundError(
                    f"No draft with EntryID '{self.template_entryid}': {exc}"
                ) from exc

        drafts = self._namespace.GetDefaultFolder(OL_FOLDER_DRAFTS)
        for item in drafts.Items:
            try:
                if str(item.Subject).strip() == self.template_subject:
                    self._template = item
                    log.info("Template located by subject: '%s'", self.template_subject)
                    self._cache_template_as_oft()
                    return item
            except Exception:  # pragma: no cover - odd items
                continue

        raise TemplateNotFoundError(
            f"No draft found with subject '{self.template_subject}'"
        )

    def _cache_template_as_oft(self) -> None:
        """Save the template as a temp .oft so new drafts can be created from it.

        ``MailItem.Copy()`` fails on *inline response* items (templates created
        via Reply/Forward) with "This method can't be used with an inline
        response mail item." Creating drafts from an .oft template sidesteps
        that entirely while preserving formatting, images and the signature.
        Falls back silently to ``Copy()`` if the template can't be saved.
        """
        self._oft_path = None
        try:
            fd, path = tempfile.mkstemp(suffix=".oft")
            os.close(fd)
            os.remove(path)  # SaveAs needs to create the file itself
            self._template.SaveAs(path, OL_SAVE_AS_TEMPLATE)
            self._oft_path = path
            log.info("Template cached as Outlook template (.oft)")
        except Exception as exc:  # pragma: no cover - COM quirks
            log.warning(
                "Could not cache template as .oft (%s); falling back to Copy()",
                exc,
            )
            self._oft_path = None

    def _new_from_template(self):
        """Return a fresh mail item duplicated from the template.

        Prefers ``CreateItemFromTemplate`` (works for inline-response
        templates); falls back to ``MailItem.Copy()``.
        """
        if self._oft_path:
            return self._app.CreateItemFromTemplate(self._oft_path)
        return self._template.Copy()

    # --- draft creation ---------------------------------------------------
    def create_draft_for(self, contact: Contact) -> str:
        """Create a personalized draft for *contact*. Returns the new EntryID.

        A fresh item is duplicated from the template (via an .oft template, or
        ``MailItem.Copy`` as a fallback) so the original stays untouched. Only
        the recipient, Cc, subject and the generated HTML table are changed —
        all other formatting is preserved.
        """
        if self._template is None:
            self.locate_template()

        new_mail = self._new_from_template()

        # 1. Recipient — column A.
        new_mail.To = contact.recipient

        # 1b. Cc — fixed address(es) applied to every draft, if configured.
        if self.cc_address:
            new_mail.CC = self.cc_address

        # 2. Subject — first N columns joined by spaces.
        new_mail.Subject = contact.subject(self.subject_columns)

        # 3. Body — insert a freshly generated table for every column.
        table_html = generate_html_table(contact.values)
        try:
            new_mail.HTMLBody = self._insert_table(new_mail.HTMLBody, table_html)
        except Exception:  # pragma: no cover - plain-text template fallback
            new_mail.Body = f"{new_mail.Body}\n\n{table_html}"

        new_mail.Save()  # saved as a draft; never .Send()

        entry_id = str(new_mail.EntryID)
        log.info(
            "Draft created for '%s' (row %d) -> EntryID %s",
            contact.recipient,
            contact.row_number,
            entry_id,
        )
        return entry_id

    # --- helpers ----------------------------------------------------------
    def _insert_table(self, html_body: Optional[str], table_html: str) -> str:
        """Insert *table_html* into *html_body*.

        Replaces the configured placeholder when present; otherwise appends the
        table just before ``</body>`` (or at the end) so the existing HTML —
        fonts, images, signature — is preserved.
        """
        body = html_body or ""

        if self.table_placeholder and self.table_placeholder in body:
            return body.replace(self.table_placeholder, table_html)

        lower = body.lower()
        idx = lower.rfind("</body>")
        if idx != -1:
            return body[:idx] + table_html + body[idx:]
        return body + table_html


__all__ = ["OutlookService"]
