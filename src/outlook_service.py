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
import re
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


def _save_debug_html(html: str) -> None:
    """Write *html* to ``debug_template.html`` in the working directory.

    Called once per run so you can open the file in a browser to inspect
    exactly what HTML Outlook returned before any placeholder replacement.
    """
    try:
        path = os.path.join(os.getcwd(), "debug_template.html")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(html)
        log.info("Template HTML saved for inspection: %s", path)
    except Exception as exc:
        log.warning("Could not save debug_template.html: %s", exc)


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
        subject_columns: int = 8,
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
        subject = contact.subject(self.subject_columns)
        new_mail.Subject = subject

        # 3. Debug: capture the template HTML before any modification so it
        #    can be inspected in debug_template.html if something goes wrong.
        raw_html = new_mail.HTMLBody or ""
        raw_plain = new_mail.Body or ""
        ph_found = self._any_placeholder_in(raw_html, raw_plain)
        log.info("Checking template HTML for placeholder...")
        log.info("Placeholder present: %s", ph_found)
        if not ph_found:
            log.warning(
                "Placeholder diagnostics — "
                "'TABLE' word in HTML: %s | "
                "'{' brace in HTML: %s | "
                "HTML body length: %d chars | "
                "Plain body length: %d chars. "
                "Open debug_template.html and search for the placeholder text.",
                "TABLE" in raw_html.upper(),
                "{" in raw_html,
                len(raw_html),
                len(raw_plain),
            )
        _save_debug_html(raw_html)

        # 4. Body — generate table (with headers when available) and insert.
        table_html = generate_html_table(
            contact.values,
            headers=contact.headers or None,
        )
        strategy = "Append at end"
        try:
            html_body, plain_body, strategy = self._insert_table(
                raw_html,
                table_html,
                raw_plain,
            )
            if html_body is not None:
                new_mail.HTMLBody = html_body
            if plain_body is not None:
                new_mail.Body = plain_body
        except Exception as exc:  # pragma: no cover
            log.warning("_insert_table failed (%s); appending to Body", exc)
            new_mail.Body = f"{raw_plain}\n\n{table_html}"
            strategy = "Plain body append (exception fallback)"

        new_mail.Save()  # saved as a draft; never .Send()

        entry_id = str(new_mail.EntryID)
        log.info(
            "Recipient: %s | Subject: %s | Placeholder Found: %s | "
            "Insertion Strategy: %s | Draft Saved: EntryID %s",
            contact.recipient,
            subject,
            ph_found,
            strategy,
            entry_id,
        )
        return entry_id

    # --- helpers ----------------------------------------------------------
    # Unicode bidirectional / zero-width control characters that Outlook's
    # Word HTML engine silently inserts between characters of LTR text
    # (e.g. {{TABLE}}) embedded in an Arabic (RTL) email body.
    _BIDI_MARKS: str = (
        "\u200B\u200C\u200D\u200E\u200F"   # zero-width & directional marks
        "\u202A\u202B\u202C\u202D\u202E"   # LTR/RTL embedding & override
        "\uFEFF"                            # BOM / zero-width no-break space
    )
    # Additional fixed placeholders tried before the configured token so
    # users can embed a cleaner marker in their master template.
    _HTML_COMMENT_PH: str = "<!--TABLE_PLACEHOLDER-->"
    _DIV_PH: str = '<div id="TABLE_PLACEHOLDER"></div>'

    def _any_placeholder_in(self, html: str, plain: str) -> bool:
        """Return True if any supported placeholder is found in *html* or *plain*."""
        candidates = [self._HTML_COMMENT_PH, self._DIV_PH, self.table_placeholder]
        return any(ph and (ph in html or ph in plain) for ph in candidates)

    def _insert_table(
        self,
        html_body: Optional[str],
        table_html: str,
        plain_body: Optional[str] = None,
    ) -> tuple[Optional[str], Optional[str], str]:
        """Insert *table_html* into the draft body.

        Returns ``(html_body, plain_body, strategy_name)``.  Exactly one of
        *html_body* / *plain_body* is non-None when a placeholder is replaced;
        on fallback append both html variants may be non-None.

        Replacement strategies tried in order:
        1. HTML comment  ``<!--TABLE_PLACEHOLDER-->``
        2. DIV element   ``<div id="TABLE_PLACEHOLDER"></div>``
        3. Literal       configured token (default ``{{TABLE}}``)
        4. Entity-encoded curly braces (``&#123;``/``&#125;``)
        5. Space-padded  ``{{ TABLE }}``
        6. Regex — tolerates Unicode bidi marks and ``<span>``/``<font>`` tags
           inserted between characters (Arabic / RTL Outlook drafts)
        7. Plain-text body match
        8. Append before ``</body>`` (final fallback — logs a warning)
        """
        body = html_body or ""

        # 1. HTML comment placeholder
        if self._HTML_COMMENT_PH in body:
            return (
                body.replace(self._HTML_COMMENT_PH, table_html, 1),
                None,
                "HTML comment placeholder",
            )

        # 2. DIV placeholder
        if self._DIV_PH in body:
            return (
                body.replace(self._DIV_PH, table_html, 1),
                None,
                "DIV placeholder",
            )

        if self.table_placeholder:
            ph = self.table_placeholder

            # 3. Literal match
            if ph in body:
                return body.replace(ph, table_html, 1), None, "Literal {{TABLE}}"

            # 4. HTML entity-encoded curly braces (&#123; / &#125;)
            encoded = ph.replace("{", "&#123;").replace("}", "&#125;")
            if encoded != ph and encoded in body:
                return body.replace(encoded, table_html, 1), None, "HTML entity replacement"

            # 5. Space-padded variant some Outlook versions produce
            padded = ph.replace("{{", "{{ ").replace("}}", " }}")
            if padded != ph and padded in body:
                return body.replace(padded, table_html, 1), None, "Space-padded {{TABLE}}"

            # 6. Regex: between every character of the placeholder allow
            #    Unicode bidi/zero-width marks and Word-HTML span/font tags.
            #    Non-raw string so \uXXXX are real Unicode code points.
            _between = (
                "(?:"
                "[" + self._BIDI_MARKS + "]"
                "|</?(?:span|font)[^>]{0,400}>"
                ")*"
            )
            pattern = _between.join(re.escape(c) for c in ph)
            m = re.search(pattern, body, re.DOTALL)
            if m:
                return (
                    body[: m.start()] + table_html + body[m.end() :],
                    None,
                    "Regex (bidi/span-tolerant)",
                )

            # 7. Plain-text body match
            if plain_body and ph in plain_body:
                return (
                    None,
                    plain_body.replace(ph, table_html, 1),
                    "Plain text replacement",
                )

        # 8. No placeholder found — warn and append before </body>
        log.warning(
            "Placeholder '%s' not found in template HTML or plain body; "
            "appending table before </body>.",
            self.table_placeholder,
        )
        lower = body.lower()
        idx = lower.rfind("</body>")
        if idx != -1:
            return body[:idx] + table_html + body[idx:], None, "Append before </body>"
        return body + table_html, None, "Append at end"


__all__ = ["OutlookService"]
