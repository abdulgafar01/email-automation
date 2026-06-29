---
title: Glossary
---

# Glossary

Short definitions of terms used throughout these docs.

`AppConfig`
:   The frozen dataclass holding all settings. Produced by `load_config()`. See
    [Configuration reference](configuration.md).

COM (Component Object Model)
:   The Windows technology pywin32 uses to control Outlook. Works only on the
    **same machine** as Outlook, in an interactive session.

`Contact`
:   The in-memory representation of one spreadsheet row. Exposes `recipient`,
    `subject(n)` and the raw `values`. See [`models.py`](../models.md).

Draft
:   An unsent mail item saved in the **Drafts** folder. This tool's only output.

EntryID
:   A unique identifier Outlook assigns to each item. Can be used to pin the
    template precisely (`TEMPLATE_ENTRYID`).

Inline response item
:   A mail item created by Reply/Forward and edited in the reading pane. Outlook
    forbids `Copy()` on these — the reason we use `.oft` templates.

`.oft`
:   An **Outlook File Template**. We save the master draft as one so each new
    draft can be created from it with full formatting preserved.

openpyxl
:   The pure-Python library used to read `.xlsx` files. The project uses **no
    pandas**.

Placeholder (`{{TABLE}}`)
:   A token in the template body that gets replaced by the generated HTML table.
    If absent, the table is appended before `</body>`.

Project root
:   The folder containing `src/`. Relative paths are resolved against it so the
    app is location-independent.

pywin32
:   The library exposing Windows COM APIs (including Outlook) to Python.

Read-only mode
:   openpyxl mode that streams rows for speed/low memory. It trusts the file's
    declared dimension — the cause of the
    [“only N rows” gotcha](data-contract.md#notes-gotchas).

Recipient
:   The `To` value, taken from **column A** of each row.

Template (master draft)
:   The existing Outlook draft used as the design source. Located by subject
    (default `MASTER TEMPLATE`) or EntryID. Never modified.
