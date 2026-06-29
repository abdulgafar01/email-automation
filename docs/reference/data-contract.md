---
title: Data contract (Excel)
---

# Data contract — how your spreadsheet is interpreted

This is the agreement between your `.xlsx` and the app. Follow it and every row
becomes a correct draft.

## Rules

| Rule | Detail |
| --- | --- |
| **Sheet** | The **first** worksheet is read. |
| **Header row** | Row 1 is treated as headers and **skipped** (used only to count columns). |
| **Data rows** | Start at row 2. |
| **Recipient** | **Column A** (`values[0]`) → the draft's `To`. |
| **Subject** | First **7** columns joined by single spaces (configurable via `SUBJECT_COLUMNS`). |
| **Table** | **Every** column of the row becomes one `<td>` in a generated `<table>`. |
| **Empty rows** | Fully blank rows are ignored. |
| **Whitespace** | Every cell is trimmed. |
| **Formulas** | Read as their computed value (`data_only=True`). |

## Worked example

Given this sheet:

| Account | Reference | Contract | Name | ID | Status | Amount | Notes |
| --- | --- | --- | --- | --- | --- | --- | --- |
| 80409 | 121014959 | KOC-10048 -25/26 | YAQOUB … ALKANDARI | 313111600751 | SUSPENDED | 110 | first note |

The app produces:

- **To** → `80409`
- **Subject** → `80409 121014959 KOC-10048 -25/26 YAQOUB … ALKANDARI 313111600751 SUSPENDED 110`
- **Table** (all 8 columns):

```html
<table border="1" cellspacing="0" cellpadding="5">
  <tr>
    <td>80409</td><td>121014959</td><td>KOC-10048 -25/26</td>
    <td>YAQOUB … ALKANDARI</td><td>313111600751</td>
    <td>SUSPENDED</td><td>110</td><td>first note</td>
  </tr>
</table>
```

## Notes & gotchas

!!! note "The recipient need not be an email address"
    Column A can be an account/reference number (e.g. `514`, `80409`). Outlook
    resolves it. The app does **not** validate it as an email.

!!! warning "“Only N rows were read”"
    In openpyxl read-only mode, the row count comes from the worksheet's declared
    **dimension**. Some exporters (SAP, web tools, certain exports) under-declare
    it, truncating the data. Fix by recomputing the range with
    `worksheet.reset_dimensions()` or reading with `read_only=False`. Details in
    [`excel_reader.py`](../excel_reader.md#known-gotcha-only-n-rows-were-read).

!!! tip "Want labelled table cells?"
    The table currently renders values only (per spec). The generator already
    supports an optional header row — see
    [`html_table_generator.py`](../html_table_generator.md).

See also: [Configuration reference](configuration.md) ·
[`models.py`](../models.md) · [`excel_reader.py`](../excel_reader.md).
