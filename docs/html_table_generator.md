---
title: html_table_generator.py
---

# `src/html_table_generator.py` — Building the HTML table

!!! abstract "At a glance"
    **Responsibility:** turn a row's values into a standalone HTML `<table>`.
    **Depends on:** stdlib (`html.escape`). **Pure:** yes — the easiest module to
    test.

## Why it exists

The spec says: don't edit an existing table — **generate a new one** for every
row, containing every column, for any number of columns. Isolating this makes it
**independently testable** (no Outlook, no Excel) and reusable.

## Reference

### `generate_html_table(values, headers=None) -> str`

| Param | Meaning |
| --- | --- |
| `values` | The row's cell values → one `<td>` each |
| `headers` | Optional labels → renders a `<th>` header row |

**Returns:** a complete `<table>…</table>` string.

```python
from src.html_table_generator import generate_html_table

generate_html_table(["80409", "121014959", "SUSPENDED"])
```
```html
<table border="1" cellspacing="0" cellpadding="5">
  <tr><td>80409</td><td>121014959</td><td>SUSPENDED</td></tr>
</table>
```

With headers:

```python
generate_html_table(["80409", "SUSPENDED"], headers=["Account", "Status"])
```
```html
<table border="1" cellspacing="0" cellpadding="5">
  <tr><th>Account</th><th>Status</th></tr>
  <tr><td>80409</td><td>SUSPENDED</td></tr>
</table>
```

### `_cell(value) -> str` (internal)

Renders one escaped `<td>`; `None` becomes an empty cell.

## Design decisions

??? note "Why a pure function?"
    Data in, string out — no Outlook, no files, no globals. Pure functions are the
    easiest code to test and reason about; you can assert the exact HTML.

??? note "Why escape every cell? (security)"
    ```python
    from html import escape
    f"<td>{escape(text)}</td>"
    ```
    Spreadsheet text may contain `<`, `>`, `&`, even `<script>`. Without escaping,
    such content could **break the email's HTML** or inject markup. `html.escape`
    turns `<b>` into `&lt;b&gt;`, so the cell shows literal text and the
    surrounding Outlook body stays intact. This addresses the OWASP injection risk
    for generated markup.

??? note "Why support optional headers if the flow doesn't use them?"
    Cheap future-proofing: labelled tables become a one-argument change, with zero
    cost today.

## Tested behaviour

| Input | Output cell |
| --- | --- |
| `"110"` | `<td>110</td>` |
| `"<b>html</b>"` | `<td>&lt;b&gt;html&lt;/b&gt;</td>` |
| `None` | `<td></td>` |
| `" Dave "` | `<td>Dave</td>` |

## See also

- [`outlook_service.py`](outlook_service.md) — inserts this table into the body
- [Data contract](reference/data-contract.md) — which columns appear
