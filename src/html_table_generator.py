"""Reusable HTML table generator.

Builds a standalone HTML ``<table>`` from a row's values, regardless of the
number of columns. Values are HTML-escaped to avoid breaking the surrounding
Outlook HTML body (and to prevent injection from spreadsheet content).
"""

from __future__ import annotations

from html import escape
from typing import Iterable, Optional, Sequence

_TABLE_ATTRS = 'border="1" cellspacing="0" cellpadding="5" dir="ltr" style="direction:ltr; text-align:left;"'


def _cell(value: object) -> str:
    """Render a single ``<td>`` from *value* (escaped)."""
    text = "" if value is None else str(value).strip()
    return f'<td dir="ltr" style="direction:ltr; text-align:left;">{escape(text)}</td>'


def generate_html_table(
    values: Sequence[object],
    headers: Optional[Iterable[object]] = None,
) -> str:
    """Generate an HTML table containing every value in *values*.

    Args:
        values:  The row's cell values, one ``<td>`` per value.
        headers: Optional column headers; when provided they are rendered as a
                 bold ``<th>`` header row above the data row.

    Returns:
        A complete ``<table>...</table>`` HTML string. Works for any number of
        columns.
    """
    rows_html = []

    if headers is not None:
        header_cells = "".join(
            f"<th>{escape('' if h is None else str(h).strip())}</th>"
            for h in headers
        )
        rows_html.append(f"<tr>{header_cells}</tr>")

    data_cells = "".join(_cell(v) for v in values)
    rows_html.append(f"<tr>{data_cells}</tr>")

    body = "".join(rows_html)
    return f"<table {_TABLE_ATTRS}>{body}</table>"


__all__ = ["generate_html_table"]
