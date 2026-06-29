# `src/html_table_generator.py` — Building the HTML table

## What it does
Turns a row's values into a standalone HTML `<table>` string. One reusable
function: `generate_html_table(values, headers=None)`.

## Why it exists
The spec says: don't edit an existing HTML table — **generate a new one** for
every row, containing every column, for any number of columns. Keeping this in
its own module means it can be **tested on its own** (no Outlook, no Excel) and
reused anywhere.

## Why it is built this way

### A pure function
It takes data in and returns a string out — no Outlook, no files, no global
state. Pure functions are the easiest things in a codebase to test and reason
about. You can verify the exact HTML with a one-line test.

### Works for any number of columns
```python
data_cells = "".join(_cell(v) for v in values)
```
It simply loops over whatever values it's given, so a 5-column row and a
20-column row both work with no code change — exactly as required.

### HTML escaping (security + correctness)
```python
from html import escape
return f"<td>{escape(text)}</td>"
```
Spreadsheet text can contain `<`, `>`, `&` or even `<script>`. Without escaping,
such content could **break the email's HTML** or inject markup. `html.escape`
turns `<b>` into `&lt;b&gt;`, so the cell shows the literal text and the
surrounding Outlook body stays intact. This addresses the OWASP "injection"
risk for generated markup.

### Optional headers
```python
def generate_html_table(values, headers=None):
```
The current flow passes only `values` (a single data row, per the spec), but the
function can also render a `<th>` header row if you later want labelled tables —
a small bit of future-proofing that costs nothing now.

### Matches the spec's markup
```python
_TABLE_ATTRS = 'border="1" cellspacing="0" cellpadding="5"'
```
The output mirrors the example in the requirements so the rendered table looks
as expected inside Outlook.

## Example
```python
generate_html_table(["80409", "121014959", "SUSPENDED"])
# -> <table border="1" cellspacing="0" cellpadding="5">
#      <tr><td>80409</td><td>121014959</td><td>SUSPENDED</td></tr>
#    </table>
```
