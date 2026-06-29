# `src/excel_reader.py` — Reading the spreadsheet

## What it does
Opens an `.xlsx` workbook with **openpyxl**, treats the first row as headers,
and returns one `Contact` per non-empty data row.

## Why it exists
This is the single boundary between "messy spreadsheet" and "clean Python
objects". All the awkward parts of reading Excel — empty cells, whitespace,
ragged rows — are handled here so no other module has to care.

## Why it is built this way

### openpyxl, not pandas
The project requirement is explicitly **no pandas**. openpyxl is lightweight,
reads `.xlsx` natively, and has no heavy numeric dependencies — ideal for a
simple "read rows" job on a locked-down work PC.

### Read-only + data-only mode
```python
load_workbook(self.path, read_only=True, data_only=True)
```
- `read_only=True` streams rows instead of loading the whole file into memory —
  fast and light, and it can't accidentally modify your file.
- `data_only=True` returns the **computed value** of formula cells, not the
  formula text — so a cell showing `110` gives you `110`, not `=A1*2`.

### First row = headers, data starts at row 2
```python
header_row = next(rows)
for row_number, raw in enumerate(rows, start=2):
```
Matches the spec ("treat the first row as headers"). `start=2` keeps
`row_number` aligned with what you see in Excel.

### Column count from the header row
```python
col_count = self._last_non_empty_index(headers) + 1
```
The number of real columns is taken from the last non-empty header. Each data
row is then read to exactly that width, so trailing blank cells don't create
stray empty `<td>`s later, and short rows are padded with `""`.

### Trimming and empty-row handling
```python
@staticmethod
def _clean(value): return "" if value is None else str(value).strip()
...
if contact.is_empty(): continue
```
- Every cell is converted to a trimmed string (the spec asks to trim
  whitespace), so `" Dave "` becomes `"Dave"`.
- Fully blank rows are skipped (the spec asks to ignore empty rows).

### Clear failure modes
Raises `WorkbookError` (file missing) or `EmptyWorkbookError` (no header / no
data). These are the *fatal* family, so `main.py` stops the run with a helpful
message instead of crashing.

## Known gotcha: "only N rows were read"
In `read_only=True` mode, openpyxl trusts the worksheet's **declared dimension**
stored in the file. Some exporters (SAP, web tools, certain Outlook/Excel
exports) write a too-small dimension, so `iter_rows` stops early. If you ever see
fewer rows than the file really has, the fix is to ask openpyxl to recompute the
range with `worksheet.reset_dimensions()` (or read with `read_only=False`). This
is a property of the source file, not a bug in your data.
