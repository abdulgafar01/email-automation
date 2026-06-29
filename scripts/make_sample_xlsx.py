"""Generate a sample contacts workbook for testing the Excel reader.

Run: python scripts/make_sample_xlsx.py
"""

from pathlib import Path

from openpyxl import Workbook

DATA = Path("data")
DATA.mkdir(exist_ok=True)

wb = Workbook()
ws = wb.active
ws.title = "Contacts"

# Header row (col A is the recipient; subject = first 7 columns).
ws.append([
    "Account", "Reference", "Contract", "Name",
    "ID", "Status", "Amount", "Notes",
])
ws.append([
    "80409", "121014959", "KOC-10048 -25/26",
    "YAQOUB YOUSEF MOHAMMAD S. ALKANDARI",
    "313111600751", "SUSPENDED", "110", "first note",
])
ws.append([
    "80410", "121014960", "KOC-10049 -25/26",
    "SAMPLE CONTACT TWO",
    "313111600752", "ACTIVE", "250", "second note",
])
ws.append([None, None, None, None, None, None, None, None])  # empty -> ignored
ws.append([
    " 80411 ", " 121014961 ", " KOC-10050 -25/26 ",
    " THIRD CONTACT ", " 313111600753 ", " PENDING ", " 90 ", " <b>html</b> ",
])

out = DATA / "contacts.xlsx"
wb.save(out)
print(f"Wrote {out}")
