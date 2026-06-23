"""Find what column ANCHOR and ANCHOR_F named ranges point to."""
import zipfile, xml.etree.ElementTree as ET

FILE = 'KernelV0_v2_copy.xlsx'

# ── Read named ranges from workbook.xml ──────────────────────────────────────
with zipfile.ZipFile(FILE) as z:
    wb_xml = ET.fromstring(z.read('xl/workbook.xml'))

ns_main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
defined_names = wb_xml.find(f'{{{ns_main}}}definedNames')
if defined_names is None:
    print("No definedNames element found in workbook.xml")
else:
    for dn in defined_names.findall(f'{{{ns_main}}}definedName'):
        name = dn.get('name')
        val  = dn.text or ''
        if 'ANCHOR' in name.upper() or 'anchor' in name.lower():
            print(f"  Named range: {name!r:30s} = {val}")

# ── Also read the ANCHOR sheet to see what cells are populated ───────────────
import openpyxl
wb = openpyxl.load_workbook(FILE, data_only=True)
ws = wb['ANCHOR']
print(f"\nANCHOR sheet dimensions: {ws.dimensions}")
print("Non-empty cells in ANCHOR sheet:")
for row in ws.iter_rows():
    for cell in row:
        if cell.value is not None:
            print(f"  {cell.coordinate} (col={cell.column}, row={cell.row}): {repr(cell.value)}")
wb.close()

# ── Check all defined names ──────────────────────────────────────────────────
print("\nAll defined names:")
with zipfile.ZipFile(FILE) as z:
    wb_xml = ET.fromstring(z.read('xl/workbook.xml'))
defined_names = wb_xml.find(f'{{{ns_main}}}definedNames')
if defined_names is not None:
    for dn in defined_names.findall(f'{{{ns_main}}}definedName'):
        name = dn.get('name')
        val  = dn.text or ''
        print(f"  {name!r:40s} = {val}")
