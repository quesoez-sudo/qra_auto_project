"""Find and print the ImpactThermMatrix AB1 formula, then compare
Python vs Excel cell-centre x/y conventions for srcX=156.424m."""
import zipfile, xml.etree.ElementTree as ET, math, numpy as np

ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
FILE = 'KernelV0_v2_copy.xlsx'

# ── 1. Find the sheet XML file for ImpactThermMatrix ────────────────────────
with zipfile.ZipFile(FILE) as z:
    # Read workbook.xml to get sheet-name -> rId mapping
    wb_xml = ET.fromstring(z.read('xl/workbook.xml'))
    sheet_map = {}   # name -> rId
    for sh in wb_xml.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
        sheet_map[sh.get('name')] = sh.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')

    # Read workbook.xml.rels to get rId -> target file
    rels_xml = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    rid_map = {}
    for rel in rels_xml.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        rid_map[rel.get('Id')] = rel.get('Target')

    therm_rid    = sheet_map['ImpactThermMatrix']
    therm_target = rid_map[therm_rid]
    if not therm_target.startswith('xl/'):
        therm_target = 'xl/' + therm_target

    print(f"ImpactThermMatrix XML: {therm_target}")

    # ── 2. Find all formulas in row 1 ────────────────────────────────────────
    xml_data = z.read(therm_target)
    root = ET.fromstring(xml_data)
    print("\n=== Row 1 formulas ===")
    found = False
    for row_el in root.findall('.//x:row', ns):
        if row_el.get('r') != '1':
            continue
        for c_el in row_el.findall('x:c', ns):
            ref  = c_el.get('r', '')
            f_el = c_el.find('x:f', ns)
            if f_el is not None:
                ftype = f_el.get('t', 'normal')
                text  = f_el.text or '(empty text - spill reference)'
                print(f"\n  Cell {ref} [type={ftype}, ref={f_el.get('ref','')}]:")
                print(f"  {text[:2000]}")
                found = True
    if not found:
        print("  (no formulas found in row 1 - formula may be elsewhere)")

# ── 3. Compare grid conventions for srcX=156.424m ────────────────────────────
QX, QY  = 315, 317
SX, SY  = 1.0698412698412698, 1.069425
SRC_X   = 156.42416666666668
SRC_Y   = 98.60372423769823

cols = np.arange(1, QX+1)
rows = np.arange(1, QY+1)

# Python convention (current engine)
x_py = SX * (cols - 0.5)          # centre: 0.5, 1.5, ..., 314.5 * SX
y_py = SY * (QY - rows + 0.5)     # centre: 316.5, 315.5, ..., 0.5 * SY

# Hypothesis A: Excel uses SX*col (right-edge, or: cell index = 1..N, centre at SX*i)
x_xl_A = SX * cols

# Hypothesis B: Excel uses SX*(col-1) (left-edge)
x_xl_B = SX * (cols - 1)

# Hypothesis C: Excel uses SX*col - SX/2 (same as Python, just written differently)
# (this would produce no offset, confirming same convention)

print("\n=== Cell-centre x for first/last cols ===")
for label, x_arr in [("Python  SX*(c-0.5)", x_py), ("Excel-A SX*c", x_xl_A), ("Excel-B SX*(c-1)", x_xl_B)]:
    col_nearest = int(np.argmin(np.abs(x_arr - SRC_X))) + 1
    x_nearest   = x_arr[col_nearest - 1]
    print(f"  {label:<22}: col1={x_arr[0]:.4f}m  col315={x_arr[-1]:.4f}m  "
          f"srcX nearest col={col_nearest}  (x={x_nearest:.4f}m, err={x_nearest-SRC_X:+.4f}m)")

# Delta between Python and each Excel hypothesis at srcX
col_py = int(np.argmin(np.abs(x_py - SRC_X))) + 1
col_A  = int(np.argmin(np.abs(x_xl_A - SRC_X))) + 1
col_B  = int(np.argmin(np.abs(x_xl_B - SRC_X))) + 1

print(f"\n  Python centre col  = {col_py}")
print(f"  Hyp-A centre col  = {col_A}  (delta = {col_py - col_A:+d})")
print(f"  Hyp-B centre col  = {col_B}  (delta = {col_py - col_B:+d})")

# What SX convention gives exactly 10 cells to the left of Python?
target_col = col_py - 10
target_x   = x_py[target_col - 1]
print(f"\n  Python is 10 cols RIGHT of Excel => Excel peak at col {target_col}")
print(f"  That cell centre (Python coords) = {target_x:.4f} m")
print(f"  If Excel formula: x=SX*col, srcX maps to col {SRC_X/SX:.4f}")
print(f"  If Excel formula: x=SX*(col-1), srcX maps to col {SRC_X/SX+1:.4f}")
print(f"  SRC_X / SX = {SRC_X/SX:.6f}  (offset from index-0 origin)")

# Check if using SX*(col) shifts by exactly 10 compared to Python
# Python peak: SRC_X / SX + 0.5
# Hyp-A peak:  SRC_X / SX
# Difference:  0.5 cells  (NOT 10)
# => Neither formula difference explains 10 cells

# What offset in X (in metres) corresponds to 10 cells?
offset_m = 10 * SX
print(f"\n  10 cells x SX = {offset_m:.4f} m")
print(f"  SRC_X + 10*SX = {SRC_X + offset_m:.4f} m  (Python source if Excel is right)")
print(f"  SRC_X - 10*SX = {SRC_X - offset_m:.4f} m  (Excel source if Python is right)")
