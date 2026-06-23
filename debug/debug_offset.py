"""
Diagnose the 10-cell horizontal offset for E7213_carcasa/L7/HVN_FL50/H/1mDia thermal.

Excel reads from AA5/AA6 via INDEX/MATCH on CoreExample col A:
  AA5 (srcX) = 156.42416666666668
  AA6 (srcY) = 98.60372423769823

We need to confirm:
  1. What Python's load_core() returns for this key
  2. Exactly how Python maps srcX -> column index vs how Excel does it
  3. Whether the 10-cell offset comes from coordinate mismatch or grid formula difference
"""
import math, numpy as np, openpyxl

QX, QY = 315, 317
SX, SY = 1.0698412698412698, 1.069425

KEY      = "E7213_carcasa/L7/HVN_FL50/H/1mDia"
# Values Excel reads from AA5/AA6 for this scenario
XL_SRC_X = 156.42416666666668
XL_SRC_Y = 98.60372423769823

# ── 1. What does Python's load_core() get? ──────────────────────────────────
wb = openpyxl.load_workbook('KernelV0_v2_copy.xlsx', data_only=True, read_only=True)
ws = wb['CoreExample']

py_x, py_y = None, None
matched_row = None
for r_idx, row in enumerate(ws.iter_rows(min_row=4, values_only=True), start=4):
    if row[0] == KEY:
        py_x, py_y = row[3], row[4]
        matched_row = r_idx
        break
wb.close()

print(f"Key: {KEY}")
print(f"\nCoreExample row matched: {matched_row}")
print(f"  Python srcX = {py_x}")
print(f"  Python srcY = {py_y}")
print(f"\nExcel  srcX (AA5) = {XL_SRC_X}")
print(f"Excel  srcY (AA6) = {XL_SRC_Y}")
print(f"\nDiff X = {(py_x or 0) - XL_SRC_X:.6f} m")
print(f"Diff Y = {(py_y or 0) - XL_SRC_Y:.6f} m")

# ── 2. What column does each srcX map to? ───────────────────────────────────
def src_to_col_python(src_x):
    # Python: cell c has centre at SX*(c-0.5), so closest col to src:
    return src_x / SX + 0.5   # fractional column (1-based)

def src_to_col_excel_candidate1(src_x):
    # Hypothesis: Excel uses SX*c (no -0.5 offset, i.e. cell right-edge convention)
    return src_x / SX         # fractional column

def src_to_col_excel_candidate2(src_x):
    # Hypothesis: Excel uses SX*(c-1) (left-edge convention)
    return src_x / SX + 1.0

print("\n── Column mapping (1-based) ──")
print(f"Python  srcX col  : {src_to_col_python(py_x):.4f}  (centre = SX*(c-0.5))")
print(f"Excel   srcX col  : {src_to_col_python(XL_SRC_X):.4f}  (same formula)")
print(f"Delta columns     : {src_to_col_python(py_x) - src_to_col_python(XL_SRC_X):.4f}")

print(f"\nPython  srcX / SX : {py_x / SX:.4f}")
print(f"Excel   srcX / SX : {XL_SRC_X / SX:.4f}")
print(f"Delta             : {(py_x - XL_SRC_X) / SX:.4f} cells")

# ── 3. Simulate where the peak (minimum distance) lands in each case ─────────
cols = np.arange(1, QX + 1)
rows = np.arange(1, QY + 1)
x_m = SX * (cols - 0.5)          # Python cell centres
y_m = SY * (QY - rows + 0.5)

# Python grid: find column of minimum distance row for each approach
dist_py = np.sqrt((x_m - py_x)**2 + (SY*(QY - 159 + 0.5) - py_y)**2)  # arbitrary row 159
peak_col_py = int(np.argmin(np.abs(x_m - py_x))) + 1

dist_xl = np.sqrt((x_m - XL_SRC_X)**2 + (SY*(QY - 159 + 0.5) - XL_SRC_Y)**2)
peak_col_xl = int(np.argmin(np.abs(x_m - XL_SRC_X))) + 1

print(f"\n── Nearest column to srcX (Python cell-centre grid) ──")
print(f"Python  srcX={py_x:.4f}  -> nearest col {peak_col_py}  (x={x_m[peak_col_py-1]:.4f} m)")
print(f"Excel   srcX={XL_SRC_X:.4f}  -> nearest col {peak_col_xl}  (x={x_m[peak_col_xl-1]:.4f} m)")
print(f"Col diff = {peak_col_py - peak_col_xl}  cells")

# ── 4. Also check if Excel formula might use SX*col instead of SX*(col-0.5) ─
x_m_xl_hyp = SX * cols   # Excel hypothesis: no -0.5
peak_col_xl_hyp = int(np.argmin(np.abs(x_m_xl_hyp - XL_SRC_X))) + 1
print(f"\nIf Excel uses SX*col (no -0.5):")
print(f"  Excel srcX={XL_SRC_X:.4f} -> nearest col {peak_col_xl_hyp}  "
      f"(x={x_m_xl_hyp[peak_col_xl_hyp-1]:.4f} m)")
print(f"  Col diff vs Python = {peak_col_py - peak_col_xl_hyp} cells")

# ── 5. Print first few rows of CoreExample around E7213 for manual check ────
print(f"\n── CoreExample rows near E7213 ──")
wb2 = openpyxl.load_workbook('KernelV0_v2_copy.xlsx', data_only=True, read_only=True)
ws2 = wb2['CoreExample']
for r_idx, row in enumerate(ws2.iter_rows(min_row=4, values_only=True), start=4):
    key_val = row[0]
    if key_val and 'E7213' in str(key_val):
        print(f"  row {r_idx}: A={row[0]}  X={row[3]}  Y={row[4]}")
wb2.close()
