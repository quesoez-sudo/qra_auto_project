"""
Read the cached kernel AB1:MD317 values from KernelV0_v2_copy.xlsx
and compare with the Python thermal CSV for E7213.

Also prints AA3/AA5/AA6 to confirm which scenario is cached.
"""
import csv, zipfile, xml.etree.ElementTree as ET
import numpy as np, openpyxl

QX, QY = 315, 317
FILE    = "KernelV0_v2_copy.xlsx"
CSV_PATH = "output_v2/thermal/E7213_carcasa_L7_HVN_FL50_H_1mDia.csv"

# ── 1. Read AA params to confirm scenario ────────────────────────────────────
wb = openpyxl.load_workbook(FILE, data_only=True)
ws = wb['ImpactThermMatrix']

def cell_val(ref):
    return ws[ref].value

print("=== AA control cells (cached values) ===")
for r, label in [(2,"mode"), (3,"scenario key"), (5,"srcX"), (6,"srcY"),
                  (7,"freq"), (8,"expTime"), (13,"SX"), (14,"SY"),
                  (15,"QX"), (16,"QY")]:
    print(f"  AA{r:2d} ({label:<12}): {cell_val(f'AA{r}')!r}")

# ── 2. Read AB1 spill range (AB1:MD317) ──────────────────────────────────────
# AB=col 28, MD=col28+315-1=342
# Check a few cells first
print(f"\n=== Sample of AB1 spill range (AB1:MD317) ===")
AA3_val = cell_val('AA3')
print(f"Scenario in AA3: {AA3_val!r}")

# Count non-None cells in a few rows
non_none_count = 0
total = 0
kernel_mat = np.zeros((QY, QX), dtype=float)
for r_idx in range(1, QY+1):
    for c_idx in range(28, 28+QX):
        col_letter = ''
        n = c_idx
        while n > 0:
            n, rem = divmod(n - 1, 26)
            col_letter = chr(ord('A') + rem) + col_letter
        cell = ws[f'{col_letter}{r_idx}']
        v = cell.value
        total += 1
        if v is not None:
            non_none_count += 1
            kernel_mat[r_idx-1, c_idx-28] = float(v) if isinstance(v, (int, float)) else 0.0

wb.close()

print(f"AB1:MD317: {total} cells,  {non_none_count} non-None,  {total-non_none_count} None (0)")
nz_k = np.count_nonzero(kernel_mat)
print(f"Non-zero kernel cells: {nz_k}")

if nz_k > 0:
    nz_rows, nz_cols = np.where(kernel_mat > 0)
    print(f"Non-zero region:")
    print(f"  Col range: {nz_cols.min()+1} to {nz_cols.max()+1}  centre {(nz_cols.min()+nz_cols.max())/2+1:.1f}")
    print(f"  Row range: {nz_rows.min()+1} to {nz_rows.max()+1}  centre {(nz_rows.min()+nz_rows.max())/2+1:.1f}")
else:
    print("Kernel matrix is all zeros -- cached values not set for this scenario.")
    print("Check if user last saved with a different scenario in AA3.")

# ── 3. Load Python CSV ────────────────────────────────────────────────────────
py_mat = []
with open(CSV_PATH, newline='') as f:
    for row in csv.reader(f):
        if row and row[0].startswith('#'):
            continue
        py_mat.append([float(v) for v in row])
py_mat = np.array(py_mat)

nz_py_rows, nz_py_cols = np.where(py_mat > 0)
print(f"\nPython non-zero region:")
print(f"  Col range: {nz_py_cols.min()+1} to {nz_py_cols.max()+1}  centre {(nz_py_cols.min()+nz_py_cols.max())/2+1:.1f}")
print(f"  Row range: {nz_py_rows.min()+1} to {nz_py_rows.max()+1}  centre {(nz_py_rows.min()+nz_py_rows.max())/2+1:.1f}")

# ── 4. Compare ────────────────────────────────────────────────────────────────
if nz_k > 0:
    diff = np.abs(kernel_mat - py_mat)
    match_rate = np.mean(diff < 1e-4)
    print(f"\nDiff (kernel vs Python):")
    print(f"  Match rate (|diff|<1e-4): {match_rate:.1%}")
    print(f"  Max abs error: {diff.max():.4e}")

    # Find column shift by cross-correlation on a middle row
    mid_r = (nz_py_rows.min() + nz_py_rows.max()) // 2
    row_py = py_mat[mid_r, :]
    row_kn = kernel_mat[mid_r, :]
    if row_py.sum() > 0 and row_kn.sum() > 0:
        corr = np.correlate(row_py, row_kn, mode='full')
        shift = np.argmax(corr) - (QX - 1)
        print(f"\n  Cross-correlation col shift at row {mid_r+1}: {shift:+d} cols")
        print(f"  (positive = Python is shifted right vs kernel)")

    # Check row 225, cols 130-165
    print(f"\n  Row 225 comparison (cols 130-165):")
    print(f"  {'Col':>5} {'Kernel':>12} {'Python':>12} {'Diff':>12}")
    for c in range(129, 165):
        kv = kernel_mat[224, c]
        pv = py_mat[224, c]
        if kv != 0 or pv != 0:
            print(f"  {c+1:>5} {kv:>12.4f} {pv:>12.4f} {pv-kv:>12.4f}")
