"""
Find the centre of the non-zero region in Python's thermal CSV
for E7213_carcasa/L7/HVN_FL50/H/1mDia, then compare with
what the grid formula predicts for the same srcX/srcY.
"""
import csv, numpy as np

QX, QY = 315, 317
SX, SY = 1.0698412698412698, 1.069425
SRC_X  = 156.42416666666668
SRC_Y  = 98.60372423769823

# ── Load the thermal impact CSV ───────────────────────────────────────────────
CSV_PATH = "output_v2/thermal/E7213_carcasa_L7_HVN_FL50_H_1mDia.csv"
mat = []
with open(CSV_PATH, newline='') as f:
    for row in csv.reader(f):
        if row and row[0].startswith('#'):
            continue
        mat.append([float(v) for v in row])
mat = np.array(mat)
print(f"Matrix shape: {mat.shape}  (expected {QY}x{QX})")

# ── Where is the peak? ────────────────────────────────────────────────────────
peak_idx  = np.unravel_index(np.argmax(mat), mat.shape)
peak_row  = peak_idx[0] + 1   # 1-based
peak_col  = peak_idx[1] + 1   # 1-based
peak_val  = mat[peak_idx]

# Also find bounding box of non-zero cells
nz_rows, nz_cols = np.where(mat > 0)
if len(nz_cols):
    print(f"\nNon-zero region:")
    print(f"  Col range : {nz_cols.min()+1} to {nz_cols.max()+1}  (centre {(nz_cols.min()+nz_cols.max())/2+1:.1f})")
    print(f"  Row range : {nz_rows.min()+1} to {nz_rows.max()+1}  (centre {(nz_rows.min()+nz_rows.max())/2+1:.1f})")
else:
    print("All zeros!")

print(f"\nPeak value: {peak_val:.4f} kW/m2 at (row={peak_row}, col={peak_col})")

# ── Where does srcX/srcY map to in the Python grid? ──────────────────────────
cols = np.arange(1, QX+1)
rows = np.arange(1, QY+1)
x_m  = SX * (cols - 0.5)
y_m  = SY * (QY - rows + 0.5)

src_col_exact = SRC_X / SX + 0.5        # exact (fractional)
src_row_exact = QY - SRC_Y / SY + 0.5  # exact (fractional)

print(f"\n── Expected source position in Python grid ──")
print(f"  srcX = {SRC_X:.6f} m  -> col {src_col_exact:.4f}  (nearest col {round(src_col_exact)})")
print(f"  srcY = {SRC_Y:.6f} m  -> row {src_row_exact:.4f}  (nearest row {round(src_row_exact)})")
print(f"  x at nearest col {round(src_col_exact)}: {x_m[round(src_col_exact)-1]:.6f} m")
print(f"  y at nearest row {round(src_row_exact)}: {y_m[round(src_row_exact)-1]:.6f} m")

print(f"\n── Comparison ──")
print(f"  Python CSV non-zero centre col : {(nz_cols.min()+nz_cols.max())/2+1:.1f}")
print(f"  Expected src col               : {src_col_exact:.1f}")
print(f"  Difference                     : {(nz_cols.min()+nz_cols.max())/2+1 - src_col_exact:.1f} cols")

# ── What would the centre be if Excel starts at AB (col 28)? ─────────────────
# If user sees Excel output starting at AB1, their "col 1" = spreadsheet col 28
# and Excel's impact-centre col (in spreadsheet) should also be ~src_col_exact
# So in Excel's AB1:MD317 output, the k-th column = spreadsheet col 27+k
# Python's k-th CSV column = Python col k
# Both k should be the same (~147) since same formula convention

print(f"\n── If user is comparing AB1:MD317 (Excel) vs Python CSV ──")
print(f"  The 147th column of both outputs should have the same x position")
print(f"  Excel spill col 147 = Excel spreadsheet col {28+147-1} = col {28+146}")
print(f"  Python CSV col 147")
print(f"  Both x = SX*(147-0.5) = {SX*146.5:.4f} m  (vs srcX={SRC_X:.4f} m)")

# ── Print a small window around the peak ──────────────────────────────────────
print(f"\n── Matrix slice near peak (rows {max(1,peak_row-2)}-{min(QY,peak_row+2)}, "
      f"cols {max(1,peak_col-5)}-{min(QX,peak_col+5)}) ──")
r0 = max(0, peak_row-3); r1 = min(QY, peak_row+2)
c0 = max(0, peak_col-6); c1 = min(QX, peak_col+5)
print(f"  cols: {c0+1} .. {c1}")
for r in range(r0, r1):
    vals = [f"{mat[r,c]:5.1f}" for c in range(c0, c1)]
    print(f"  row {r+1:3d}: {' '.join(vals)}")
