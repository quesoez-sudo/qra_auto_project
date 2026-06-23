"""
compare_thermal_all.py
======================
Modular kernel-vs-Python comparison for all QRA event types.

For each scenario the script:
  1. Sets the kernel's scenario key cell to the active key (xlwings)
  2. Forces recalculation, reads the output matrix from the kernel sheet
  3. Loads the matching Python CSV from output_v2/<event>/
  4. Reports per scenario:
       NZ_COL_START / NZ_COL_END / NZ_WIDTH   (X-axis non-zero extent)
       NZ_ROW_START / NZ_ROW_END / NZ_HEIGHT  (Y-axis, row 1 = north)
       NZ_COUNT     -- total non-zero cells
       MATCH_RATE   -- fraction of cells where |Excel - Python| < TOLERANCE
       MAX_ABS_ERR  -- worst absolute error across all cells
       MEAN_ABS_ERR -- mean absolute error inside the non-zero union region
       SRC_X / SRC_Y -- confirmed from kernel after recalculation

Toggle events on/off by editing ACTIVE_EVENTS inside main().

Usage:
  python tests/compare_thermal_all.py

Requires:
  - xlwings + Excel installed
  - KernelV0_v2_copy.xlsx at the project root
  - output_v2/<event>/ folders populated by qra_engine_v2.py
"""

import csv
import time
from pathlib import Path
import numpy as np
import openpyxl
import xlwings as xw

# ── Paths ─────────────────────────────────────────────────────────────────────
_ROOT       = Path(__file__).resolve().parent.parent
KERNEL_FILE = str(_ROOT / "KernelV0_v2_copy.xlsx")
OUTPUT_V2   = _ROOT / "output_v2"

# ── Grid constants ─────────────────────────────────────────────────────────────
QX, QY    = 315, 317
TOLERANCE = 1e-4

# ── Event catalogue ────────────────────────────────────────────────────────────
# Each entry describes one event type: which kernel sheet drives it, which cell
# holds the scenario key, where the output matrix starts, and which Python
# sub-folder holds the generated CSVs.
#
# mat_col : 1-based Excel column number where the output matrix begins
#   AB = 28  (thermal / ff / toxic)
#   AL = 38  (jf — control area occupies AJ:AK, matrix starts at AL)
EVENT_CONFIGS = {
    'thermal': {
        'display':    'Thermal (LPF)',
        'sheet':      'ImpactThermMatrix',
        'key_cell':   'AA3',
        'src_x_cell': 'AA5',
        'src_y_cell': 'AA6',
        'mat_col':    28,        # column AB
        'py_subdir':  'thermal',
    },
    'ff': {
        'display':    'Flash Fire (FF)',
        'sheet':      'ImpactFFMatrix',
        'key_cell':   'AA3',
        'src_x_cell': 'AA5',
        'src_y_cell': 'AA6',
        'mat_col':    28,
        'py_subdir':  'ff',
    },
    'toxic': {
        'display':    'Toxic (TOX)',
        'sheet':      'ImpactToxMatrix',
        'key_cell':   'AA3',
        'src_x_cell': 'AA5',
        'src_y_cell': 'AA6',
        'mat_col':    28,
        'py_subdir':  'toxic',
    },
    'jf': {
        'display':    'Jet Fire (JF)',
        'sheet':      'ImpactJFMatrix',
        'key_cell':   'AK3',
        'src_x_cell': 'AK5',
        'src_y_cell': 'AK6',
        'mat_col':    38,        # column AL
        'py_subdir':  'jf',
    },
}

# ── Helpers ────────────────────────────────────────────────────────────────────

def _col_letter(n):
    """1-based column number -> Excel letter string (e.g. 28 -> 'AB')."""
    letters = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord('A') + rem) + letters
    return letters


def _mat_range(mat_col):
    """Build 'AB1:MD317'-style range for a given start column number."""
    start = _col_letter(mat_col)
    end   = _col_letter(mat_col + QX - 1)
    return f'{start}1:{end}{QY}'


def _safe_name(s):
    return (s.replace('/', '_').replace('\\', '_')
             .replace(' ', '').replace('(', '').replace(')', ''))


def _load_event_keys(sheet_name):
    """Scan col B (rows 2+) of a kernel sheet and return unique, ordered keys."""
    wb   = openpyxl.load_workbook(KERNEL_FILE, data_only=True, read_only=True)
    ws   = wb[sheet_name]
    keys, seen = [], set()
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=2, values_only=True):
        v = row[0]
        if v and str(v).strip() and v not in seen:
            seen.add(v)
            keys.append(str(v).strip())
    wb.close()
    return keys


def _load_py_csv(path):
    mat = []
    with open(path, newline='') as f:
        for row in csv.reader(f):
            if row and row[0].startswith('#'):
                continue
            mat.append([float(v) for v in row])
    return np.array(mat)


def _nz_metrics(mat):
    """
    Non-zero region statistics (threshold: value > 0).
    Returns (col_start, col_end, width, row_start, row_end, height, count).
    All indices are 1-based; row 1 = northmost row.
    """
    nz_rows, nz_cols = np.where(mat > 0)
    if len(nz_cols) == 0:
        return None, None, 0, None, None, 0, 0
    cs = int(nz_cols.min()) + 1
    ce = int(nz_cols.max()) + 1
    rs = int(nz_rows.min()) + 1
    re = int(nz_rows.max()) + 1
    return cs, ce, ce - cs + 1, rs, re, re - rs + 1, int(len(nz_cols))


def _diff_label(a, b):
    if a is None or b is None:
        return 'N/A'
    return f'{b - a:+d}'


def _read_kernel_matrix(ws_xw, mat_col):
    """
    Read the QY x QX output matrix starting at mat_col from an xlwings sheet.
    IFERROR fallback values (<= 0) are treated as 0.
    """
    rng = ws_xw.range(_mat_range(mat_col))
    raw = rng.value
    mat = np.zeros((QY, QX), dtype=float)
    if raw is None:
        return mat
    for r, row_data in enumerate(raw):
        if row_data is None:
            continue
        for c, v in enumerate(row_data):
            if v is not None:
                fv = float(v)
                mat[r, c] = fv if fv > 0 else 0.0
    return mat


def _print_event_header(display, sheet, mat_range_str, n_scenarios):
    print(f"\n{'=' * 70}")
    print(f"  {display}")
    print(f"  Sheet: {sheet}   Matrix: {mat_range_str}   Scenarios: {n_scenarios}")
    print(f"{'=' * 70}")
    col_w = 46
    print(f"  {'Scenario':<{col_w}} {'NZcS':>5} {'NZcE':>5} {'W':>4} "
          f"{'NZrS':>5} {'NZrE':>5} {'H':>4} "
          f"{'NZcnt':>6} {'Match%':>7} {'MaxErr':>10} {'MeanErr':>10}")
    print("  " + "-" * (col_w + 5+5+4+5+5+4+6+7+10+10 + 12))


def _print_event_row(key, k_metrics, p_metrics, match_r, max_err, mean_err, elapsed):
    col_w = 46
    k_cs, k_ce, k_w, k_rs, k_re, k_h, k_nz = k_metrics
    p_cs, p_ce, p_w, p_rs, p_re, p_h, p_nz = p_metrics
    print(f"  {key[:col_w]:<{col_w}} "
          f"{k_cs or 0:>5} {k_ce or 0:>5} {k_w:>4} "
          f"{k_rs or 0:>5} {k_re or 0:>5} {k_h:>4} "
          f"{k_nz:>6} {match_r:>7.1%} {max_err:>10.2e} {mean_err:>10.2e}"
          f"  [{elapsed:.1f}s]")
    # Flag mismatches
    if (k_cs != p_cs or k_ce != p_ce or k_rs != p_rs or k_re != p_re
            or match_r < 0.999):
        print(f"    !! MISMATCH  col={_diff_label(k_cs,p_cs)}/{_diff_label(k_ce,p_ce)}"
              f"  row={_diff_label(k_rs,p_rs)}/{_diff_label(k_re,p_re)}"
              f"  Py_nz={p_nz}")


def _build_result_row(key, src_x, src_y, k_m, p_m, match_r, max_err, mean_err, elapsed):
    k_cs, k_ce, k_w, k_rs, k_re, k_h, k_nz = k_m
    p_cs, p_ce, p_w, p_rs, p_re, p_h, p_nz = p_m
    return {
        'scenario':       key,
        'knl_srcX':       f'{src_x:.6f}' if src_x else '',
        'knl_srcY':       f'{src_y:.6f}' if src_y else '',
        'k_nz_col_start': k_cs, 'k_nz_col_end': k_ce, 'k_nz_width':  k_w,
        'k_nz_row_start': k_rs, 'k_nz_row_end': k_re, 'k_nz_height': k_h,
        'k_nz_count':     k_nz,
        'p_nz_col_start': p_cs, 'p_nz_col_end': p_ce, 'p_nz_width':  p_w,
        'p_nz_row_start': p_rs, 'p_nz_row_end': p_re, 'p_nz_height': p_h,
        'p_nz_count':     p_nz,
        'dcol_start':     (p_cs - k_cs) if (p_cs and k_cs) else '',
        'dcol_end':       (p_ce - k_ce) if (p_ce and k_ce) else '',
        'drow_start':     (p_rs - k_rs) if (p_rs and k_rs) else '',
        'drow_end':       (p_re - k_re) if (p_re and k_re) else '',
        'match_rate':     f'{match_r:.6f}',
        'max_abs_error':  f'{max_err:.4e}',
        'mean_abs_error': f'{mean_err:.4e}',
        'elapsed_s':      f'{elapsed:.2f}',
    }


def _print_event_summary(name, results):
    if not results:
        print(f"  (no results for {name})")
        return
    match_rates = [float(r['match_rate']) for r in results]
    max_errs    = [float(r['max_abs_error']) for r in results]
    col_diffs   = [r['dcol_start'] for r in results if isinstance(r['dcol_start'], int)]
    row_diffs   = [r['drow_start'] for r in results if isinstance(r['drow_start'], int)]
    perfect     = sum(mr >= 1 - 1e-9 for mr in match_rates)
    mismatches  = [r['scenario'] for r in results
                   if float(r['match_rate']) < 0.999
                   or (isinstance(r['dcol_start'], int) and r['dcol_start'] != 0)
                   or (isinstance(r['drow_start'], int) and r['drow_start'] != 0)]

    print(f"\n  {name.upper()}")
    print(f"    Scenarios   : {len(results)}")
    print(f"    Avg match   : {np.mean(match_rates):.2%}")
    print(f"    100% match  : {perfect}/{len(results)}")
    print(f"    Max error   : {max(max_errs):.4e}")
    if col_diffs:
        print(f"    Col drift   : mean={np.mean(col_diffs):+.2f}  "
              f"range [{min(col_diffs)}, {max(col_diffs)}]")
    if row_diffs:
        print(f"    Row drift   : mean={np.mean(row_diffs):+.2f}  "
              f"range [{min(row_diffs)}, {max(row_diffs)}]")
    if mismatches:
        print(f"    Mismatches ({len(mismatches)}):")
        for s in mismatches:
            print(f"      - {s}")
    else:
        print(f"    Result      : ALL PASS — location and values match exactly")


# ── Per-event runner (called once per event while workbook is open) ────────────

def _match_event_keys(cfg):
    """
    Pre-load scenario keys (openpyxl) and pair each with its Python CSV path.
    Call this BEFORE opening the workbook with xlwings to avoid file-lock conflicts.
    Returns (matched [(key, Path)], skipped [key]).
    """
    py_dir   = OUTPUT_V2 / cfg['py_subdir']
    all_keys = _load_event_keys(cfg['sheet'])
    matched, skipped = [], []
    for key in all_keys:
        p = py_dir / f"{_safe_name(key)}.csv"
        (matched if p.exists() else skipped).append(
            (key, p) if p.exists() else key
        )
    return matched, skipped


def run_event(app, wb, name, cfg, matched, skipped):
    """
    Run comparison for one event type.
    matched : list of (key, csv_path) — pre-loaded before xlwings opened the file
    Returns list of result dicts (one per scenario).
    """
    sheet      = cfg['sheet']
    key_cell   = cfg['key_cell']
    src_x_cell = cfg['src_x_cell']
    src_y_cell = cfg['src_y_cell']
    mat_col    = cfg['mat_col']
    mrange     = _mat_range(mat_col)

    _print_event_header(cfg['display'], sheet, mrange, len(matched))
    if skipped:
        print(f"  (skipped {len(skipped)} keys with no Python CSV)")

    if not matched:
        py_dir = OUTPUT_V2 / cfg['py_subdir']
        print(f"  No Python CSVs found in {py_dir} — run qra_engine_v2.py first.")
        return []

    ws      = wb.sheets[sheet]
    results = []
    t_start = time.time()

    for key, csv_path in matched:
        t0 = time.time()

        ws.range(key_cell).value = key
        app.calculate()

        src_x = ws.range(src_x_cell).value
        src_y = ws.range(src_y_cell).value

        knl_mat = _read_kernel_matrix(ws, mat_col)
        py_mat  = _load_py_csv(csv_path)

        k_m = _nz_metrics(knl_mat)
        p_m = _nz_metrics(py_mat)

        diff     = np.abs(knl_mat - py_mat)
        match_r  = float(np.mean(diff < TOLERANCE))
        max_err  = float(np.max(diff))
        nz_mask  = (knl_mat > 0) | (py_mat > 0)
        mean_err = float(np.mean(diff[nz_mask])) if nz_mask.any() else 0.0
        elapsed  = time.time() - t0

        _print_event_row(key, k_m, p_m, match_r, max_err, mean_err, elapsed)
        results.append(_build_result_row(
            key, src_x, src_y, k_m, p_m, match_r, max_err, mean_err, elapsed))

    total = time.time() - t_start
    print(f"\n  Done in {total:.1f}s  ({total/len(matched):.1f}s/scenario)")
    return results


def _save_results(name, results):
    if not results:
        return
    path = OUTPUT_V2 / f"{name}_comparison_results.csv"
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=list(results[0].keys()))
        writer.writeheader()
        writer.writerows(results)
    print(f"  Saved -> {path}")


# ── Main ───────────────────────────────────────────────────────────────────────

def main():
    # ── Toggle events here ───────────────────────────────────────────────────
    # Remove an event name to skip it, add it back to run it.
    ACTIVE_EVENTS = [
        'thermal',
        'ff',
        'toxic',
        'jf',
    ]
    # ────────────────────────────────────────────────────────────────────────

    unknown = [e for e in ACTIVE_EVENTS if e not in EVENT_CONFIGS]
    if unknown:
        print(f"Unknown event(s): {unknown}. Valid: {list(EVENT_CONFIGS)}")
        return

    print("=" * 70)
    print("QRA Kernel vs Python Comparison")
    print(f"Events : {', '.join(ACTIVE_EVENTS)}")
    print(f"Kernel : {KERNEL_FILE}")
    print(f"Tol    : {TOLERANCE}")
    print("=" * 70)

    # Pre-load scenario keys with openpyxl BEFORE xlwings locks the file
    print("\nPre-loading scenario keys...")
    preloaded = {}
    for name in ACTIVE_EVENTS:
        cfg     = EVENT_CONFIGS[name]
        matched, skipped = _match_event_keys(cfg)
        preloaded[name] = (matched, skipped)
        print(f"  {cfg['display']:<22}: {len(matched)} matched, {len(skipped)} skipped")

    all_results = {}

    app = xw.App(visible=False)
    try:
        wb = app.books.open(KERNEL_FILE)

        for name in ACTIVE_EVENTS:
            cfg = EVENT_CONFIGS[name]
            matched, skipped = preloaded[name]
            results = run_event(app, wb, name, cfg, matched, skipped)
            all_results[name] = results
            _save_results(name, results)

        wb.close()
    finally:
        app.quit()

    # ── Overall summary ───────────────────────────────────────────────────────
    print("\n" + "=" * 70)
    print("SUMMARY — all active events")
    print("=" * 70)
    for name in ACTIVE_EVENTS:
        _print_event_summary(name, all_results.get(name, []))


if __name__ == '__main__':
    main()
