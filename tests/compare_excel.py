"""
tests/compare_excel.py
======================
Compares Python formula output against Excel's native computation on a
per-scenario basis.

For N sampled scenarios per event this script:
  1. Writes the scenario's parameters to the impact sheet's AK/AL control cells.
  2. Forces Excel to recalculate (xlwings COM call).
  3. Reads the output matrix from AM1:MO317.
  4. Runs Python's formula with the same inputs.
  5. Reports match_rate, max_abs_error, mean_abs_error, and non-zero cell counts.

Requirements:
  - Excel installed and xlwings:  pip install xlwings
  - KernelV0 (version 1).xlsx at _EXCEL_PATH

Usage:
  python tests/compare_excel.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import csv
import numpy as np
import openpyxl
import xlwings as xw

from qra_engine import (
    QX, QY, SX, SY,
    read_core,
    read_thermal_scenarios,
    read_jf_scenarios,
    read_toxic_scenarios,
    read_explosion_scenarios,
    read_ff_scenarios,
    formula_thermal,
    formula_jf,
    formula_toxic,
    formula_explosion,
    formula_ff,
    dist_grid,
)

# ── Configuration ─────────────────────────────────────────────────────────────
_WORKSPACE   = r'C:\Users\herman.ramirez\OneDrive - Wood PLC\ODS\QRA_cod_project'
_EXCEL_PATH  = _WORKSPACE + r'\KernelV0 (version 1).xlsx'
_N_SCENARIOS = 10      # scenarios to sample per event
_TOLERANCE   = 1e-4    # cell match threshold  |xl − py| < _TOLERANCE → match
_RANDOM_SEED = 42
_OUTPUT_DIR  = os.path.join(os.path.dirname(__file__), '..', 'output')

# Restrict to specific events (empty list = all events)
_EVENTS_FILTER = []   # e.g. ['JF', 'TOXIC']

# ── Helpers ───────────────────────────────────────────────────────────────────
def _col_num_to_letter(n):
    """1-indexed column number → Excel letter string. 39 → 'AM'."""
    letters = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord('A') + rem) + letters
    return letters

_END_COL = _col_num_to_letter(39 + QX - 1)        # 'MO' for QX=315
_MATRIX_RANGE = f'AM1:{_END_COL}{QY}'             # 'AM1:MO317'


def _extract_weather(key):
    """Return 'Noche' or 'Dia' from a ScenarioWeather key string."""
    if key.endswith('Noche'):
        return 'Noche'
    if key.endswith('Dia'):
        return 'Dia'
    # Fallback: last path segment
    return key.rsplit('/', 1)[-1]


def _read_ak_labels(ws_xl):
    """
    Scan AK1:AK50 for non-None labels.
    Returns {label_string: row_number} (1-indexed, matching the sheet row).
    """
    vals = ws_xl.range('AK1:AK50').value
    # xlwings returns a flat list for a single-column range
    if vals is None:
        return {}
    label_row = {}
    for i, v in enumerate(vals, start=1):
        if v is not None:
            label_row[str(v).strip()] = i
    return label_row


def _set_scenario(ws_xl, label_row, key, weather, x, y):
    """
    Write scenario parameters to AK/AL cells.
    Sets Impact/Risk=0 so Excel computes raw impact (not impact × probability).
    """
    def _set(label, val):
        r = label_row.get(label)
        if r is not None:
            ws_xl.range(f'AL{r}').value = val

    _set('Impact/Risk', 0)
    _set('Scenario',    key)
    _set('Weather',     weather)
    _set('X',           x)
    _set('Y',           y)
    _set('Probability', 0.0)


def _read_matrix(ws_xl):
    """
    Read AM1:MO317 from Excel into a (QY, QX) float64 numpy array.
    None cells (formula errors, blanks) are treated as 0.0.
    """
    raw = ws_xl.range(_MATRIX_RANGE).value
    if raw is None:
        return np.zeros((QY, QX))
    arr = np.zeros((QY, QX), dtype=float)
    for i, row in enumerate(raw):
        if row is None:
            continue
        for j, v in enumerate(row):
            arr[i, j] = float(v) if v is not None else 0.0
    return arr


def _check_grid_params(ws_xl, label_row):
    """
    Read QX, QY, SX, SY from AK/AL and compare to Python module constants.
    Returns True if all match within tolerance.
    """
    def _al(label):
        r = label_row.get(label)
        return ws_xl.range(f'AL{r}').value if r else None

    xl_qx = _al('QX')
    xl_qy = _al('QY')
    xl_sx = _al('SX')
    xl_sy = _al('SY')

    ok = True
    for name, xl_val, py_val, check in [
        ('QX', xl_qx, QX, lambda a, b: a is not None and int(a) == b),
        ('QY', xl_qy, QY, lambda a, b: a is not None and int(a) == b),
        ('SX', xl_sx, SX, lambda a, b: a is not None and abs(a - b) < 1e-9),
        ('SY', xl_sy, SY, lambda a, b: a is not None and abs(a - b) < 1e-9),
    ]:
        match = check(xl_val, py_val)
        flag = '✓' if match else '✗'
        print(f'    {flag} {name}: Excel={xl_val}  Python={py_val}')
        if not match:
            ok = False
    return ok


def _metrics(xl_mat, py_mat):
    """
    Compute quality metrics between two (QY, QX) matrices.
    Returns (match_rate, max_abs_err, mean_abs_err, nonzero_xl, nonzero_py).
    """
    diff = np.abs(xl_mat - py_mat)
    match_rate = float(np.mean(diff < _TOLERANCE))
    max_err    = float(np.max(diff))
    nz_mask    = (xl_mat != 0) | (py_mat != 0)
    mean_err   = float(np.mean(diff[nz_mask])) if nz_mask.any() else 0.0
    nz_xl      = int(np.count_nonzero(xl_mat))
    nz_py      = int(np.count_nonzero(py_mat))
    return match_rate, max_err, mean_err, nz_xl, nz_py


# ── Python formula wrappers (one per event type) ──────────────────────────────

def _py_thermal(sc):
    return formula_thermal(dist_grid(sc['sx'], sc['sy']), sc['therm_dists'])

def _py_jf(sc):
    return formula_jf(sc['sx'], sc['sy'],
                      sc['dist_vals'], sc['halfW_vals'], sc['center_vals'])

def _py_toxic(sc):
    return formula_toxic(dist_grid(sc['sx'], sc['sy']),
                         sc['tox_dists'], sc['tox_probs'])

def _py_explosion_cve(sc):
    sx = sc['ign_x'] if sc.get('ign_x') is not None else sc['sx']
    sy = sc['ign_y'] if sc.get('ign_y') is not None else sc['sy']
    return formula_explosion(dist_grid(sx, sy), sc['exp_dists'])

def _py_explosion_blv(sc):
    return formula_explosion(dist_grid(sc['sx'], sc['sy']), sc['exp_dists'])

def _py_ff(sc):
    return formula_ff(dist_grid(sc['sx'], sc['sy']),
                      sc['lfl_dist'], sc['lflf_dist'])


# ── Event catalogue ───────────────────────────────────────────────────────────
# (event_name, sheet_name, reader_fn, python_formula_fn, x_fn, y_fn)
# x_fn / y_fn: lambda sc → the coordinate to set in AK/AL X and Y cells
#   (for CVE this is the ignition point, not the source equipment)

_CATALOGUE = [
    ('TOXIC', 'ImpactToxMatrix',  read_toxic_scenarios,     _py_toxic,
     lambda sc: sc['sx'], lambda sc: sc['sy']),
    ('JF',    'ImpactJFMatrix',   read_jf_scenarios,        _py_jf,
     lambda sc: sc['sx'], lambda sc: sc['sy']),
    ('LPF',   'ImpactThermMatrix',read_thermal_scenarios,   _py_thermal,
     lambda sc: sc['sx'], lambda sc: sc['sy']),
    ('CVE',   'ImpactExpMatrix',  read_explosion_scenarios, _py_explosion_cve,
     lambda sc: sc['ign_x'] if sc.get('ign_x') else sc['sx'],
     lambda sc: sc['ign_y'] if sc.get('ign_y') else sc['sy']),
    ('FF',    'ImpactFFMatrix',   read_ff_scenarios,        _py_ff,
     lambda sc: sc['sx'], lambda sc: sc['sy']),
]


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    os.makedirs(_OUTPUT_DIR, exist_ok=True)
    rng = np.random.default_rng(_RANDOM_SEED)

    # Filter catalogue
    catalogue = [e for e in _CATALOGUE
                 if not _EVENTS_FILTER or e[0] in _EVENTS_FILTER]

    # ── Phase 1: load scenarios via openpyxl (no Excel needed) ────────────────
    print('Phase 1 — loading scenarios from workbook (openpyxl)...')
    wb_opx = openpyxl.load_workbook(_EXCEL_PATH, data_only=True)
    core = read_core(wb_opx)
    print(f'  Core: {len(core)} scenarios')

    scenario_cache = {}
    event_scenarios = {}
    for ev_name, sheet, reader_fn, _, _, _ in catalogue:
        if sheet not in scenario_cache:
            scenario_cache[sheet] = reader_fn(wb_opx, core)
        event_scenarios[ev_name] = scenario_cache[sheet]
        print(f'  {ev_name:<8}: {len(event_scenarios[ev_name]):4d} scenarios  ({sheet})')
    wb_opx.close()

    # ── Phase 2: xlwings comparison ───────────────────────────────────────────
    print(f'\nPhase 2 — opening workbook with xlwings (Excel required)')
    print(f'  File:         {_EXCEL_PATH}')
    print(f'  Matrix range: {_MATRIX_RANGE}  ({QX} cols × {QY} rows)')
    print(f'  Scenarios:    {_N_SCENARIOS} per event  (tolerance={_TOLERANCE})')

    all_rows = []

    app = xw.App(visible=False)
    try:
        wb_xl = app.books.open(_EXCEL_PATH)

        for ev_name, sheet, _, py_fn, x_fn, y_fn in catalogue:
            scenarios = event_scenarios[ev_name]
            if not scenarios:
                print(f'\n{ev_name}: 0 scenarios — skipped')
                continue

            ws_xl = wb_xl.sheets[sheet]
            label_row = _read_ak_labels(ws_xl)

            print(f'\n── {ev_name}  ({sheet}) ──')
            print('  Grid params:')
            _check_grid_params(ws_xl, label_row)

            # Sample N scenarios
            n = min(_N_SCENARIOS, len(scenarios))
            idx = rng.choice(len(scenarios), n, replace=False)
            sample = [scenarios[i] for i in sorted(idx)]

            hdr = f'  {"Scenario key":<44} {"match%":>7} {"maxErr":>10} {"meanErr":>10} {"nz_XL":>7} {"nz_Py":>7}'
            print(hdr)
            print('  ' + '-' * (len(hdr) - 2))

            for sc in sample:
                weather = _extract_weather(sc['key'])
                x = x_fn(sc)
                y = y_fn(sc)

                _set_scenario(ws_xl, label_row, sc['key'], weather, x, y)
                app.calculate()

                xl_mat = _read_matrix(ws_xl)
                py_mat = py_fn(sc)

                mr, max_e, mean_e, nz_xl, nz_py = _metrics(xl_mat, py_mat)

                key_short = sc['key'][:44]
                print(f'  {key_short:<44} {mr:>7.1%} {max_e:>10.3e} {mean_e:>10.3e} {nz_xl:>7d} {nz_py:>7d}')

                all_rows.append({
                    'event':          ev_name,
                    'sheet':          sheet,
                    'scenario':       sc['key'],
                    'weather':        weather,
                    'x':              x,
                    'y':              y,
                    'match_rate':     f'{mr:.6f}',
                    'max_abs_error':  f'{max_e:.6e}',
                    'mean_abs_error': f'{mean_e:.6e}',
                    'nonzero_xl':     nz_xl,
                    'nonzero_py':     nz_py,
                })

        wb_xl.close()
    finally:
        app.quit()

    # ── Save results ──────────────────────────────────────────────────────────
    if all_rows:
        out_path = os.path.join(_OUTPUT_DIR, 'compare_excel_results.csv')
        with open(out_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=list(all_rows[0].keys()))
            writer.writeheader()
            writer.writerows(all_rows)
        print(f'\nResults saved → {out_path}')

    # ── Summary ───────────────────────────────────────────────────────────────
    if all_rows:
        print('\n── Summary ──')
        by_event = {}
        for r in all_rows:
            by_event.setdefault(r['event'], []).append(r)
        for ev, rows in by_event.items():
            avg_mr   = np.mean([float(r['match_rate']) for r in rows])
            avg_maxe = np.mean([float(r['max_abs_error']) for r in rows])
            print(f'  {ev:<8}  avg_match={avg_mr:.1%}  avg_maxErr={avg_maxe:.3e}  n={len(rows)}')


if __name__ == '__main__':
    main()
