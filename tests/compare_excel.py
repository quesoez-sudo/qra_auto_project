"""
tests/compare_excel.py
======================
Compares qra_v6_engine formula output against Excel's KernelV0 native
computation on a per-scenario basis.  All grid parameters and thresholds
are loaded dynamically from the MacroQRAV6 General sheet — no hardcoded
QRA constants in this script.

Flow
----
Phase 0 — MacroQRAV6 (xlwings):
  Open workbook, call load_general_params() to populate qra_v6_engine._P.
  This initialises the XX/YY coordinate grids and all threshold arrays.
  qra_engine reader-function constants are then overwritten with the same
  values so the two modules stay consistent.

Phase 1a — MacroQRAV6 (xlwings):
  For each event, set Core!C2 = impact_id, recalculate, read Core A4:R89.
  Produces a core dict {key: {x, y, size, probs}} per event.

Phase 1b — KernelV0 (openpyxl):
  For each event call reader_fn(wb_kernel, core) from qra_engine.
  Reads consequence distances from ImpactXxxMatrix sheets matched by key.

Phase 2 — KernelV0 (xlwings):
  For each sampled scenario:
    - Set control cells (AK/AL or Z/AA): key, weather, x, y
    - Force recalculate, read output matrix from KernelV0 -> xl_mat
    - Run qra_v6_engine formula with same x/y + distances from 1b -> py_mat
    - Report match_rate, max_abs_error, mean_abs_error, non-zero counts

Requirements:
  - Excel installed + xlwings:  pip install xlwings
  - MacroQRAV6 (version 1).xlsm at _MACRO_PATH
  - KernelV0 (version 1) JetFireFormula.xlsx at _EXCEL_PATH

Usage:
  python tests/compare_excel.py
"""
import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import csv
import time
import numpy as np
import openpyxl
import xlwings as xw
from dotenv import load_dotenv
load_dotenv()

# ── qra_engine: reader functions only (KernelV0 openpyxl) ────────────────────
# Grid constants and formula functions are NOT imported from here.
# After Phase 0 loads _P, we patch the relevant qra_engine module attributes so
# that all threshold lookups and min-prob filters use the General sheet values.
import qra_engine as _qra_engine
from qra_engine import (
    read_thermal_scenarios,
    read_jf_scenarios,
    read_toxic_scenarios,
    read_explosion_scenarios,
    read_ff_scenarios,
    read_wind_direction,
)

# ── qra_v6_engine: formulas + dynamic parameter loader ───────────────────────
# All formula functions reference _P internally (populated by load_general_params).
# We import the module as _v6 so helper functions can access _v6._P after loading.
import qra_v6_engine as _v6
from qra_v6_engine import (
    formula_thermal,
    formula_jf,
    formula_toxic,
    formula_explosion,
    formula_ff,
    dist_grid,
    load_general_params,
    read_core_scenarios,
    IMPACT_CONFIG,
)

# ── Configuration ─────────────────────────────────────────────────────────────
_WORKSPACE   = os.getenv('WORKSPACE', os.getcwd())
_EXCEL_PATH  = _WORKSPACE + r'\KernelV0 (version 1) JetFireFormula.xlsx'

def _resolve_macro_path():
    """Return the first MacroQRAV6 file that exists: .xlsm then .xlsx."""
    for ext in ('.xlsm', '.xlsx'):
        p = _WORKSPACE + r'\MacroQRAV6 (version 1)' + ext
        if os.path.exists(p):
            return p
    raise FileNotFoundError(
        f'MacroQRAV6 (version 1).xlsm/.xlsx not found in {_WORKSPACE!r}'
    )

_MACRO_PATH  = _resolve_macro_path()
_N_SCENARIOS = 10      # scenarios to sample per event
_TOLERANCE   = 1e-4    # cell match threshold  |xl - py| < _TOLERANCE -> match
_RANDOM_SEED = 42
_OUTPUT_DIR  = os.path.join(os.path.dirname(__file__), '..', 'output')

# Restrict to specific events (empty list = all events)
_EVENTS_FILTER = []   # e.g. ['JF', 'TOXIC']

# Map event names -> v6 impact IDs (for setting Core!C2)
_EVENT_IMPACT_ID = {cfg['event']: iid for iid, cfg in IMPACT_CONFIG.items()}


# ── Helpers ───────────────────────────────────────────────────────────────────

def _col_num_to_letter(n):
    """1-indexed column number -> Excel letter string.  39 -> 'AM'."""
    letters = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(ord('A') + rem) + letters
    return letters


def _col_letter_to_num(letters):
    """Excel column letters -> 1-based number.  'AM' -> 39."""
    n = 0
    for c in letters.upper():
        n = n * 26 + (ord(c) - ord('A') + 1)
    return n


def _make_matrix_range(mat_col):
    """Build 'AB1:MD317'-style range string.

    Grid dimensions are read from _v6._P (populated at runtime from the
    General sheet); no hardcoded QX/QY here.
    """
    QX = int(_v6._P['QX'])
    QY = int(_v6._P['QY'])
    start = _col_letter_to_num(mat_col)
    end   = _col_num_to_letter(start + QX - 1)
    return f'{mat_col}1:{end}{QY}'


def _extract_weather(key):
    """Return 'Noche' or 'Dia' from a scenario key string."""
    if key.endswith('Noche'):
        return 'Noche'
    if key.endswith('Dia'):
        return 'Dia'
    return key.rsplit('/', 1)[-1]


def _v6_core_to_engine_core(v6_scenarios):
    """Convert read_core_scenarios() list -> qra_engine-style core dict.

    qra_engine reader functions expect {key: {x, y, size, probs: {col_idx: float}}}
    where col_idx is the 1-based openpyxl column for the probability in Core.
    """
    core = {}
    for sc in v6_scenarios:
        core[sc['key']] = {
            'x':    sc['x'],
            'y':    sc['y'],
            'size': sc['size'],
            'probs': {
                10: sc.get('prob',  0.0),   # TOXIC   (Core col J)
                11: sc.get('p_jf',  0.0),   # JF      (Core col K)
                12: sc.get('p_lpf', 0.0),   # LPF     (Core col L)
                13: sc.get('p_epf', 0.0),   # EPF     (Core col M)
                14: sc.get('p_fb',  0.0),   # FB      (Core col N)
                15: sc.get('p_cve', 0.0),   # CVE     (Core col O)
                16: sc.get('p_blv', 0.0),   # BLV     (Core col P)
                17: sc.get('p_ff',  0.0),   # FF      (Core col Q)
            }
        }
    return core


def _read_ak_labels(ws_xl, label_col='AK'):
    """Scan {label_col}1:{label_col}50 for non-None labels -> {label: row}."""
    scan_range = f'{label_col}1:{label_col}50'
    vals = ws_xl.range(scan_range).value
    if vals is None:
        return {}
    label_row = {}
    for i, v in enumerate(vals, start=1):
        if v is not None:
            label_row[str(v).strip()] = i
    return label_row


def _set_scenario(ws_xl, label_row, key, weather, x, y, value_col='AL'):
    """Write scenario parameters into the value column of KernelV0."""
    def _set(label, val):
        r = label_row.get(label)
        if r is not None:
            ws_xl.range(f'{value_col}{r}').value = val

    _set('Impact/Risk', 0)
    _set('Scenario',    key)
    _set('Weather',     weather)
    _set('X',           x)
    _set('Y',           y)
    _set('Probability', 0.0)


def _read_matrix(ws_xl, mat_col='AM'):
    """Read QY×QX matrix starting at mat_col -> (QY, QX) float64 array.

    Grid dimensions come from _v6._P (General sheet).
    Excel kernel row 1 = largest Y (top of map); flip so index 0 = smallest Y.
    """
    QY = int(_v6._P['QY'])
    QX = int(_v6._P['QX'])
    matrix_range = _make_matrix_range(mat_col)
    raw = ws_xl.range(matrix_range).value
    if raw is None:
        return np.zeros((QY, QX))
    arr = np.zeros((QY, QX), dtype=float)
    for i, row in enumerate(raw):
        if row is None:
            continue
        for j, v in enumerate(row):
            arr[i, j] = float(v) if v is not None else 0.0
    return arr[::-1, :]   # flip Y: Excel row 1 (top=large Y) -> Python index QY-1


def _check_grid_params(ws_xl, label_row, value_col='AL'):
    """Compare KernelV0 control-cell grid params against _v6._P values."""
    def _al(label):
        r = label_row.get(label)
        return ws_xl.range(f'{value_col}{r}').value if r else None

    xl_qx = _al('QX')
    xl_qy = _al('QY')
    xl_sx = _al('SX')
    xl_sy = _al('SY')
    p = _v6._P

    ok = True
    for name, xl_val, py_val, check in [
        ('QX', xl_qx, p['QX'], lambda a, b: a is not None and int(a) == int(b)),
        ('QY', xl_qy, p['QY'], lambda a, b: a is not None and int(a) == int(b)),
        ('SX', xl_sx, p['SX'], lambda a, b: a is not None and abs(a - b) < 1e-9),
        ('SY', xl_sy, p['SY'], lambda a, b: a is not None and abs(a - b) < 1e-9),
    ]:
        match = check(xl_val, py_val)
        flag = 'OK' if match else 'XX'
        print(f'    {flag} {name}: Excel={xl_val}  Python={py_val}')
        if not match:
            ok = False
    return ok


def _sync_grid_to_kernel(wb_xl):
    """Read KernelV0 grid params (SX/SY/QX/QY) and rebuild _v6 coordinate grids
    if they differ from the MacroQRAV6 General sheet values.

    KernelV0 may have been authored with a different SY than MacroQRAV6.
    For the comparison to be valid, Python cell-centre coordinates must match
    exactly what the kernel formula uses.  We patch _v6._P and rebuild
    _v6.XX / _v6.YY before running any scenario.

    The production qra_v6_engine run (driven by MacroQRAV6) is a separate
    process and is unaffected by this patch.
    """
    # ImpactJFMatrix has labelled control cells in AK/AL — most reliable source
    try:
        ws        = wb_xl.sheets['ImpactJFMatrix']
        label_row = _read_ak_labels(ws, 'AK')

        def _al(label):
            r = label_row.get(label)
            return ws.range(f'AL{r}').value if r else None

        k_qx = _al('QX')
        k_qy = _al('QY')
        k_sx = _al('SX')
        k_sy = _al('SY')
    except Exception as exc:
        print(f'  [WARN] Could not read KernelV0 grid params from ImpactJFMatrix: {exc}')
        return

    p       = _v6._P
    changed = False

    for name, k_val, p_key, cast in [
        ('QX', k_qx, 'QX', lambda v: int(round(v))),
        ('QY', k_qy, 'QY', lambda v: int(round(v))),
        ('SX', k_sx, 'SX', float),
        ('SY', k_sy, 'SY', float),
    ]:
        if k_val is None:
            continue
        k_cast = cast(k_val)
        p_cast = cast(p[p_key])
        if abs(k_cast - p_cast) > 1e-9:
            print(f'  [GRID SYNC] {name}: KernelV0={k_val}  MacroQRAV6={p[p_key]}')
            print(f'              Using KernelV0 value so comparison grid matches kernel.')
            p[p_key] = k_cast
            changed = True

    if changed:
        QX = int(p['QX'])
        QY = int(p['QY'])
        xc = p['SX'] * (np.arange(QX) + 0.5)
        yc = p['SY'] * (np.arange(QY) + 0.5)
        _v6.XX, _v6.YY = np.meshgrid(xc, yc)
        print(f'  [GRID SYNC] Rebuilt: {QX}×{QY}  SX={p["SX"]:.10f}  SY={p["SY"]:.10f}')
    else:
        print(f'  [GRID SYNC] KernelV0 grid matches MacroQRAV6 — no rebuild needed.')


def _metrics(xl_mat, py_mat):
    """Compute match_rate, max_abs_err, mean_abs_err (nonzero region), counts."""
    diff = np.abs(xl_mat - py_mat)
    match_rate = float(np.mean(diff < _TOLERANCE))
    max_err    = float(np.max(diff))
    nz_mask    = (xl_mat != 0) | (py_mat != 0)
    mean_err   = float(np.mean(diff[nz_mask])) if nz_mask.any() else 0.0
    nz_xl      = int(np.count_nonzero(xl_mat))
    nz_py      = int(np.count_nonzero(py_mat))
    return match_rate, max_err, mean_err, nz_xl, nz_py


# ── Python formula wrappers (use qra_v6_engine formulas via _P) ───────────────

def _py_thermal(sc):
    # formula_thermal uses _P['THERM_THRESHOLDS'] internally — no hardcoded arg
    return formula_thermal(dist_grid(sc['sx'], sc['sy']), sc['therm_dists'])


def _py_jf(sc):
    return formula_jf(sc['sx'], sc['sy'],
                      sc['dist_vals'], sc['halfW_vals'], sc['center_vals'])


def _py_toxic(sc):
    return formula_toxic(dist_grid(sc['sx'], sc['sy']),
                         sc['tox_dists'], sc['tox_probs'])


def _py_explosion_cve(sc):
    # CVE uses ignition-point coordinates when available
    sx = sc['ign_x'] if sc.get('ign_x') is not None else sc['sx']
    sy = sc['ign_y'] if sc.get('ign_y') is not None else sc['sy']
    return formula_explosion(dist_grid(sx, sy), sc['exp_dists'])


def _py_ff(sc):
    return formula_ff(dist_grid(sc['sx'], sc['sy']),
                      sc['lfl_dist'], sc['lflf_dist'])


# ── Event catalogue ───────────────────────────────────────────────────────────
# (event_name, kernel_sheet, reader_fn, python_formula_fn, x_fn, y_fn,
#  label_col, value_col, mat_col)
#
# JF uses AK/AL/AM (original layout); all others use Z/AA/AB.
_CATALOGUE = [
    ('TOXIC', 'ImpactToxMatrix',   read_toxic_scenarios,     _py_toxic,
     lambda sc: sc['sx'], lambda sc: sc['sy'],
     'Z', 'AA', 'AB'),
    ('JF',    'ImpactJFMatrix',    read_jf_scenarios,        _py_jf,
     lambda sc: sc['sx'], lambda sc: sc['sy'],
     'AK', 'AL', 'AM'),
    ('LPF',   'ImpactThermMatrix', read_thermal_scenarios,   _py_thermal,
     lambda sc: sc['sx'], lambda sc: sc['sy'],
     'Z', 'AA', 'AB'),
    ('CVE',   'ImpactExpMatrix',   read_explosion_scenarios, _py_explosion_cve,
     lambda sc: sc['ign_x'] if sc.get('ign_x') else sc['sx'],
     lambda sc: sc['ign_y'] if sc.get('ign_y') else sc['sy'],
     'Z', 'AA', 'AB'),
    ('FF',    'ImpactFFMatrix',    read_ff_scenarios,        _py_ff,
     lambda sc: sc['sx'], lambda sc: sc['sy'],
     'Z', 'AA', 'AB'),
]


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    os.makedirs(_OUTPUT_DIR, exist_ok=True)
    rng = np.random.default_rng(_RANDOM_SEED)

    catalogue = [e for e in _CATALOGUE
                 if not _EVENTS_FILTER or e[0] in _EVENTS_FILTER]

    # ── Phase 0 + 1a: load General params, then read Core per event ───────────
    print('Phase 0+1a -- loading General params and Core from MacroQRAV6 (xlwings)...')
    print(f'  File: {_MACRO_PATH}')

    core_by_event = {}

    app_macro = xw.App(visible=False)
    try:
        macro_path = _MACRO_PATH
        try:
            wb_macro = app_macro.books.open(macro_path)
        except Exception as open_err:
            xlsx_fallback = macro_path.replace('.xlsm', '.xlsx')
            if os.path.exists(xlsx_fallback):
                print(f'  [WARN] Cannot open {os.path.basename(macro_path)}: {open_err}')
                print(f'  [INFO] Falling back to {os.path.basename(xlsx_fallback)}')
                macro_path = xlsx_fallback
                wb_macro = app_macro.books.open(macro_path)
            else:
                raise RuntimeError(
                    f'Cannot open MacroQRAV6:\n  {open_err}\n\n'
                    f'Fix: open the file in Excel and resave it as '
                    f'"MacroQRAV6 (version 1).xlsx" (Excel Workbook, no macros).'
                ) from open_err

        # Phase 0: populate qra_v6_engine._P from MacroQRAV6 General sheet.
        # All formula functions in qra_v6_engine read _P at call time, so this
        # must happen before any formula is invoked.
        _P = load_general_params(wb_macro)
        print(f'  Grid            : QX={int(_P["QX"])}  QY={int(_P["QY"])}  '
              f'SX={_P["SX"]:.8f}  SY={_P["SY"]:.8f}')
        print(f'  Thermal kW/m²   : {_P["THERM_THRESHOLDS"].tolist()}')
        print(f'  Overpressure bar: {_P["EXP_THRESHOLDS"].tolist()}')
        print(f'  TOX min prob    : {_P["TOX_MIN_PROB"]}')
        print(f'  FF zones        : {_P["FF_OUTSIDE"]} / {_P["FF_TRANSITION"]} / {_P["FF_INSIDE_LFL"]}')

        # Propagate dynamic values to qra_engine so its reader functions use
        # the same thresholds and filters as the v6 engine.
        _qra_engine.THERM_THRESHOLDS = _P['THERM_THRESHOLDS']
        _qra_engine.EXP_THRESHOLDS   = _P['EXP_THRESHOLDS']
        _qra_engine.TOX_MIN_PROB      = float(_P['TOX_MIN_PROB'])
        _qra_engine.FF_OUTSIDE        = float(_P['FF_OUTSIDE'])
        _qra_engine.FF_TRANSITION     = float(_P['FF_TRANSITION'])
        _qra_engine.FF_INSIDE_LFL     = float(_P['FF_INSIDE_LFL'])

        ws_core = wb_macro.sheets['Core']

        for ev_name, _, _, _, _, _, _, _, _ in catalogue:
            impact_id = _EVENT_IMPACT_ID.get(ev_name)
            if impact_id is None:
                print(f'  {ev_name:<8}: no impact_id mapping -- skipped')
                continue

            ws_core.range('C2').value = impact_id
            time.sleep(1.0)
            app_macro.api.Calculate()
            time.sleep(0.5)

            v6_scenarios = read_core_scenarios(ws_core)
            core_by_event[ev_name] = _v6_core_to_engine_core(v6_scenarios)
            print(f'  {ev_name:<8}: impact_id={impact_id}  '
                  f'{len(v6_scenarios)} Core rows -> {len(core_by_event[ev_name])} keys')

        wb_macro.close()
    finally:
        app_macro.quit()

    # ── Phase 1b: KernelV0 (openpyxl) -- read consequence distances ──────────
    print(f'\nPhase 1b -- reading consequence distances from KernelV0 (openpyxl)...')
    print(f'  File: {_EXCEL_PATH}')

    wb_opx = openpyxl.load_workbook(_EXCEL_PATH, data_only=True)
    _qra_engine.WIND_ANGLE_DEG = read_wind_direction(wb_opx)
    print(f'  Wind direction: {_qra_engine.WIND_ANGLE_DEG:.1f} deg '
          f'(math convention: 0=E, 90=N) from WindMatrix')
    event_scenarios = {}
    sheet_cache = {}

    for ev_name, sheet, reader_fn, _, _, _, _, _, _ in catalogue:
        core = core_by_event.get(ev_name, {})
        if not core:
            print(f'  {ev_name:<8}: no Core data -- skipped')
            continue
        if sheet not in sheet_cache:
            sheet_cache[sheet] = reader_fn(wb_opx, core)
        event_scenarios[ev_name] = sheet_cache[sheet]
        print(f'  {ev_name:<8}: {len(event_scenarios[ev_name]):4d} scenarios  ({sheet})')

    wb_opx.close()

    # ── Phase 2: KernelV0 (xlwings) -- kernel vs Python comparison ───────────
    QX = int(_v6._P['QX'])
    QY = int(_v6._P['QY'])
    print(f'\nPhase 2 -- kernel comparison via xlwings...')
    print(f'  File:         {_EXCEL_PATH}')
    print(f'  Grid size:    {QX} cols x {QY} rows  (from General sheet)')
    print(f'  Scenarios:    {_N_SCENARIOS} per event  (tolerance={_TOLERANCE})')

    all_rows = []

    app = xw.App(visible=False)
    try:
        wb_xl = app.books.open(_EXCEL_PATH)

        # Align Python grid to KernelV0's own SX/SY before any formula runs.
        print('  Checking KernelV0 grid params...')
        _sync_grid_to_kernel(wb_xl)

        for ev_name, sheet, _, py_fn, x_fn, y_fn, label_col, value_col, mat_col in catalogue:
            scenarios = event_scenarios.get(ev_name, [])
            if not scenarios:
                print(f'\n{ev_name}: 0 scenarios -- skipped')
                continue

            ws_xl     = wb_xl.sheets[sheet]
            label_row = _read_ak_labels(ws_xl, label_col)

            print(f'\n-- {ev_name}  ({sheet}) --')
            print(f'  Control cols: labels={label_col}  values={value_col}  matrix starts={mat_col}')
            print('  Grid params (KernelV0 vs General sheet):')
            _check_grid_params(ws_xl, label_row, value_col)

            n   = min(_N_SCENARIOS, len(scenarios))
            idx = rng.choice(len(scenarios), n, replace=False)
            sample = [scenarios[i] for i in sorted(idx)]

            hdr = (f'  {"Scenario key":<44} {"match%":>7} '
                   f'{"maxErr":>10} {"meanErr":>10} {"nz_XL":>7} {"nz_Py":>7}')
            print(hdr)
            print('  ' + '-' * (len(hdr) - 2))

            for sc in sample:
                weather = _extract_weather(sc['key'])
                x = x_fn(sc)
                y = y_fn(sc)

                _set_scenario(ws_xl, label_row, sc['key'], weather, x, y, value_col)
                app.calculate()

                xl_mat = _read_matrix(ws_xl, mat_col)
                py_mat = py_fn(sc)

                mr, max_e, mean_e, nz_xl, nz_py = _metrics(xl_mat, py_mat)

                key_short = sc['key'][:44]
                print(f'  {key_short:<44} {mr:>7.1%} {max_e:>10.3e} '
                      f'{mean_e:>10.3e} {nz_xl:>7d} {nz_py:>7d}')

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

    # ── Save CSV ──────────────────────────────────────────────────────────────
    if all_rows:
        out_path = os.path.join(_OUTPUT_DIR, 'compare_excel_results.csv')
        with open(out_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=list(all_rows[0].keys()))
            writer.writeheader()
            writer.writerows(all_rows)
        print(f'\nResults saved -> {out_path}')

    # ── Summary ───────────────────────────────────────────────────────────────
    if all_rows:
        print('\n-- Summary --')
        by_event = {}
        for r in all_rows:
            by_event.setdefault(r['event'], []).append(r)
        for ev, rows in by_event.items():
            avg_mr   = np.mean([float(r['match_rate']) for r in rows])
            avg_maxe = np.mean([float(r['max_abs_error']) for r in rows])
            print(f'  {ev:<8}  avg_match={avg_mr:.1%}  '
                  f'avg_maxErr={avg_maxe:.3e}  n={len(rows)}')


if __name__ == '__main__':
    main()
