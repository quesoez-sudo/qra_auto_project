"""
test_live_readers.py
=====================
Regression check: each NEW read_<event>_live() (xlwings, reads the ALREADY
OPEN workbook) must return exactly the same scenario data as the existing
read_<event>() (openpyxl, static read) for the same file. This is step 2 of
the CalcBU button integration testing plan (docs/superpowers/specs/2026-07-14
-calcbu-button-integration-design.md) — confirms the openpyxl->xlwings port
introduces no data differences before it's ever wired to the IRiskMap button.

Requires: xlwings + Excel installed, and MacroQRAV6 (version 1)-Win11-
cWvAMQq4D.xlsm present at the project root (see qra_engine_v2_1.INPUT_FILE).
Opens Excel invisibly; does not save or modify the workbook.

Usage:
  python tests/test_live_readers.py
"""
import math
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import openpyxl
import xlwings as xw

import qra_engine_v2_1 as eng

FAILURES = []


def _check(label, cond, detail=''):
    status = 'PASS' if cond else 'FAIL'
    print(f'  [{status}] {label}' + (f' — {detail}' if detail and not cond else ''))
    if not cond:
        FAILURES.append(f'{label}: {detail}')


def _floats_equal(a, b):
    a_nan = isinstance(a, float) and math.isnan(a)
    b_nan = isinstance(b, float) and math.isnan(b)
    if a_nan or b_nan:
        return a_nan and b_nan
    return a == b or abs(a - b) < 1e-9


def _values_equal(sv, lv):
    if isinstance(sv, list):
        return isinstance(lv, list) and len(sv) == len(lv) and all(
            _floats_equal(a, b) for a, b in zip(sv, lv))
    if isinstance(sv, float):
        return _floats_equal(sv, lv)
    return sv == lv


def _rows_equal(label, static_rows, live_rows):
    _check(f'{label}: same scenario count', len(static_rows) == len(live_rows),
           f'static={len(static_rows)} live={len(live_rows)}')
    n = min(len(static_rows), len(live_rows))
    mismatches = 0
    for i in range(n):
        s, l = static_rows[i], live_rows[i]
        for key, sv in s.items():
            lv = l.get(key)
            if not _values_equal(sv, lv):
                mismatches += 1
                print(f'      row {i} key_scenario={s.get("path_code")!r} '
                      f'field={key!r}: static={sv!r} live={lv!r}')
    _check(f'{label}: all {n} rows match field-for-field', mismatches == 0,
           f'{mismatches} field mismatch(es)')


def main():
    print(f'Opening {eng.INPUT_FILE} (static, openpyxl)...')
    wb_static = openpyxl.load_workbook(eng.INPUT_FILE, data_only=True)

    # NOTE: load_general_params() must run BEFORE Excel opens the same file live
    # below — once xlwings/Excel has the workbook open, Excel's OS-level file
    # lock blocks openpyxl from re-opening it (PermissionError), which is
    # exactly the bug load_general_params_live() exists to work around in
    # _run_pipeline. Capturing static_params here (pre-lock) lets us still
    # compare it against the live reader's output further down.
    static_params = eng.load_general_params()

    print(f'Opening {eng.INPUT_FILE} (live, xlwings)...')
    app = xw.App(visible=False)
    app.display_alerts = False
    app.api.AutomationSecurity = 3
    try:
        wb_live = app.books.open(eng.INPUT_FILE)

        print('\n=== load_general_params vs load_general_params_live (General) ===')
        live_params = eng.load_general_params_live(wb_live.sheets['General'])
        for key in static_params:
            sv, lv = static_params[key], live_params.get(key)
            if isinstance(sv, list):
                ok = isinstance(lv, list) and len(sv) == len(lv) and all(
                    _floats_equal(a, b) for a, b in zip(sv, lv))
            elif isinstance(sv, float):
                ok = _floats_equal(sv, lv)
            else:
                ok = sv == lv
            _check(f'General params: {key} matches', ok, f'static={sv!r} live={lv!r}')

        print('\n=== read_therm vs read_therm_live (Late Pool Fire Results) ===')
        static_rows = eng.read_therm(wb_static['Late Pool Fire Results'])
        live_rows = eng.read_therm_live(wb_live.sheets['Late Pool Fire Results'])
        _rows_equal('thermal', static_rows, live_rows)

        print('\n=== read_ff vs read_ff_live (Flash Fire Results) ===')
        static_rows = eng.read_ff(wb_static['Flash Fire Results'])
        live_rows = eng.read_ff_live(wb_live.sheets['Flash Fire Results'])
        _rows_equal('flash fire', static_rows, live_rows)

        print('\n=== read_tox vs read_tox_live (Outdoor Toxic Results) ===')
        static_rows = eng.read_tox(wb_static['Outdoor Toxic Results'])
        live_rows = eng.read_tox_live(wb_live.sheets['Outdoor Toxic Results'])
        _rows_equal('toxic', static_rows, live_rows)

        print('\n=== read_exp vs read_exp_live (CVE Results) ===')
        static_rows = eng.read_exp(wb_static['CVE Results'])
        live_rows = eng.read_exp_live(wb_live.sheets['CVE Results'])
        _rows_equal('explosion', static_rows, live_rows)

        print('\n=== read_jf vs read_jf_live (Jet Fire Results) ===')
        static_rows = eng.read_jf(wb_static['Jet Fire Results'])
        live_rows = eng.read_jf_live(wb_live.sheets['Jet Fire Results'])
        _rows_equal('jet fire', static_rows, live_rows)

        wb_live.close()
    finally:
        app.quit()
    wb_static.close()

    print(f'\n{"="*60}')
    if FAILURES:
        print(f'{len(FAILURES)} FAILURE(S):')
        for f in FAILURES:
            print(f'  - {f}')
        raise SystemExit(1)
    print('ALL CHECKS PASSED')


if __name__ == '__main__':
    main()
