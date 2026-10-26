"""
test_qra_v2_1_helpers.py
=========================
Fast, Excel-free unit tests for qra_engine_v2_1's button-support helpers:
SHEET_IMPACT/SHEET_RISK constants, _bulk_read, _frozen, _progress,
_format_warning_summary, _show_msgbox, _timed, _format_timing_report. No
xlwings App is opened here — Fake objects stand in for xlwings Sheet/App/Api
so these run in well under a second. Excel-dependent behavior is covered
separately by tests/test_live_readers.py and
tests/test_run_from_button_headless.py.

Usage:
  python tests/test_qra_v2_1_helpers.py
"""
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import qra_engine_v2_1 as eng

FAILURES = []


def _check(label, cond, detail=''):
    status = 'PASS' if cond else 'FAIL'
    print(f'  [{status}] {label}' + (f' — {detail}' if detail and not cond else ''))
    if not cond:
        FAILURES.append(f'{label}: {detail}')


def test_sheet_constants():
    _check('SHEET_IMPACT is CalcBUImpact', eng.SHEET_IMPACT == 'CalcBUImpact',
           f'got {eng.SHEET_IMPACT!r}')
    _check('SHEET_RISK is CalcBURisk', eng.SHEET_RISK == 'CalcBURisk',
           f'got {eng.SHEET_RISK!r}')


class _FakeApi:
    def __init__(self):
        self.EnableEvents = True
        self.StatusBar = None


class _FakeApp:
    def __init__(self):
        self.screen_updating = True
        self.calculation = 'automatic'
        self.api = _FakeApi()


def test_frozen_sets_and_restores_on_success():
    app = _FakeApp()
    with eng._frozen(app):
        _check('screen_updating False inside the with-block', app.screen_updating is False)
        _check("calculation 'manual' inside the with-block", app.calculation == 'manual')
        _check('EnableEvents False inside the with-block', app.api.EnableEvents is False)
    _check('screen_updating restored after the with-block', app.screen_updating is True)
    _check('calculation restored after the with-block', app.calculation == 'automatic')
    _check('EnableEvents restored after the with-block', app.api.EnableEvents is True)


def test_frozen_restores_on_exception():
    app = _FakeApp()
    try:
        with eng._frozen(app):
            raise RuntimeError('boom')
    except RuntimeError:
        pass
    _check('screen_updating restored after an exception', app.screen_updating is True)
    _check('calculation restored after an exception', app.calculation == 'automatic')
    _check('EnableEvents restored after an exception', app.api.EnableEvents is True)


def test_progress_sets_status_bar_text():
    app = _FakeApp()
    eng._progress(app, 'Procesando TOXIC (1/7)...')
    _check('StatusBar set to the progress text',
           app.api.StatusBar == 'Procesando TOXIC (1/7)...', f'got {app.api.StatusBar!r}')


def test_format_warning_summary_empty():
    out = eng._format_warning_summary([])
    _check('empty warnings -> "Sin advertencias."', out == 'Sin advertencias.', f'got {out!r}')


def test_format_warning_summary_with_items():
    out = eng._format_warning_summary(['Impact 22 (BLV): no reader', 'ERROR FATAL: boom'])
    _check('summary starts with the count', out.startswith('2 advertencia(s):'), f'got {out!r}')
    _check('item 1 is numbered', '1. Impact 22 (BLV): no reader' in out)
    _check('item 2 is numbered', '2. ERROR FATAL: boom' in out)


def test_show_msgbox_calls_windows_api_with_expected_args():
    calls = []
    original = eng.ctypes.windll.user32.MessageBoxW
    eng.ctypes.windll.user32.MessageBoxW = lambda hwnd, text, title, flags: calls.append((hwnd, text, title, flags))
    try:
        eng._show_msgbox('hello', title='T')
    finally:
        eng.ctypes.windll.user32.MessageBoxW = original
    _check('MessageBoxW called exactly once', len(calls) == 1, f'{len(calls)} call(s)')
    if calls:
        _check('called with (0, text, title, MB_ICONINFORMATION)',
               calls[0] == (0, 'hello', 'T', 0x40), f'got {calls[0]!r}')


def test_timed_accumulates_across_multiple_uses():
    eng._timings.clear()
    with eng._timed('reading'):
        time.sleep(0.01)
    with eng._timed('reading'):
        time.sleep(0.01)
    with eng._timed('writing'):
        time.sleep(0.01)
    _check("'reading' accumulated across two with-blocks",
           eng._timings['reading'] >= 0.02, f"got {eng._timings.get('reading')!r}")
    _check("'writing' recorded separately from 'reading'",
           eng._timings['writing'] >= 0.01 and eng._timings['writing'] < eng._timings['reading'],
           f"reading={eng._timings.get('reading')!r} writing={eng._timings.get('writing')!r}")


def test_timed_stops_the_clock_on_exception():
    eng._timings.clear()
    try:
        with eng._timed('calculation'):
            time.sleep(0.01)
            raise RuntimeError('boom')
    except RuntimeError:
        pass
    _check("'calculation' still recorded despite the exception",
           eng._timings.get('calculation', 0.0) >= 0.01)


def test_format_timing_report_shape():
    report = eng._format_timing_report(
        {'reading': 1.0, 'calculation': 2.5, 'iteration': 0.5, 'writing': 3.0})
    _check('report mentions all four stages',
           all(s in report for s in ['Reading', 'Calculation', 'Iteration', 'Writing']))
    _check('report includes a Total line', 'Total' in report)
    _check('total is the sum of the four stages', '7.0' in report, f'got {report!r}')


class _FakeRange:
    def __init__(self, value):
        self.value = value


class _FakeLiveSheet:
    """Mimics xlwings' Sheet.range((r0,c0),(r1,c1)).value shape behavior:
    scalar for 1x1, flat list for 1xN or Nx1, nested list for NxM."""
    def __init__(self, grid):
        self._grid = grid   # {(row, col): value}, 1-based

    def range(self, start, end):
        r0, c0 = start
        r1, c1 = end
        if r0 == r1 and c0 == c1:
            return _FakeRange(self._grid.get((r0, c0)))
        if r0 == r1:
            return _FakeRange([self._grid.get((r0, c)) for c in range(c0, c1 + 1)])
        if c0 == c1:
            return _FakeRange([self._grid.get((r, c0)) for r in range(r0, r1 + 1)])
        return _FakeRange([[self._grid.get((r, c)) for c in range(c0, c1 + 1)]
                            for r in range(r0, r1 + 1)])


def test_bulk_read_normalizes_shapes():
    ws = _FakeLiveSheet({(1, 1): 'a', (1, 2): 'b', (2, 1): 'c', (2, 2): 'd'})
    _check('2x2 block', eng._bulk_read(ws, 1, 2, 1, 2) == [['a', 'b'], ['c', 'd']])
    _check('1x2 row normalized to [[..]]', eng._bulk_read(ws, 1, 1, 1, 2) == [['a', 'b']])
    _check('2x1 col normalized to [[..],[..]]', eng._bulk_read(ws, 1, 2, 1, 1) == [['a'], ['c']])
    _check('1x1 cell normalized to [[..]]', eng._bulk_read(ws, 1, 1, 1, 1) == [['a']])
    empty_ws = _FakeLiveSheet({})
    _check('missing 1x1 cell -> [[None]]', eng._bulk_read(empty_ws, 5, 5, 5, 5) == [[None]])


def test_compute_event_live_impact_max_risk_sum():
    # Two overlapping FF scenarios at the same source: impact must be the
    # cell-wise MAX (kernel MaxMatrix), risk the cell-wise SUM (kernel SumMatrix).
    evt = {'formula': 'ff', 'freq_col': 'P_FF'}
    core = [
        {'key': 'A', 'X': 100.0, 'Y': 100.0, 'size': 'M', 'P_FF': 1e-4},
        {'key': 'B', 'X': 100.0, 'Y': 100.0, 'size': 'M', 'P_FF': 2e-4},
    ]
    results_idx = {'A': {'lfl_r': 10.0, 'lflf_r': 20.0},
                   'B': {'lfl_r': 10.0, 'lflf_r': 20.0}}
    imp, risk, matched, _ = eng.compute_event_live(evt, core, results_idx)
    _check('both scenarios matched', matched == 2, f'got {matched}')
    for sz in ('Total', 'M'):
        _check(f'[{sz}] impact max is MAX of scenarios (2), not SUM (4)',
               imp[sz].max() == 2.0, f'got {imp[sz].max()}')
        _check(f'[{sz}] risk max is SUM of frequencies (3e-4)',
               abs(risk[sz].max() - 3e-4) < 1e-12, f'got {risk[sz].max()}')
    _check('[S] untouched', imp['S'].max() == 0.0 and risk['S'].max() == 0.0)


# ── CVE ───────────────────────────────────────────────────────────────────────
import numpy as np

_CVE_LEVELS = [0.04, 0.1, 0.35, 0.5, 1.0]
_CVE_DISTS  = [3241.44, 1421.44, 469.474, 345.199, 195.768]   # D-4766 Dia, 1E row

# VUL_OVP rows (AJ, AK, AL:AO) copied from 'MacroQRAV6 Primera corrida.xlsm'
_VUL_ROWS = [
    ['CIA1', 'CIA1_1', 0.01, 0.45, 0.55, 0.6],
    ['CIA1', 'CIA1_2', 0.55, 0.6, 1, 1],
    ['OPEN AREA (ONSHORE)', 'OPEN AREA (ONSHORE)_1', 0.01, 0.05, 0.3, 0.35],
    ['OPEN AREA (ONSHORE)', 'OPEN AREA (ONSHORE)_2', 0.3, 0.35, 0.999, 0.5],
    ['OPEN AREA (ONSHORE)', 'OPEN AREA (ONSHORE)_3', 0.999, 0.5, 1, 1],
    [None, None, None, None, None, None],
]


def test_compute_exp_linear_interpolation():
    d = np.array([[_CVE_DISTS[0], _CVE_DISTS[2], 100.0, 4000.0,
                   (_CVE_DISTS[1] + _CVE_DISTS[2]) / 2]])
    imp = eng.compute_exp(d, _CVE_DISTS, _CVE_LEVELS, mode='linear')
    _check('CVE linear: at J distance -> 0.04 bar', abs(imp[0, 0] - 0.04) < 1e-12, f'got {imp[0, 0]}')
    _check('CVE linear: at L distance -> 0.35 bar', abs(imp[0, 1] - 0.35) < 1e-12, f'got {imp[0, 1]}')
    _check('CVE linear: inside N -> 1 bar cap', imp[0, 2] == 1.0, f'got {imp[0, 2]}')
    _check('CVE linear: beyond J -> 0', imp[0, 3] == 0.0, f'got {imp[0, 3]}')
    _check('CVE linear: K-L midpoint -> 0.225 bar', abs(imp[0, 4] - 0.225) < 1e-12, f'got {imp[0, 4]}')


def test_compute_exp_loglog_interpolation():
    mid = (_CVE_DISTS[1] + _CVE_DISTS[2]) / 2
    d = np.array([[_CVE_DISTS[1], _CVE_DISTS[2], mid]])
    imp = eng.compute_exp(d, _CVE_DISTS, _CVE_LEVELS, mode='loglog')
    slope = np.log(0.35 / 0.1) / np.log(_CVE_DISTS[2] / _CVE_DISTS[1])
    expected = 0.1 * (mid / _CVE_DISTS[1]) ** slope
    _check('CVE loglog: at K distance -> 0.1 bar', abs(imp[0, 0] - 0.1) < 1e-12, f'got {imp[0, 0]}')
    _check('CVE loglog: at L distance -> 0.35 bar', abs(imp[0, 1] - 0.35) < 1e-12, f'got {imp[0, 1]}')
    _check('CVE loglog: K-L midpoint follows power law', abs(imp[0, 2] - expected) < 1e-12,
           f'got {imp[0, 2]} expected {expected}')
    _check('CVE loglog: midpoint below linear (0.225)', imp[0, 2] < 0.225, f'got {imp[0, 2]}')


def test_compute_exp_default_mode_is_linear():
    _check('CVE_INTERP_MODE default is linear', eng.CVE_INTERP_MODE == 'linear',
           f'got {eng.CVE_INTERP_MODE!r}')


def test_compute_exp_all_zero_row_is_empty():
    d = np.array([[0.0, 10.0, 1000.0]])
    imp = eng.compute_exp(d, [0, 0, 0, 0, 0], _CVE_LEVELS)
    risk = eng.compute_exp(d, [0, 0, 0, 0, 0], _CVE_LEVELS, return_risk=True, frequency=1.0,
                           vul_curve=eng._parse_vul_rows(_VUL_ROWS)['CIA1'])
    _check('CVE all-zero row: impact empty', imp.max() == 0.0)
    _check('CVE all-zero row: risk empty', risk.max() == 0.0)


def test_vul_ovp_lethality_matches_sheet():
    curves = eng._parse_vul_rows(_VUL_ROWS)
    _check('VUL_OVP: two building types parsed', set(curves) == {'CIA1', 'OPEN AREA (ONSHORE)'},
           f'got {set(curves)}')
    oa = curves['OPEN AREA (ONSHORE)']
    ovp = np.array([0.01, 0.05, 0.35, 0.5, 1.0, 1.5])
    L = eng._ovp_lethality(ovp, oa)
    expected = [0.0, 0.01, 0.3, 0.999, 1.0, 1.0]
    _check('VUL_OVP open area: breakpoints', np.allclose(L, expected), f'got {L.tolist()}')
    # CIA1_1 at OVP=0.55 -> sheet AU column gives 0.1636734711960405
    L1 = eng._ovp_lethality(np.array([0.55]), curves['CIA1'])[0]
    _check('VUL_OVP CIA1 at 0.55 bar matches sheet', abs(L1 - 0.1636734711960405) < 1e-9, f'got {L1}')


def test_location_rows_parse():
    rows = [['1E -Horno H4753 ', 1, 'Oil refinery', True, 'id', 1, 0, 1, 'OPEN AREA (ONSHORE)'],
            ['2U - RIE', 1, 'Oil refinery', True, 'id', 1, 0, 0, 'CIA4'],
            [None] * 9]
    bt = eng._parse_location_rows(rows)
    _check('Location Data: name -> building type',
           bt == {'1E -Horno H4753': 'OPEN AREA (ONSHORE)', '2U - RIE': 'CIA4'}, f'got {bt}')


def test_compute_event_live_cve_rows_per_ignition():
    # One scenario, two ignition rows at different centres: impact = MAX over rows,
    # risk = SUM with FULL P_CVE per row (TODO(CVE-freq) — double counting kept on purpose).
    evt = {'formula': 'exp', 'freq_col': 'P_CVE'}
    core = [{'key': 'A', 'X': 0.0, 'Y': 0.0, 'size': 'INST', 'P_CVE': 1e-4}]
    dists = [60.0, 40.0, 20.0, 15.0, 10.0]
    rows = [
        {'path_code': 'A', 'location': 'L1', 'xign': 100.0, 'yign': 100.0, 'distances': dists},
        {'path_code': 'A', 'location': 'L1', 'xign': 101.0, 'yign': 100.0, 'distances': dists},
        {'path_code': 'A', 'location': 'NOPE', 'xign': float('nan'), 'yign': 1.0, 'distances': dists},
    ]
    vul = {'L1': eng._parse_vul_rows(_VUL_ROWS)['OPEN AREA (ONSHORE)']}
    imp, risk, matched, unmatched = eng.compute_event_live(evt, core, eng._group_by_key(rows), vul)
    _check('CVE: scenario matched', matched == 1 and unmatched == 0, f'got {matched}/{unmatched}')
    _check('CVE: impact capped at 1 bar (MAX, not SUM)', imp['INST'].max() == 1.0,
           f'got {imp["INST"].max()}')
    _check('CVE: risk = 2 rows x full P_CVE inside 1 bar', abs(risk['INST'].max() - 2e-4) < 1e-12,
           f'got {risk["INST"].max()}')
    _check('CVE: Total equals INST', np.array_equal(risk['Total'], risk['INST']))
    _check('CVE: multi-row key detected', eng._cve_multi_row_keys(eng._group_by_key(rows)) == ['A'])


if __name__ == '__main__':
    test_compute_exp_linear_interpolation()
    test_compute_exp_loglog_interpolation()
    test_compute_exp_default_mode_is_linear()
    test_compute_exp_all_zero_row_is_empty()
    test_vul_ovp_lethality_matches_sheet()
    test_location_rows_parse()
    test_compute_event_live_cve_rows_per_ignition()
    test_compute_event_live_impact_max_risk_sum()
    test_sheet_constants()
    test_frozen_sets_and_restores_on_success()
    test_frozen_restores_on_exception()
    test_progress_sets_status_bar_text()
    test_format_warning_summary_empty()
    test_format_warning_summary_with_items()
    test_show_msgbox_calls_windows_api_with_expected_args()
    test_timed_accumulates_across_multiple_uses()
    test_timed_stops_the_clock_on_exception()
    test_format_timing_report_shape()
    test_bulk_read_normalizes_shapes()

    print(f'\n{"="*60}')
    if FAILURES:
        print(f'{len(FAILURES)} FAILURE(S):')
        for f in FAILURES:
            print(f'  - {f}')
        raise SystemExit(1)
    print('ALL CHECKS PASSED')
