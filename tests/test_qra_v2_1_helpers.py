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


if __name__ == '__main__':
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
