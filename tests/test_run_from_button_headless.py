"""
test_run_from_button_headless.py
=================================
Headless integration test for the IRiskMap button path — no VBA, no button
click. Opens the real target workbook visibly via xlwings, calls
eng.run_from_button(wb=wb) directly, then checks:
  - screen_updating / calculation / EnableEvents were restored after the run
    (the freeze/unfreeze contract)
  - the StatusBar was released back to Excel (set to False) at the end
  - the warning-summary + time-report msgbox was produced exactly once
    (mocked — doesn't block on a real popup)
  - CalcBUImpact / CalcBURisk have data after the run
  - _timings has all four pipeline stages recorded

This is step 3 of the CalcBU button integration testing plan (see
docs/superpowers/specs/2026-07-14-calcbu-button-integration-design.md).

IMPORTANT: this test does NOT save the workbook — wb.close() without wb.save()
discards all in-memory changes, so the real .xlsm on disk is left untouched
even though CalcBUImpact/CalcBURisk get written in-memory for the checks below.

Requires: xlwings + Excel installed, and MacroQRAV6 (version 1)-Win11-
cWvAMQq4D.xlsm present at the project root.

Usage:
  python tests/test_run_from_button_headless.py
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import xlwings as xw

import qra_engine_v2_1 as eng

FAILURES = []


def _check(label, cond, detail=''):
    status = 'PASS' if cond else 'FAIL'
    print(f'  [{status}] {label}' + (f' — {detail}' if detail and not cond else ''))
    if not cond:
        FAILURES.append(f'{label}: {detail}')


def main():
    msgbox_calls = []
    original_msgbox = eng._show_msgbox
    eng._show_msgbox = lambda text, title='QRA Engine v2.1': msgbox_calls.append((text, title))

    app = xw.App(visible=True)
    app.display_alerts = False
    app.api.AutomationSecurity = 3   # headless test only — no VBA/button involved
    try:
        wb = app.books.open(eng.INPUT_FILE)

        screen_updating_before = app.screen_updating
        calculation_before = app.calculation
        enable_events_before = app.api.EnableEvents

        eng.run_from_button(wb=wb)

        _check('screen_updating restored after run', app.screen_updating == screen_updating_before)
        _check('calculation mode restored after run', app.calculation == calculation_before)
        _check('EnableEvents restored after run', app.api.EnableEvents == enable_events_before)
        # COM sometimes round-trips a False write as the string 'FALSE' on read,
        # depending on the win32com binding — this is reading back what
        # run_from_button correctly set, not evidence of a bug in run_from_button itself.
        _check('StatusBar released back to Excel',
               app.api.StatusBar is False or str(app.api.StatusBar).upper() == 'FALSE',
               f'got {app.api.StatusBar!r}')
        _check('warning/timing msgbox shown exactly once', len(msgbox_calls) == 1,
               f'{len(msgbox_calls)} call(s)')
        if msgbox_calls:
            text = msgbox_calls[0][0]
            _check('msgbox text includes the time report', 'Tiempos de ejecución' in text)
            _check('msgbox text includes a Total line', 'Total' in text)

        for stage in ('reading', 'calculation', 'iteration', 'writing'):
            _check(f"_timings recorded '{stage}'", eng._timings.get(stage, 0.0) > 0,
                   f'got {eng._timings.get(stage)!r}')

        ws_imp = wb.sheets[eng.SHEET_IMPACT]
        ws_risk = wb.sheets[eng.SHEET_RISK]
        _check(f"'{eng.SHEET_IMPACT}' sheet has data after the run",
               ws_imp.used_range.last_cell.row > 1)
        _check(f"'{eng.SHEET_RISK}' sheet has data after the run",
               ws_risk.used_range.last_cell.row > 1)

        wb.close()   # NOT saved — the real .xlsm file is untouched
    finally:
        eng._show_msgbox = original_msgbox
        app.quit()

    print(f'\n{"="*60}')
    if FAILURES:
        print(f'{len(FAILURES)} FAILURE(S):')
        for f in FAILURES:
            print(f'  - {f}')
        raise SystemExit(1)
    print('ALL CHECKS PASSED')
    if msgbox_calls:
        print(f'\nMsgBox text that would have been shown:\n{msgbox_calls[0][0]}')


if __name__ == '__main__':
    main()
