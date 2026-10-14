"""
qra_button_entry.py
====================
PyInstaller entry point for the frozen IRiskMap button executable. Frozen
separately from qra_engine_v2_1.py so the module's own CLI entry point
(python qra_engine_v2_1.py -> run_live()) stays untouched.

When launched by xlwings' RunFrozenPython, Excel appends
--wb=... --from_xl=1 --app=... --hwnd=... to argv; run_from_button(wb=None)
resolves the calling workbook via xw.Book.caller(), which reads those same
argv flags.

Usage (built into a standalone .exe via PyInstaller, invoked from VBA):
  RunFrozenPython "C:\\path\\to\\QRAEngine.exe"
"""
import qra_engine_v2_1 as eng

if __name__ == '__main__':
    eng.run_from_button()
