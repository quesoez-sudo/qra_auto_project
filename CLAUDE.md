# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

Python automation layer for a Quantitative Risk Assessment (QRA) tool. The engine replicates the calculations that the Excel macro workbook (`MacroQRAV6 (version 1).xlsm`) performs, computing **317×315 impact and risk matrices** for each event type × 6 leak sizes (Total, S, M, L, XL, INST).

This repo is the **Windows/Excel-bound** generation. A web app replacing the whole workbook is being built separately (see "Migration" below).

## The Final Engine: `qra_engine_v2_1.py`

**`qra_engine_v2_1.py` is the current, production engine.** It reads everything from the production workbook: `General` (grid + thresholds), `Core` (scenarios, X/Y, frequencies, per-scenario JF/thermal params), `PageControl`, and the real `* Results` sheets. It writes per-size impact/risk matrices into the `CalcBUImpact` / `CalcBURisk` sheets.

### Entry points
| How | Call | Notes |
|---|---|---|
| **Excel button (production)** | `QRAEngine.exe` → `qra_button_entry.py` → `run_from_button()` | IRiskMap button in VBA calls xlwings `RunFrozenPython "...\QRAEngine.exe"`; resolves the open workbook via `xw.Book.caller()` |
| CLI, live Excel | `python qra_engine_v2_1.py --live` → `run_live()` | Opens `INPUT_FILE` in a hidden Excel with VBA disabled; writes `EXPORT_FILE` (never overwrites the `.xlsm`) |
| CLI, openpyxl snapshot | `python qra_engine_v2_1.py` → `main()` | Reads cached values; CSV output to `output_v2.1/` |

`run_live()` and `run_from_button()` share `_run_pipeline(wb, app)`. `load_general_params_live()` overwrites the module-level grid/threshold globals at the start of each run; the literals at the top of the file are only fallbacks for tests.

### Building the exe
```
pyinstaller QRAEngine.spec      # entry: qra_button_entry.py → dist/QRAEngine.exe
```
`QRAEngine.spec`, `build/` and `dist/` are not tracked in git. The exe is Windows-only (xlwings COM + `ctypes`).

### Event coverage (`EVENT_CONFIG`, impact IDs match `PageControl!L`)
| ID | Event | Sheet | Formula |
|---|---|---|---|
| 16 | TOXIC | Outdoor Toxic Results | Linear interpolation on CSV-blob distance/probability profile |
| 17 | JF | Jet Fire Results | **Elliptical** zones, per-scenario direction count |
| 18/19/20 | LPF / EPF / FB | Late/Early Pool Fire, Fireball Results | Circular distance → kW/m² → thermal probit → normal CDF |
| 21 | CVE | CVE Results | Step function at overpressure threshold, uses ignition coords (Xign/Yign) |
| 22 | BLV | BLEVE Results | **No reader yet** → zero |
| 23 | FF | Flash Fire Results | Step on LFL / LFLF radius → 0, 1, 2 |

## Older Engines (kept for reference, not production)

| File | What it is |
|---|---|
| `qra_engine.py` | First version. openpyxl, reads `KernelV0 (version 1).xlsx` (file no longer in the folder), routing from `Directions` |
| `qra_v6_engine.py` | xlwings. Sets `Core!C2` per Impact ID to force a live Core recalculation, exports `MacroQRAV6_export_result.xlsx` |
| `qra_engine_v2.py` | Kernel-pinned v2: reads `KernelV0_v2_copy.xlsx` / CoreExample. v2.1 branched from it |

## Architecture

### Computation Pattern
For each event, the engine:
1. Reads scenario rows from the event's Results sheet
2. Matches each row to `Core` by scenario key
3. Computes distance from the source point to every cell in the 317×315 grid
4. Applies the event formula to get a per-cell impact value
5. Accumulates: `impact_matrix += cell_impact`, `risk_matrix += cell_impact × frequency`
6. Repeats per size filter ("Total" uses all scenarios)

### Grid Constants (do not change without reading `docs/engine_architecture.md`)
- `QX=315`, `QY=317`, `SX≈1.0698 m/cell`, `SY≈1.069425 m/cell` (v2.1 reads them from `General` at runtime)
- Cell centers: `x = SX × (i + 0.5)`, `y = SY × (j + 0.5)`
- Open question: `General` gives Y = 339.5 m / 317 = 1.07098 m/cell, which disagrees with SY above

## Dependencies (`requirements.txt`)
- `numpy`: grid math
- `scipy`: `erf` for the normal CDF in v2.1 (older engines use an inline A&S approximation)
- `openpyxl`: snapshot reads
- `xlwings`: live Excel COM (Windows + Excel required for `--live` and the exe)

## Tests
- `tests/test_qra_v2_1_helpers.py`: unit tests for v2.1 helpers
- `tests/test_live_readers.py`, `tests/test_run_from_button_headless.py`: need Excel + the `...-Win11-cWvAMQq4D.xlsm` workbook at the project root; the headless test closes without saving
- `tests/compare_excel.py`, `tests/compare_thermal_all.py`: Python vs Excel matrix parity

## Known Issues
See `docs/open_questions.md` and `parameters_joint.md`. Key items:
- **BLEVE:** no reader in v2.1 (own column layout, psi units, uses Core X/Y not Xign/Yign)
- **Core snapshot:** v2.1 reads Core once; Core's scenario set is formula-driven per Impact ID (`qra_v6_engine.py` handles this by recalculating per ID)
- **Late Explosion:** no scenarios / formula yet
- **QY dimension:** source data gives 317.46; engine uses `floor(317)`

## Migration (web app)
The full workbook (inputs + PHAST + results) is being rebuilt as a web app, developed on macOS. The VBA source, form layouts and the starting prompt are exported to `QRA_Macro_Export/` (kept outside this repo: `MASTER_PROMPT.md`, `forms_viewer.html`, `vba/`). In that project, `qra_engine_v2_1.py` is the reference for Section 2 (results/matrices). Its calculation functions are portable; its xlwings/exe layer is not.

Excel files are git-ignored (and the `.xlsm` exceeds GitHub's 100 MB limit). Copy them manually.
