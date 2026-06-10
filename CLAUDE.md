# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What This Project Does

Python automation layer for a Quantitative Risk Assessment (QRA) tool. The engine replicates the calculations that the Excel macro workbook (`MacroQRAV6 (version 1).xlsm`) performs, computing **317×315 impact and risk matrices** for 9 event types × 6 leak sizes (Total, S, M, L, XL, INST).

## Running the Engines

There are two standalone scripts — choose based on the input workbook available:

**`qra_engine.py`** — uses `openpyxl` (no Excel required, reads cached `.xlsx` values):
```
python qra_engine.py
```
Before running, set `_WORKSPACE` at the top to the folder containing `KernelV0 (version 1).xlsx`. Outputs CSV files to `output/` and writes back to `ImpactMatrix0`/`RiskMatrix0` sheets inside the workbook.

**`qra_v6_engine.py`** — uses `xlwings` (requires Excel installed and open):
```
python qra_v6_engine.py
```
Before running, set `_WORKSPACE` at the top. Reads `MacroQRAV6 (version 1).xlsm`, sets `Core!C2` for each impact ID to force recalculation, then exports results to `MacroQRAV6_export_result.xlsx`.

## Architecture

### Data Flow
```
Excel workbook
  ├── Core sheet          → source coordinates (X/Y) + event frequencies per scenario
  ├── ImpactToxMatrix     → toxic dispersion profiles (CSV blob per scenario, col G)
  ├── ImpactThermMatrix   → radiation threshold distances (cols G–P, 10 thresholds)
  ├── ImpactExpMatrix     → overpressure threshold distances (cols J–N, 5 thresholds)
  ├── ImpactFFMatrix      → LFL/LFLF radii per scenario
  └── Directions/PageControl → destination cell ranges for result placement
```

### Computation Pattern
For each event, the engine:
1. Reads scenario rows from the relevant `ImpactXXMatrix` sheet
2. Matches each row to `Core` by scenario key (col A in Core, col B in effect sheets)
3. Computes Euclidean distance from the source point to every cell in the 317×315 grid
4. Applies a formula to get a per-cell impact value (see below)
5. Accumulates: `impact_matrix += cell_impact`, `risk_matrix += cell_impact × frequency`
6. Runs this accumulation 6 times (once per size filter; "Total" uses all scenarios)

### Formulas by Event Type

| Formula | Events | Key logic |
|---------|--------|-----------|
| `thermal` | JF, LPF, EPF, FB | Interpolate kW/m² from distance, apply probit (`-36.38 + 2.56 × ln(...)`), then normal CDF |
| `explosion` | CVE, BLV | Interpolate bar from distance, step function at `limitOV2=0.3 bar` → 0 or 1 |
| `ff` | FF | Step function on LFL radius and LFLF radius → values 0, 1, or 2 |
| `toxic` | TOXIC | Linear interpolation on distance-probability profile from CSV blob |
| `zero` | LATE_EXP | No scenarios; always zero matrices |

**CVE special case:** uses ignition point coordinates from `ImpactExpMatrix` cols V/W instead of the source coordinates from `Core`.

### Grid Constants (do not change without reading `engine_architecture.md`)
- `QX=315`, `QY=317`, `SX≈1.0698 m/cell`, `SY≈1.069425 m/cell`
- Cell centers: `x = SX × (i + 0.5)`, `y = SY × (j + 0.5)`

### Key Differences Between the Two Scripts
- `qra_engine.py`: reads routing from `Directions` sheet (rows 2–10, cols E–J); event config is hardcoded in `EVENTS` list
- `qra_v6_engine.py`: reads routing from `PageControl` sheet (cols L–R); event config in `IMPACT_CONFIG` dict (IDs 16–24); uses `xlwings` for live Excel COM interaction

## Dependencies

- `numpy` — all grid math
- `openpyxl` — Excel I/O in `qra_engine.py`
- `xlwings` — live Excel COM in `qra_v6_engine.py` (requires Excel installed)

No `scipy` — the normal CDF is implemented inline via Abramowitz & Stegun approximation in `_norm_cdf()`.

## Known Issues & Deferred Items

See `docs/open_questions.md` for the full list. Key blockers:
- **BLEVE (BLV):** `ImpactExpMatrix` has 0 BLEVE rows → zero matrices
- **Fireball (FB):** D7203/ST1 scenarios missing from ImpactThermMatrix rows 2–31
- **Late Explosion:** no scenarios yet; formula/probability column TBD
- **QY dimension:** source data gives 317.46; engine uses `floor(317)` — may drop ~0.5m bottom edge
