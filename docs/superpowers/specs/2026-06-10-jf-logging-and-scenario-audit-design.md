# Design: JF Parameter Logging + Scenario Audit Diagnostic

**Date:** 2026-06-10  
**Status:** Approved

---

## Problem

The Python engines produce impact/risk matrices with values lower than Excel's native
computation for several events (TOXIC, LPF, EPF, etc.).  Locations appear correct but
magnitudes are too low, which is consistent with **missing scenario contributions** —
Python processes fewer scenarios than Excel does for the same event.

---

## Changes

### 1. JF Parameter Logging (both engines)

**`qra_engine.py`** — add a print block in `main()` after the existing grid/output print,
showing the four JF constants so they can be verified visually at startup:

```
JF formula params : directions=8  angle_offset=0.0°  angle_step=0  thresholds=[1.6, 5.0, ...]
```

**`qra_v6_engine.py`** — extend the existing `print_params()` function with a JF block
after the `Thermal t_exp` line, reading from `_P`:

```
Jet Fire dirs     : 8  offset=0.0°  step=0.0°
Jet Fire kW/m²    : [1.6, 5.0, 7.3, ...]
```

---

### 2. `tests/scenario_audit.py` — Scenario Audit + Contribution Log

A standalone diagnostic script that reads the KernelV0 workbook (openpyxl, no Excel
required) and reports on scenario coverage and per-scenario impact contributions.

#### Configuration block (top of file)

| Variable | Purpose |
|---|---|
| `_WORKSPACE` | Folder containing `KernelV0 (version 1).xlsx` |
| `RUN_CONTRIBUTION_LOG` | `True` to run Section B (slower, runs formulas) |
| `EVENTS_FILTER` | List of event names to audit; empty = all events |

#### Section A — Inventory Audit (always runs)

For each event:
- Total rows in its impact sheet
- Matched to Core count
- Skipped count (key not found in Core) + full list of skipped keys
- Size distribution of matched scenarios (`S/M/L/XL/INST/unknown`)
- Probability column stats: min, max, non-zero count, zero count

This directly answers: *"Is Python reading all the scenarios Excel uses?"*

#### Section B — Per-Scenario Contribution Log (controlled by `RUN_CONTRIBUTION_LOG`)

For each matched scenario per event, runs the formula and prints:

| Column | Description |
|---|---|
| Event | Event name |
| Key | Scenario key string |
| Size | Size label |
| Prob | Probability value from Core |
| MaxImpact | Max cell value in the scenario's impact matrix |
| SumImpact | Sum of all cells in the impact matrix |
| NonZeroCells | Count of cells with impact > 0 |

Output is printed as a formatted table and optionally saved to
`output/scenario_audit_B.csv` for spreadsheet inspection.

This directly answers: *"Is each matched scenario contributing what it should?"*

---

## What this does NOT cover

- Formula correctness verification (Approach C from brainstorming) — reserved for
  when A and B are clean but mismatch persists.
- v6 engine / MacroQRAV6 workbook — audit targets KernelV0 only.

---

## Files changed

| File | Change |
|---|---|
| `qra_engine.py` | Add JF parameter print block in `main()` |
| `qra_v6_engine.py` | Extend `print_params()` with JF block |
| `tests/scenario_audit.py` | New file |
