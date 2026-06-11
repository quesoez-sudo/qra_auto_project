# Design: Excel vs Python Matrix Comparison Test

**Date:** 2026-06-10  
**Status:** Approved

---

## Problem

Python impact/risk matrices produce lower values than Excel's native computation
for several event types. The per-scenario contribution diagnostic (`tests/scenario_audit.py`)
checks what Python is reading and computing, but does not compare against Excel's
actual formula output. This test does the direct per-scenario comparison.

---

## Approach

Live xlwings comparison: for N sampled scenarios per event, set the scenario
parameters in the sheet's AK/AL control cells, force Excel to recalculate, read
the output matrix (AM1:MO317), run Python's formula with the same inputs, and
compute quality metrics.

**Quality metrics (both):**
- `match_rate`: fraction of cells where |Excel − Python| < tolerance
- `max_abs_error`, `mean_abs_error`: error magnitude
- `nonzero_xl`, `nonzero_py`: non-zero cell counts (shows if one side is producing zeros where the other doesn't)

---

## AK/AL Parameter Cell Layout

All impact sheets share the same AK/AL structure.  Labels are in col AK; values in col AL.

| Row | Label | Usage |
|-----|-------|-------|
| 1 | `impact id` | identifier — do not change |
| 2 | `Impact/Risk` | 0 = impact only, 1 = impact × probability — **set to 0** for this test |
| 3 | `Scenario` | full ScenarioWeather key (matches col B of the data rows) |
| 4 | `Weather` | weather suffix (e.g. `Dia` or `Noche`) |
| 5 | `X` | source X coordinate [m] (or ignition X for CVE) |
| 6 | `Y` | source Y coordinate [m] (or ignition Y for CVE) |
| 7 | `Probability` | set to 0.0 for impact-only comparison |
| 11 | `SX` | grid cell width [m] |
| 12 | `SY` | grid cell height [m] |
| 13 | `QX` | grid columns (315) |
| 14 | `QY` | grid rows (317.46 → floor = 317) |
| 15–19 | `limit1`–`limit5` | event-specific thresholds (read-only for this test) |
| 23+ | event params | e.g. JF: index, angle offset, directions, step |

AK labels are read dynamically (no hardcoded row numbers) so the script is robust
to future row additions.

---

## Output Matrix Range

All sheets share the same output range: **AM1:MO317**
- AM = column 39; MO = column 353 (39 + 315 − 1)
- Rows 1–317 = QY rows

Grid params from AK/AL are cross-checked against Python's module constants at startup.

---

## Event Catalogue

| Event | Sheet | Python formula |
|-------|-------|---------------|
| TOXIC | ImpactToxMatrix | `formula_toxic(dist, tox_dists, tox_probs)` |
| JF | ImpactJFMatrix | `formula_jf(sx, sy, dist_vals, halfW_vals, center_vals)` |
| LPF | ImpactThermMatrix | `formula_thermal(dist, therm_dists)` |
| EPF | ImpactThermMatrix | `formula_thermal(dist, therm_dists)` |
| FB | ImpactThermMatrix | `formula_thermal(dist, therm_dists)` |
| CVE | ImpactExpMatrix | `formula_explosion(dist_grid(ign_x, ign_y), exp_dists)` |
| BLV | ImpactExpMatrix | `formula_explosion(dist_grid(sx, sy), exp_dists)` |
| FF | ImpactFFMatrix | `formula_ff(dist, lfl_dist, lflf_dist)` |
| LATE_EXP | — | skip (zero matrix) |

---

## Files

| File | Change |
|------|--------|
| `docs/superpowers/specs/2026-06-10-compare-excel-design.md` | this spec |
| `tests/compare_excel.py` | new comparison test script |
