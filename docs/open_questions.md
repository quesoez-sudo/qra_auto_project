# QRA Tool — Open Questions & Unresolved Items

This document tracks questions that arose during analysis that are either unresolved,
deferred, or need confirmation before finalizing the calculation engine.

---

## 1. Fireball Scenario Mismatch

**Status:** Unresolved

**Issue:** Core sheet has only 2 scenarios with P_FB > 0:
- `D7203/ST1/HYD_FL14/NA/1mDia` (INST size)
- `D7203/ST1/HYD_FL14/NA/1mNoche` (INST size)

The Directions sheet assigns `$B$1:$P$31` (first 30 rows) of ImpactThermMatrix to Fireball.
However, none of those 30 rows match the D7203/ST1 scenarios — and D7203/ST1 does not appear
anywhere in the first 30 rows (checked via Python).

**Question:** Should D7203/ST1 INST scenarios have fireball radiation distances in
ImpactThermMatrix rows 2–31? Are they accidentally missing from ImpactThermMatrix,
or is the fireball calculated differently (e.g. a separate consequence sheet)?

---

## 2. BLEVE Scenario Data Range = 1 Row

**Status:** Unresolved

**Issue:** Directions row 8 (BLEVE Results) has Scenario Data Range = `$B$1:$H$1`.
This covers only the header row of ImpactExpMatrix — 0 data rows.

**Question:** Is BLEVE not yet populated in the kernel workbook? Should this range be
`$B$1:$H$N` for some N > 1? Currently BLEVE will produce a zero matrix.

---

## 3. ImpactToxMatrix Computed Table — Unknown Column Titles

**Status:** Unresolved (noted in ImpactToxMatrix_formula_analysis.md as `[NEEDS TITLE]`)

**Issue:** The J:N computed table inside ImpactToxMatrix has 5 columns. Only J (distance,
descending) and M (probability) are used by the AB1 formula. Columns K, L, N are computed
but their meaning is unknown.

**Question:** What do columns K, L, N of the J:N table represent?

---

## 4. Identical Thermal Impact Matrices for JF / LPF / EPF

**Status:** Accepted — noted for documentation

**Finding:** JF, LPF, and EPF all read from ImpactThermMatrix. The same 108 Core scenarios
have non-zero probability for all three events. The radiation distance data is identical
per scenario regardless of fire type.

**Result:** The three Impact matrices (radiation kW/m² consequence) will be numerically identical.
The three Risk matrices will differ only in magnitude because P_JF << P_LPF ≈ P_EPF per scenario.

**Decision:** Produce separate matrices for each event as designed. Accepted as correct behavior.

---

## 5. ImpactThermMatrix Scenarios Not in Core

**Status:** Noted — not blocking

**Finding:** 2 scenario keys appear in ImpactThermMatrix but not in Core:
- `E7210_carcasa/L1/MVN_FL52/H/1mDia`
- `E7210_carcasa/L1/MVN_FL52/H/1mNoche`

**Question:** Are these intentionally excluded from the current Core scope? Should they
have entries added to Core, or should they be skipped silently?

---

## 6. Directions KEY (id-Nsheet) Column Meaning

**Status:** Unresolved

**Issue:** Directions column D ("KEY (id-Nsheet)") contains values 16–24 (one per event row).
These do NOT directly correspond to Core column indices (probability columns are in cols 10–17).
The offset appears to be +6 vs. Core column index.

**Question:** What does the KEY column represent? Is it the kernel tool's internal scenario
number for VBA lookups? It does not appear to be used for Python column selection (we use
the probability column name from the Core header instead).

---

## 7. Directions "Prob cell Kernel" Column = `$J$4` Everywhere

**Status:** Unresolved

**Issue:** Every row in Directions has `$J$4` for "Prob cell Kernel" (column 16).
This appears to be a placeholder reference to cell J4 in the kernel workbook (VBA usage).

**Question:** Is `$J$4` used by VBA only and irrelevant for Python, or should Python
read something from J4 of the kernel tool workbook?

---

## 8. Late Explosion — Formula and Column Range in ImpactFFMatrix

**Status:** Deferred — Late Explosion outputs zero matrices until resolved

**Finding:** Directions row 10 (Late Explosion) maps to ImpactFFMatrix with column
range `$B$1:$S$60` (19 columns B–S), while Flash Fire uses only `$B$1:$F$133` (6 columns B–F).
Currently there are 0 scenarios with Late Explosion data in ImpactFFMatrix.

**Decision:** Late Explosion impact and risk matrices are set to zero for now.

**Question:** When Late Explosion scenarios are added to ImpactFFMatrix, do columns G–S
contain different parameters than the FF columns (B–F), and does the formula differ
from the Flash Fire LFL-radius formula?

---

## 9. QY Grid Dimension — Non-Integer

**Status:** Minor — noted

**Finding:** Directions row 17 (Y dimension) shows Map grid qty = 317.46 (not an integer).
QX is exactly 315. We use QY = 317 (floor) for the grid.

**Question:** Should QY = 317 or 318? Using 317 drops ~0.46 cells (≈0.49 m) from the bottom edge.

---

## 10. Core Probability Column for Late Explosion

**Status:** Deferred — Late Explosion outputs zero matrices until resolved

**Issue:** Core has 8 probability columns (P_TOXIC through P_FF = cols J–Q).
The 9th event (Late Explosion, Directions id=10) has no corresponding named column
in Core's row 2 header. Currently there are 0 Late Explosion scenarios in ImpactFFMatrix.

**Decision:** Late Explosion risk matrix is set to zero for now.

**Question:** Does Late Explosion have its own probability column in Core (beyond Q),
or does it reuse P_FF? If Core needs a new column, what is its name and how is it populated?

---

---

## 11. Impact Zone Positional Shift (~5 cells in X and Y)

**Status:** ⚠️ Accepted — not blocking, within QRA spatial uncertainty

**Finding:** Impact zones computed by the Python engine appear shifted approximately 5 cells
(~5.35 m) south and east of the source equipment location as seen on the facility map.
Zone shape and magnitude are correct; only the absolute position is slightly off.

**Root cause:** PHAST derives source coordinates from precise CAD/point geometry with
sub-cell accuracy. Core stores a numeric approximation of those coordinates. Python then
maps the Core value to the nearest grid cell centre (discrete snap), accumulating the
positional error. Grid X/Y offsets in the Directions and General sheets are confirmed
zero, so no systematic offset fix is possible from the workbook alone.

**Magnitude:** ~5 cells × SX/SY ≈ 5.35 m on a 337 m × 339 m grid = ~1.6% positional error.

**Decision:** Acceptable for QRA. Risk magnitudes are unaffected. Flag for reference;
do not attempt a fix until the exact PHAST coordinate derivation method is known.

---

## Summary Table

| # | Topic | Status |
|---|-------|--------|
| 1 | Fireball scenario mismatch in ImpactThermMatrix | ❓ Unresolved |
| 2 | BLEVE scenario range = 1 row (empty) | ❓ Unresolved |
| 3 | ImpactToxMatrix columns K, L, N meaning | ❓ Unresolved |
| 4 | Identical thermal impact matrices JF/LPF/EPF | ✅ Accepted |
| 5 | ImpactThermMatrix E7210 scenarios not in Core | ❓ Low priority |
| 6 | Directions KEY column meaning | ❓ Unresolved (not blocking) |
| 7 | Directions "Prob cell Kernel" = `$J$4` everywhere | ❓ Not blocking |
| 8 | Late Explosion formula/columns in ImpactFFMatrix | ⏸️ Deferred — zero output |
| 9 | QY = 317 vs 318 (non-integer grid) | ⚠️ Minor |
| 10 | Late Explosion probability column in Core | ⏸️ Deferred — zero output |
| 11 | Impact zone positional shift ~5 cells in X and Y | ⚠️ Accepted — within QRA tolerance |
