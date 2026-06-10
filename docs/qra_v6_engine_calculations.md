/# QRA V6 Engine — Calculations & Architecture

## Overview

`qra_v6_engine.py` computes **Impact** and **Risk** matrices on a `QY × QX` grid for up to 8 consequence event types. It drives Excel via `xlwings` (COM), reads PHAST-computed consequence distances, applies fatality formulas, and writes the results back to the workbook.

---

## 1. Grid Definition

All spatial calculations take place on a regular 2-D grid whose dimensions and cell sizes come from the `General` sheet at runtime.

| Symbol | Default | Source (General sheet) | Meaning |
|--------|---------|------------------------|---------|
| `QX` | 315 | S5 | Number of cells in X direction |
| `QY` | 317 | S6 | Number of cells in Y direction |
| `SX` | ~1.0698 m | R5 | Cell width [m] |
| `SY` | ~1.069425 m | R6 | Cell height [m] |

**Cell-centre coordinates** (NumPy meshgrid, shape `(QY, QX)`):

```
x_i = SX × (i + 0.5)   for i in 0 … QX-1
y_j = SY × (j + 0.5)   for j in 0 … QY-1
XX, YY = meshgrid(x, y)
```

**Euclidean distance** from source point `(sx, sy)` to every grid cell:

```
D[j, i] = sqrt((XX[j,i] − sx)² + (YY[j,i] − sy)²)
```

---

## 2. Event Types & Configuration

Eight consequence events are processed, each with its own result sheet and fatality formula:

| Impact ID | Event | Formula | Frequency column (Core) |
|-----------|-------|---------|------------------------|
| 16 | TOXIC — Outdoor Toxic Results | `toxic` | col 9 |
| 17 | JF — Jet Fire Results | `thermal` | col 10 |
| 18 | LPF — Late Pool Fire Results | `thermal` | col 11 |
| 19 | EPF — Early Pool Fire Results | `thermal` | col 12 |
| 20 | FB — Fireball Results | `thermal` | col 13 |
| 21 | CVE — CVE Results | `explosion` (+ ignition XY) | col 14 |
| 22 | BLV — BLEVE Results | `explosion` | col 15 |
| 23 | FF — Flash Fire Results | `ff` | col 16 |

---

## 3. Fatality Formulas

### 3.1 Thermal Radiation (`thermal`)

Used for Jet Fire, Pool Fires, and Fireball.

**Step 1 — Interpolate kW/m²**

PHAST provides distances at which the radiation level equals each threshold (e.g. 1, 4, 12.5, 37.5 kW/m²). For every grid cell, the engine interpolates the radiation intensity from those distance–threshold pairs:

```
kw_at_cell = interp(D[cell], distances_descending, kW_thresholds_descending)
             (left extrapolation = max kW, right = 0 kW)
```

**Step 2 — Thermal probit**

```
Probit = −36.38 + 2.56 × ln( (1000 × kw)^(4/3) × t_exp )
```

- `kw` — radiation intensity [kW/m²]
- `t_exp` — thermal exposure time [s] (from General sheet D21)
- `1000` converts kW → W inside the probit argument

**Step 3 — Fatality probability via normal CDF**

```
P_fatality = Φ( Probit − 5.0 )
```

`Φ` is the standard normal CDF (Abramowitz & Stegun approximation, error < 1.5×10⁻⁷):

```python
t = 1 / (1 + 0.3275911 × |x|)
y = 1 − polynomial(t) × exp(−x²)
Φ(x) = 0.5 × (1 + sign(x) × y)
```

Cells at or beyond the maximum threshold are clamped to `P = 1.0`.

---

### 3.2 Explosion Overpressure (`explosion`)

Used for CVE and BLEVE.

**Step 1 — Interpolate overpressure [bar]**

Same interpolation mechanics as thermal, but on overpressure distances:

```
bar_at_cell = interp(D[cell], distances_descending, bar_thresholds_descending)
```

**Step 2 — Step-function fatality**

```
P_fatality = 1.0   if bar_at_cell >= 0.3 bar   (limitOV2)
           = 0.0   if bar_at_cell >= 0.1 bar   (limitOV1) — intermediate zone, factor = 0
           = 0.0   if bar_at_cell < 0.1 bar
```

> Note: `_EXP_F_MID = 0.0` — the intermediate zone currently contributes zero fatality. Verify with the project team.

**CVE special case:** uses the ignition point coordinates `(X, Y)` from columns V/W of the CVE Results sheet instead of the source point from Core. A single scenario may have multiple ignition entries; each is computed separately and accumulated.

---

### 3.3 Flash Fire (`ff`)

Two radii per scenario come from the Flash Fire Results sheet:

- `lfl_dist` — distance to LFL (Lower Flammable Limit) boundary [m]
- `lflf_dist` — distance to LFL Fraction boundary [m] (≤ lfl_dist)

Zone assignment per grid cell:

```
result = FF_OUTSIDE    (default = 0)
result = FF_TRANSITION  where D ≤ lflf_dist  (default = 1)
result = FF_INSIDE_LFL  where D ≤ lfl_dist   (default = 2)
```

Zone values come from General sheet row 18.

---

### 3.4 Toxic Dispersion (`toxic`)

Each scenario row in the Outdoor Toxic Results sheet contains a compressed multi-line CSV blob (column G) encoding a distance–probability profile.

**Parsing:**
1. Split blob by newlines; each line: `distance, …, …, probability`
2. Filter: keep rows where `distance ≥ 0` and `probability ≥ TOX_MIN_PROB`
3. Sort ascending by distance

**Interpolation:**

```
P_fatality = interp(D[cell], tox_distances, tox_probabilities)
             (left = max probability, right = 0)
```

Clipped to [0, 1].

---

## 4. Accumulation: Impact & Risk Matrices

For each event, 6 size-filtered variants are computed simultaneously: `Total`, `S`, `M`, `L`, `XL`, `INST`.

For every scenario in Core that matches a row in the result sheet:

```
cell_impact = formula(D, consequence_data)     # shape (QY, QX)
cell_risk   = cell_impact × frequency          # frequency from Core prob column

impact_matrix['Total'] += cell_impact
risk_matrix['Total']   += cell_risk

if scenario.size in ['S','M','L','XL','INST']:
    impact_matrix[size] += cell_impact
    risk_matrix[size]   += cell_risk
```

`'Total'` accumulates **all** scenarios regardless of size tag.

---

## 5. Pipeline (Step by Step)

```
1. Open MacroQRAV6 (version 1).xlsm via xlwings (hidden Excel instance)

2. Load General sheet parameters
   → QX, QY, SX, SY
   → Thermal kW/m² thresholds
   → Overpressure bar thresholds
   → FF zone values
   → Toxic min probability
   → Thermal exposure time [s]

3. Build XX, YY meshgrid from grid parameters

4. Read PageControl (cols J–R, rows 2–500)
   → Active Impact IDs
   → Destination cell ranges per ID per size

5. For each active Impact ID:

   5a. Set Core!C2 = impact_id → trigger Excel recalculation
       (wait 1.5s + Calculate() + wait 1.0s)

   5b. Warn if Core!G2 has a size filter active

   5c. Read Core!A4:R89 → list of scenario dicts
       {key, size, x, y, frequencies…}
       Stop at sentinel: None / 0 / -1 in column A

   5d. Read consequence result sheet (dynamic column matching):
       - Thermal: match row-1 kW/m² headers to General thresholds
       - Explosion: match row-1 bar headers; fallback to positional cols J–N
       - Flash Fire: fixed cols E (LFL) and F (LFLF)
       - Toxic: parse CSV blob in col G

   5e. compute_event():
       For each scenario × each ignition entry:
         → compute distance grid D
         → apply formula → cell_impact (QY×QX)
         → cell_risk = cell_impact × frequency
         → accumulate into 6 size-filtered matrices

6. Write Impact Matrix Result and Risk Matrix Result sheets
   → Excel events and auto-calc disabled during write
   → Bulk COM write (single Resize().Value = … call per block)
   → One block per (impact_id, size) pair

7. SaveAs .xlsx (macro-disabled export)

8. Print warning/error summary
```

---

## 6. Dynamic Column Matching

Rather than assuming fixed column positions, the engine reads **row 1 headers** of each result sheet and matches them numerically to the thresholds loaded from General.

**Thermal:** headers are the actual kW/m² values. Unmatched thresholds produce `None` in the distance list and contribute nothing.

**Explosion:** headers should be bar values. If fewer than the expected number match (e.g., formula error returning `1/-1`), the engine falls back to **positional assignment** (columns J, K, L, M, N → thresholds in order) and logs a warning.

---

## 7. Output

Two sheets written to the workbook, then exported to `MacroQRAV6_export_result.xlsx`:

| Sheet | Contents |
|-------|----------|
| `Impact Matrix Result` | Fatality probability at each grid cell, per event × size |
| `Risk Matrix Result` | `impact × frequency`, per event × size |

Each matrix block is placed at the destination range specified in `PageControl` (e.g., `$LD$318:$XF$634`).

---

## 8. Key Constants

| Constant | Value | Meaning |
|----------|-------|---------|
| `_THERM_PA` | −36.38 | Probit intercept |
| `_THERM_PB` | 2.56 | Probit slope |
| `_THERM_DOE` | 4/3 | Thermal dose exponent |
| `_THERM_PROBIT_MEAN` | 5.0 | Probit 50% point |
| `_EXP_FATAL_BAR` | 0.3 | Overpressure for certain fatality [bar] |
| `_EXP_LOW_BAR` | 0.1 | Overpressure lower limit [bar] |
| `_NC_P` | 0.3275911 | Abramowitz & Stegun polynomial coefficient |

---

## 9. Known Limitations

- **BLEVE:** `BLEVE Results` sheet typically has zero rows → zero matrices.
- **Fireball:** Some scenarios may be missing from the result sheet (D7203/ST1).
- **Explosion intermediate zone:** `_EXP_F_MID = 0.0` — cells between 0.1–0.3 bar contribute zero fatality (verify with project team).
- **Grid floor:** `QY = floor(317.46)` — drops ~0.5 m from the bottom edge.
- **Late Explosion:** No formula implemented; currently not in `IMPACT_CONFIG`.
