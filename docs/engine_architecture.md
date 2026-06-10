# QRA Engine — Architecture & Logic Reference

**Purpose:** Full documentation of the Python calculation engine that replicates the
Excel QRA kernel workbook's impact and risk matrix computations.

---

## 1. What the Engine Does

For each combination of **event type** (9) × **leak size** (6), the engine:

1. Loads consequence distances from the relevant ImpactXXMatrix Excel sheet
2. Matches scenarios against Core (source of coordinates and probabilities)
3. Computes a **317 × 315 impact matrix** (consequence probability per grid cell)
4. Computes a **317 × 315 risk matrix** = impact × frequency per grid cell
5. Saves results as CSV files

---

## 2. Input Data Sources

### 2.1 Kernel Workbook
`KernelV0 (version 1).xlsx` — read with `openpyxl`, `data_only=True` (cached values, no VBA).

### 2.2 Core Sheet
Master list of all scenarios. One row per scenario.

| Column | Letter | Content |
|--------|--------|---------|
| A | A | Scenario key (ScenarioWeather) — match key |
| D | D | X coordinate (m) of source |
| E | E | Y coordinate (m) of source |
| G | G | Leak size: S / M / L / XL / INST |
| J | J | P_TOXIC — frequency (1/yr) |
| K | K | P_JF — Jet Fire frequency |
| L | L | P_LPF — Late Pool Fire frequency |
| M | M | P_EPF — Early Pool Fire frequency |
| N | N | P_FB — Fireball frequency |
| O | O | P_CVE — CVE frequency |
| P | P | P_BLV — BLEVE frequency |
| Q | Q | P_FF — Flash Fire frequency |

Data starts at row 4 (row 2 = headers, row 3 = blank/separator).

### 2.3 Directions Sheet
Event routing map. One row per event (rows 2–10).

| Column | Content |
|--------|---------|
| A | Event id |
| B | Description |
| C | Effect sheet key (1=Tox, 2=Therm, 3=Exp, 4=FF) |
| E–J | Output ranges in ImpactMatrix0/RiskMatrix0 (Total, S, M, L, XL, INST) |
| K | Scenario data range in ImpactXXMatrix ← **ignored per design decision** |
| Q | Impact sheet name (e.g. ImpactThermMatrix) |
| T | Result data range in ImpactXXMatrix (e.g. `$AB$1:$MD$317`) |

### 2.4 ImpactXXMatrix Sheets
One sheet per consequence type. All scenarios, all rows used.

| Sheet | Events | Scenario key col | X/Y coords | Formula params |
|-------|--------|-----------------|------------|----------------|
| ImpactToxMatrix | Toxic | B | — (no grid) | AA cols |
| ImpactThermMatrix | JF, LPF, EPF, FB | B | — | AA cols |
| ImpactExpMatrix | CVE, BLV | B | V (ignition X), W (ignition Y) for CVE | AA cols |
| ImpactFFMatrix | FF, Late Exp | B | — | AA cols |

**Coordinate source:**
- Most events: X/Y from Core (D, E columns) = leak source location
- CVE only: X/Y from ImpactExpMatrix columns V/W = ignition point coordinates

---

## 3. Map Grid

| Parameter | Value | Source |
|-----------|-------|--------|
| QX (cols) | 315 | Directions row 16, col E |
| QY (rows) | 317 | Directions row 17, col E (truncated from 317.46) |
| SX (m/cell) | 1.0698412698… | Directions row 16, col D |
| SY (m/cell) | 1.069425 | Directions row 17, col D |
| X offset | 0 | Directions row 16, col C |
| Y offset | 0 | Directions row 17, col C |

Grid cell centers:
```
x_centers = X_offset + SX * (0.5 + arange(QX))   # shape (315,)
y_centers = Y_offset + SY * (0.5 + arange(QY))   # shape (317,)
```

Distance from a source point (sx, sy) to every cell:
```
dist[j, i] = sqrt((x_centers[i] - sx)^2 + (y_centers[j] - sy)^2)
```

---

## 4. Size Filtering

Core column G contains the leak size label for each scenario.
Six accumulation passes are run:

| Pass | Filter condition |
|------|----------------|
| Total | All scenarios (no filter) |
| S | G == 'S' |
| M | G == 'M' |
| L | G == 'L' |
| XL | G == 'XL' |
| INST | G == 'INST' |

---

## 5. Formula Implementations

### 5.1 Flash Fire (FF) — ImpactFFMatrix
**Events:** Flash Fire (FF)  
**Source data columns in sheet:** B (key), C (scenario code), D (weather), F (LFL dist m), G (LFL fraction dist m), H (X source), I (Y source)  
**Parameters from sheet AA column:** limit1=0, limit2=1, limit3=2, plus grid params

**Formula per cell:**
```
dist = euclidean distance from source to cell center
if dist <= LFL_dist (col F):      impact = limit3 = 2  (inside LFL = lethal)
elif dist <= LFLF_dist (col G):   impact = limit2 = 1  (transition zone)
else:                              impact = limit1 = 0  (outside)
```

Impact matrix accumulates the impact value; Risk = impact × scenario probability.

> Note: Late Explosion uses ImpactFFMatrix with the same formula.
> Currently 0 Late Explosion scenarios exist → zero matrices. See open_questions.md #8, #10.

---

### 5.2 Thermal Radiation (Therm) — ImpactThermMatrix
**Events:** JF (P_JF), LPF (P_LPF), EPF (P_EPF), FB (P_FB)  
**Source data columns in sheet:** B (key), G–P (distances for 10 radiation thresholds, descending)  
**Parameters from sheet AA column:** 10 kW/m² thresholds, exposure time (AA8)

**Radiation thresholds (kW/m²):** 1.6, 5, 7.3, 9.5, 12.5, 16, 20.9, 25, 30, 35  
**Columns G–P:** distances (m) corresponding to each threshold (G=farthest/lowest, P=closest/highest)

**Formula per cell:**
```
dist = euclidean distance from source to cell center
# linear interpolation to get kW/m² at that distance
kW = interpolate(dist, distances[G:P], thresholds)
# probit equation
probit = -36.38 + 2.56 × ln((1000 × kW)^(4/3) × t_exp)
# convert probit to probability via normal CDF
impact = Φ(probit - 5)       where Φ = standard normal CDF
impact = clamp(impact, 0, 1)
```

Returns impact in [0, 1]; Risk = impact × scenario probability.

---

### 5.3 Explosion Overpressure (Exp) — ImpactExpMatrix
**Events:** CVE (P_CVE), BLV (P_BLV)  
**Source data columns in sheet:** B (key), J–N (distances for 5 overpressure thresholds), V (ignition X), W (ignition Y)  
**Parameters from sheet AA column:** limitOV1, limitOV2, limitF1, limitF2, thresholds

**Overpressure thresholds (bar):** 0.04, 0.1, 0.35, 0.5, 1.0  
**Columns J–N:** distances (m) at each threshold (J=farthest, N=closest)

**Coordinate source:**
- CVE: X/Y from ImpactExpMatrix cols V/W (ignition point, one per scenario row)
- BLV: X/Y from Core cols D/E (leak source)

**Formula per cell:**
```
dist = euclidean distance from ignition/source to cell center
# linear interpolation to get overpressure (bar) at that distance
v = interpolate(dist, distances[J:N], thresholds)
# step function
if v < limitOV1 (0.1 bar):         impact = 0
elif v < limitOV2 (0.3 bar):       impact = limitF1 (0)
else:                               impact = limitF2 (1)
```

> Note: BLV currently has 0 data rows in ImpactExpMatrix → zero matrices. See open_questions.md #2.

---

### 5.4 Toxic Dispersion (Tox) — ImpactToxMatrix
**Events:** Toxic (P_TOXIC)  
**Source data columns in sheet:** B (key), G (multi-line CSV blob of dispersion profile)  
**Parameters from sheet AA column:** AA15 = minimum probability filter threshold

**Each scenario has a CSV blob in column G containing a dose–distance profile:**
- Column J: distance (m, descending)
- Column M: probability of harm at that distance

**Formula per cell:**
```
dist = euclidean distance from source to cell center
# linear interpolation on (J, M) columns of parsed profile
impact = interpolate(dist, distances[J], probabilities[M])
impact = clamp(impact, 0, 1)
```

---

## 6. Accumulation Logic

For each event and size filter:

```python
impact_matrix = zeros(QY, QX)   # 317 × 315
risk_matrix   = zeros(QY, QX)

for scenario in matched_scenarios:
    if scenario.size != target_size (and target != Total): skip
    cell_impact = formula(scenario, grid)          # shape (317, 315)
    impact_matrix += cell_impact
    risk_matrix   += cell_impact * scenario.probability
```

**Matching rule:** Scenario key from ImpactXXMatrix col B must exist in Core col A.
If not found → skip. If probability = 0 → impact still computed (for impact matrix),
risk contribution = 0.

---

## 7. Event Configuration Table

| Event | id | Sheet | Prob col (Core) | Formula | Special |
|-------|----|-------|-----------------|---------|---------|
| Toxic | 2 | ImpactToxMatrix | J (P_TOXIC) | Toxic interp | — |
| JF | 3 | ImpactThermMatrix | K (P_JF) | Thermal probit | — |
| LPF | 4 | ImpactThermMatrix | L (P_LPF) | Thermal probit | — |
| EPF | 5 | ImpactThermMatrix | M (P_EPF) | Thermal probit | — |
| FB | 6 | ImpactThermMatrix | N (P_FB) | Thermal probit | — |
| CVE | 7 | ImpactExpMatrix | O (P_CVE) | Exp step | Ignition X/Y from cols V/W |
| BLV | 8 | ImpactExpMatrix | P (P_BLV) | Exp step | 0 rows → zeros |
| FF | 9 | ImpactFFMatrix | Q (P_FF) | FF LFL radius | — |
| Late Exp | 10 | ImpactFFMatrix | — | — | Deferred → zeros |

---

## 8. Output Layout

### 8.1 CSV files (verification)
Location: `output/`  
Naming: `impact_{event}_{size}.csv` and `risk_{event}_{size}.csv`  
Example: `impact_JF_Total.csv`, `risk_LPF_M.csv`  
Shape: 317 rows × 315 columns

### 8.2 Excel output (future)
Target sheets: `ImpactMatrix0`, `RiskMatrix0`  
Layout (from Directions cols E–J):
- Rows: each event block is 317 rows; events stack vertically (row offsets in Directions)
- Cols: each size block is 315 cols wide; sizes stack horizontally (col offsets in Directions)

---

## 9. Key Design Decisions

| Decision | Rationale |
|----------|-----------|
| Ignore Directions K column (scenario data range) | Use all rows of each ImpactXXMatrix; P=0 contributes 0 naturally |
| Skip ImpactXXMatrix scenarios not in Core | No coordinates available; only Core scenarios contribute |
| Late Explosion → zero matrices | No scenarios exist yet; formula/probability TBD (see open_questions.md) |
| BLV → zero matrices | ImpactExpMatrix has 0 BLEVE rows (see open_questions.md) |
| CVE uses ignition coordinates | ImpactExpMatrix cols V/W, not Core D/E |
| QY = 317 (truncated) | Floor of 317.46 from Directions; drops <0.5m at bottom edge |
| Output as CSV first | Verify before committing to large Excel write |

---

## 10. Files

| File | Purpose |
|------|---------|
| `qra_engine.py` | Main calculation engine |
| `ff_risk_matrix.py` | Original FF prototype (reference) |
| `open_questions.md` | Unresolved items / deferred features |
| `engine_architecture.md` | This file |
| `output/*.csv` | Computed matrices |
| `KernelV0 (version 1).xlsx` | Source Excel workbook |
