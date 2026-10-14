# QRA Engine — Parameters & Constants Reference

This document catalogs every parameter and constant used by the QRA Python engines,
maps each one to its source cell in the production workbook (`MacroQRAV6 (version 1).xlsm`),
and flags which values were hardcoded in the v2 development kernel vs read dynamically in v6.

---

## A. Grid / Geometry

These define the spatial extent of the impact and risk matrices.

| Parameter | Symbol | v2 (KernelV0_v2_copy) | MacroQRAV6 source | Notes |
|-----------|--------|-----------------------|-------------------|-------|
| Grid columns | `QX` | Hardcoded `315` | `General` S5 (col 19) | Number of east-west cells |
| Grid rows | `QY` | Hardcoded `317` | `General` T5 → row 5 col 20; row 6 col 20 for Y | Number of north-south cells |
| Cell width | `SX` (m/cell) | Hardcoded `1.0698412698412698` | `General` S5 (col 19) row 5 | x-axis scale |
| Cell height | `SY` (m/cell) | Hardcoded `1.069425` | `General` S6 (col 19) row 6 | y-axis scale |
| Cell-centre X | `x_m = SX*(col-0.5)` | Derived | Derived | col is 1-based |
| Cell-centre Y | `y_m = SY*(QY-row+0.5)` | Derived | Derived | row 1 = northmost |

**MacroQRAV6 General sheet cell map (v6 constant names):**

```
_GEN_QX_RC  = (5, 20)   # T5  — QX
_GEN_QY_RC  = (6, 20)   # T6  — QY
_GEN_SX_RC  = (5, 19)   # S5  — SX
_GEN_SY_RC  = (6, 19)   # S6  — SY
```

---

## B. Thermal Radiation (LPF, EPF, FB)

### B1. Impact thresholds (kW/m²)

| Parameter | v2 | MacroQRAV6 source |
|-----------|----|-------------------|
| `KW_THRESHOLDS` (10 values) | Hardcoded `[1.6, 5.0, 7.3, 9.5, 12.5, 16.0, 20.9, 25.0, 30.0, 35.0]` | `General` row 17, starting col J (up to 10 consecutive non-None values) |

```
_GEN_THERM_RC = (17, 10)   # J17 — first thermal threshold
```

v6 reads these dynamically and matches them to row-1 headers of the effect result sheets.
**These values are NOT hardcoded in v6** — they come from PHAST via General row 17.

### B2. Probit model constants (hardcoded — mathematical formula, not workbook data)

| Constant | Value | Formula role |
|----------|-------|-------------|
| `PA` (`_THERM_PA`) | `-36.38` | Probit intercept |
| `PB` (`_THERM_PB`) | `2.56` | Probit slope |
| `DOE` (`_THERM_DOE`) | `4/3` | Thermal dose exponent |
| `PROBIT_MEAN` (`_THERM_PROBIT_MEAN`) | `5.0` | 50 % probit point |
| `KW_SCALE` (`_THERM_KW_SCALE`) | `1000.0` | kW → W conversion inside ln() |

Probit formula: `Y = PA + PB × ln((KW_SCALE × kW)^DOE × t_exp)`  
Fatality probability: `P = Φ((Y − PROBIT_MEAN) / √2)`

### B3. Exposure time

| Parameter | v2 | MacroQRAV6 source |
|-----------|----|-------------------|
| `EXP_TIME` (seconds) | Hardcoded `20.0` in `main()` | `General` D21 |

```
_GEN_TEXP_RC = (21, 4)   # D21 — thermal exposure time [s]
```

### B4. Normal-CDF approximation constants (Abramowitz & Stegun — mathematical, not workbook data)

| Constant | Value |
|----------|-------|
| `_NC_A1` | `0.254829592` |
| `_NC_A2` | `-0.284496736` |
| `_NC_A3` | `1.421413741` |
| `_NC_A4` | `-1.453152027` |
| `_NC_A5` | `1.061405429` |
| `_NC_P` | `0.3275911` |

Maximum approximation error: < 1.5 × 10⁻⁷.

---

## C. Flash Fire (FF)

| Parameter | v2 | MacroQRAV6 source |
|-----------|----|-------------------|
| Impact value — outside LFLF zone | Implicit `0` | `General` J18 (first value) |
| Impact value — transition (LFLF to LFL) | Implicit `1` | `General` K18 (second value) |
| Impact value — inside LFL zone | Implicit `2` | `General` L18 (third value) |
| LFL radius per scenario | Read from `ImpactFFMatrix` col E | `Flash Fire Results` col E |
| LFLF radius per scenario | Read from `ImpactFFMatrix` col F | `Flash Fire Results` col F |

```
_GEN_FF_RC = (18, 10)   # J18 — FF zone values (3 values: outside / transition / inside)
```

Risk rule: cells where `impact ≥ FF_INSIDE_LFL` contribute `frequency × 1` to risk.

---

## D. Explosion / CVE (and BLEVE)

### D1. Overpressure thresholds (bar) — **NOT hardcoded, from PHAST**

| Parameter | v2 | MacroQRAV6 source |
|-----------|----|-------------------|
| `impact_v` (5 bar values, row 1 cols R–V of ImpactExpMatrix) | Read dynamically from sheet header | `General` row 19, starting col J (up to 5 values) |
| Distance per threshold per scenario (cols R–V of ImpactExpMatrix) | Read from `ImpactExpMatrix` cols R–V per row | `CVE Results` / `BLEVE Results` cols J–N per row |

```
_GEN_EXP_RC = (19, 10)   # J19 — first overpressure threshold [bar]
```

### D2. Risk step-function limits (from kernel AA cells, fixed per workbook run)

| Parameter | Kernel cell | Current value | Meaning |
|-----------|-------------|---------------|---------|
| `risk_lo` | `ImpactExpMatrix AA17` | `0.1` | Lower overpressure limit for risk contribution |
| `risk_hi` | `ImpactExpMatrix AA18` | `0.3` | Upper overpressure limit |
| `risk_factor_mid` | `ImpactExpMatrix AA19` | `1` | Multiplier for `risk_lo ≤ impact < risk_hi` |
| `risk_factor_high` | `ImpactExpMatrix AA20` | `None` (→ IFERROR → 0) | Multiplier for `impact ≥ risk_hi` |

In v6 these are encoded in `EXP_THRESHOLDS` indices — confirmed cells AA15–AA18.

### D3. Ignition coordinates — **CVE only, NOT from Core**

| Parameter | v2 source | MacroQRAV6 source |
|-----------|-----------|-------------------|
| `Xign` (ignition X, m) | `ImpactExpMatrix` col K per row | `CVE Results` col V (header `X`) |
| `Yign` (ignition Y, m) | `ImpactExpMatrix` col L per row | `CVE Results` col W (header `Y`) |

> **Critical distinction:** CVE ignition coordinates come from the effect results sheet, not from Core. Core X/Y are the release point; CVE X/Y are where ignition occurs.
> BLEVE uses Core X/Y (no separate ignition point column).

### D4. Effective formula behaviour (step function)

The Excel kernel calls `MATCH(d, distV, -1)` on an **ascending** `distV` array. `MATCH(-1)` expects descending order; with ascending input it returns `#N/A` for all `d ≥ min(distV)`, caught by the outer `IFERROR → 0`. Result:

```
impact = 1.0   if  d < min(distV)    (inside the first/smallest threshold radius)
         0.0   otherwise
```

---

## E. Jet Fire (JF)

### E1. Direction / angle parameters

| Parameter | Symbol | v2 | MacroQRAV6 source |
|-----------|--------|----|-------------------|
| Number of directions | `JF_N_DIRS` / `JF_DIRECTIONS` | Hardcoded `8` | `General` D24 |
| Angle offset (degrees) | `JF_ANGLE_OFFSET` | Hardcoded `0.0` | `General` D23 |
| Angle step (degrees, 0 = equal spacing) | `JF_ANGLE_STEP` | Hardcoded `0.0` | `General` D25 |
| Direction index count | `JF_INDEX` | Read from kernel `AK9` | `General` D22 |

```
_GEN_JF_INDEX_RC  = (22, 4)   # D22 — JF direction index count
_GEN_JF_OFFSET_RC = (23, 4)   # D23 — JF angle offset [deg]
_GEN_JF_DIRS_RC   = (24, 4)   # D24 — JF number of directions
_GEN_JF_STEP_RC   = (25, 4)   # D25 — JF angle step [deg] (0 = equal spacing)
```

### E2. kW/m² thresholds

Same thresholds as thermal (Section B1) — `JF_THRESHOLDS` in v6 defaults to `THERM_THRESHOLDS`.
Matched dynamically to `Jet Fire Results` row-1 headers.

### E3. Ellipse parameters per scenario (cell units, not metres)

| Column | v2 (ImpactJFMatrix) | MacroQRAV6 (Jet Fire Results) | Meaning |
|--------|---------------------|-------------------------------|---------|
| F–O (cols 6–15) | `dist_v` (10 values) | Cols F–O | Forward tip distance |
| P–Y (cols 16–25) | `half_v` (10 values) | Cols P–Y | Lateral half-width |
| Z–AI (cols 26–35) | `cent_v` (10 values) | Cols Z–AI | Centre distance from source |
| E (col 5) | `flame_len` | Col E | Flame length (reference only) |

Impact formula: for each direction angle θ and each threshold level, test whether each grid cell lies inside the ellipse defined by (dist, halfW, center). Impact = maximum kW/m² over all directions and thresholds where cell is inside.

---

## F. Toxic (TOXIC)

| Parameter | v2 | MacroQRAV6 source |
|-----------|----|-------------------|
| Minimum probability filter | `0.01` (hardcoded in `_parse_tox_blob`) | `General` J20 (first value of row 20) |
| Distance–probability profile | `ImpactToxMatrix` col G (CSV blob) | `Outdoor Toxic Results` col G |

```
_GEN_TOX_RC = (20, 10)   # J20 — toxic minimum probability filter
```

Blob format: 5-column CSV per row (distance, conc, log-value, probability, secondary).
Engine filters: `distance ≥ 0` AND `probability ≥ TOX_MIN_PROB`, then interpolates linearly.

---

## G. Core Sheet — Source Coordinates & Frequencies

Both engines read from the `Core` sheet (v2 uses `CoreExample`).

| Column (0-based) | Excel col | Field | Used by events |
|-----------------|-----------|-------|----------------|
| 0 | A | `key_scenario` | All |
| 3 | D | Source X (m) | All except CVE (uses Xign) |
| 4 | E | Source Y (m) | All except CVE (uses Yign) |
| 6 | G | Size label (S/M/L/XL/INST) | All (accumulation filter) |
| 9 | J | `P_TOXIC` — frequency | TOXIC |
| 10 | K | `P_JF` — frequency | JF |
| 11 | L | `P_LPF` — frequency | LPF (thermal) |
| 12 | M | `P_EPF` — frequency | EPF (thermal) |
| 13 | N | `P_FB` — frequency | FB (thermal) |
| 14 | O | `P_CVE` — frequency | CVE (explosion) |
| 15 | P | `P_BLV` — frequency | BLV (explosion) |
| 16 | Q | `P_FF` — frequency | FF (flash fire) |
| 17 | R | `P_LEXP` — frequency | Late explosion (deferred) |

Core read range: `A4:R89` in `MacroQRAV6`. Sentinel: `None / 0 / -1` in col A = stop.

---

## H. Effect Sheet → MacroQRAV6 Sheet Mapping

| Event | v2 kernel sheet | MacroQRAV6 sheet | Formula type |
|-------|----------------|------------------|--------------|
| TOXIC | `ImpactToxMatrix` | `Outdoor Toxic Results` | `toxic` |
| JF | `ImpactJFMatrix` | `Jet Fire Results` | `jf_ellipse` |
| LPF | `ImpactThermMatrix` | `Late Pool Fire Results` | `thermal` |
| EPF | `ImpactThermMatrix` | `Early Pool Fire Results` | `thermal` |
| FB | `ImpactThermMatrix` | `Fireball Results` | `thermal` |
| CVE | `ImpactExpMatrix` | `CVE Results` | `explosion` (is_cve=True) |
| BLV | `ImpactExpMatrix` | `BLEVE Results` | `explosion` |
| FF | `ImpactFFMatrix` | `Flash Fire Results` | `ff` |

Impact IDs in `MacroQRAV6 PageControl` (col L):

| Impact ID | Event |
|-----------|-------|
| 16 | TOXIC |
| 17 | JF |
| 18 | LPF |
| 19 | EPF |
| 20 | FB |
| 21 | CVE |
| 22 | BLV |
| 23 | FF |

---

## I. PageControl Layout (v6 only)

| Constant | Value | Description |
|----------|-------|-------------|
| `_PC_READ_RANGE` | `'J2:R500'` | Read range in PageControl sheet |
| `_PC_EVENT_COL` | `0` | Offset within range → col J (event description) |
| `_PC_ID_COL` | `2` | Offset → col L (Impact ID integer) |
| `_PC_SIZE_COLS['Total']` | `3` | Offset → col M |
| `_PC_SIZE_COLS['S']` | `4` | Offset → col N |
| `_PC_SIZE_COLS['M']` | `5` | Offset → col O |
| `_PC_SIZE_COLS['L']` | `6` | Offset → col P |
| `_PC_SIZE_COLS['XL']` | `7` | Offset → col Q |
| `_PC_SIZE_COLS['INST']` | `8` | Offset → col R |

---

## J. File Paths & Workbook Names

| Item | v2 engine | v6 engine |
|------|-----------|-----------|
| Input workbook | `KernelV0_v2_copy.xlsx` (hardcoded) | `MacroQRAV6 (version 1).xlsm` (from `.env` `WORKSPACE`) |
| Core sheet | `CoreExample` (hardcoded) | `Core` (hardcoded sheet name) |
| Output directory | `output_v2/` | `output/matrices/` (relative to `WORKSPACE`) |
| Export workbook | N/A | `MacroQRAV6_export_result.xlsx` |
| Result sheets | N/A | `Impact Matrix Result`, `Risk Matrix Result` |
| Size labels | N/A | `['Total', 'S', 'M', 'L', 'XL', 'INST']` |

---

## K. Complete Parameter Table

Every parameter used by the engines, with current value, v2 source, and v6 dynamic status.

**Dynamic status key:**
- `Dynamic — General` = read at runtime from `MacroQRAV6` General sheet
- `Dynamic — sheet` = read per-scenario from an effect results sheet
- `Dynamic — Core` = read per-scenario from the Core sheet
- `Hardcoded (formula)` = mathematical constant, never changes
- `Hardcoded (structural)` = structural code constant (sheet names, column offsets, etc.)

---

### K.1 Grid / Geometry

| Parameter | Value | v2 Reference | v6 Dynamic Status |
|-----------|-------|--------------|-------------------|
| `QX` | `315` | Hardcoded | Dynamic — `General` (row 5, col 20) `_GEN_QX_RC` |
| `QY` | `317` | Hardcoded | Dynamic — `General` (row 6, col 20) `_GEN_QY_RC` |
| `SX` | `1.0698412698412698` m/cell | Hardcoded | Dynamic — `General` (row 5, col 19) `_GEN_SX_RC` |
| `SY` | `1.069425` m/cell | Hardcoded | Dynamic — `General` (row 6, col 19) `_GEN_SY_RC` |

---

### K.2 Thermal Radiation (LPF / EPF / FB)

| Parameter | Value | v2 Reference | v6 Dynamic Status |
|-----------|-------|--------------|-------------------|
| `KW_THRESHOLDS` | `[1.6, 5.0, 7.3, 9.5, 12.5, 16.0, 20.9, 25.0, 30.0, 35.0]` kW/m² | Hardcoded array | Dynamic — `General` row 17, starting col J (`_GEN_THERM_RC`) |
| `EXP_TIME` / `THERM_T_EXP` | `20.0` s | Hardcoded in `main()` | Dynamic — `General` D21 (`_GEN_TEXP_RC`) |
| `_THERM_PA` (probit intercept) | `-36.38` | Hardcoded (formula) | Hardcoded (formula) |
| `_THERM_PB` (probit slope) | `2.56` | Hardcoded (formula) | Hardcoded (formula) |
| `_THERM_DOE` (dose exponent) | `4/3 ≈ 1.3333` | Hardcoded (formula) | Hardcoded (formula) |
| `_THERM_PROBIT_MEAN` | `5.0` | Hardcoded (formula) | Hardcoded (formula) |
| `_THERM_KW_SCALE` (kW → W) | `1000.0` | Hardcoded (formula) | Hardcoded (formula) |
| Threshold distances (10 cols) | per scenario | Read from `ImpactThermMatrix` cols F–O | Dynamic — sheet `Late/Early Pool Fire Results`, `Fireball Results` |

---

### K.3 Normal CDF (Abramowitz & Stegun approximation)

| Parameter | Value | v2 Reference | v6 Dynamic Status |
|-----------|-------|--------------|-------------------|
| `_NC_A1` | `0.254829592` | Hardcoded (formula) | Hardcoded (formula) |
| `_NC_A2` | `-0.284496736` | Hardcoded (formula) | Hardcoded (formula) |
| `_NC_A3` | `1.421413741` | Hardcoded (formula) | Hardcoded (formula) |
| `_NC_A4` | `-1.453152027` | Hardcoded (formula) | Hardcoded (formula) |
| `_NC_A5` | `1.061405429` | Hardcoded (formula) | Hardcoded (formula) |
| `_NC_P` | `0.3275911` | Hardcoded (formula) | Hardcoded (formula) |

---

### K.4 Flash Fire (FF)

| Parameter | Value | v2 Reference | v6 Dynamic Status |
|-----------|-------|--------------|-------------------|
| `FF_OUTSIDE` | `0` | Hardcoded (implicit in `compute_ff`) | Dynamic — `General` J18 (`_GEN_FF_RC`) |
| `FF_TRANSITION` | `1` | Hardcoded (implicit) | Dynamic — `General` K18 |
| `FF_INSIDE_LFL` | `2` | Hardcoded (implicit) | Dynamic — `General` L18 |
| LFL radius per scenario | per scenario | Read from `ImpactFFMatrix` col E | Dynamic — sheet `Flash Fire Results` col E |
| LFLF radius per scenario | per scenario | Read from `ImpactFFMatrix` col F | Dynamic — sheet `Flash Fire Results` col F |

---

### K.5 Explosion — CVE (and BLEVE)

| Parameter | Value | v2 Reference | v6 Dynamic Status |
|-----------|-------|--------------|-------------------|
| `impact_v` (5 bar thresholds) | `[0.83, 0.34, 0.17, 0.1, 0.03]` bar (current) | Read from `ImpactExpMatrix` row 1, cols R–V (dynamic — from PHAST) | Dynamic — `General` row 19, cols J–N (`_GEN_EXP_RC`) |
| `distV` (5 distances) | per scenario | Read from `ImpactExpMatrix` cols R–V per row | Dynamic — sheet `CVE Results` / `BLEVE Results` cols J–N per row |
| `risk_lo` | `0.1` bar | Read from `ImpactExpMatrix` AA17 | Encoded as `EXP_THRESHOLDS[0]` |
| `risk_hi` | `0.3` bar | Read from `ImpactExpMatrix` AA18 | Encoded as `EXP_THRESHOLDS[1]` |
| `risk_factor_mid` | `1` | Read from `ImpactExpMatrix` AA19 | Encoded as `EXP_THRESHOLDS[2]` |
| `risk_factor_high` | `None` (→ IFERROR → 0) | Read from `ImpactExpMatrix` AA20 | Encoded as `EXP_THRESHOLDS[3]` |
| `Xign` (ignition X) | per scenario | Read from `ImpactExpMatrix` col K | Dynamic — `CVE Results` col V (header `X`) |
| `Yign` (ignition Y) | per scenario | Read from `ImpactExpMatrix` col L | Dynamic — `CVE Results` col W (header `Y`) |

---

### K.6 Jet Fire (JF)

| Parameter | Value | v2 Reference | v6 Dynamic Status |
|-----------|-------|--------------|-------------------|
| `JF_DIRECTIONS` / `JF_N_DIRS` | `8` | Hardcoded | Dynamic — `General` D24 (`_GEN_JF_DIRS_RC`) |
| `JF_ANGLE_OFFSET` | `0.0°` | Hardcoded | Dynamic — `General` D23 (`_GEN_JF_OFFSET_RC`) |
| `JF_ANGLE_STEP` | `0.0°` (uniform 360°/N) | Hardcoded | Dynamic — `General` D25 (`_GEN_JF_STEP_RC`) |
| `JF_INDEX` | `8` | Read from kernel `ImpactJFMatrix AK9` | Dynamic — `General` D22 (`_GEN_JF_INDEX_RC`) |
| `JF_THRESHOLDS` | same as `KW_THRESHOLDS` | Hardcoded (same array) | `_JF_THRESHOLDS_DEFAULT` fallback; overridden by `General` row 17 at runtime |
| `dist_v` (10 tip distances) | per scenario | Read from `ImpactJFMatrix` cols F–O | Dynamic — sheet `Jet Fire Results` cols F–O |
| `half_v` (10 half-widths) | per scenario | Read from `ImpactJFMatrix` cols P–Y | Dynamic — sheet `Jet Fire Results` cols P–Y |
| `cent_v` (10 centre offsets) | per scenario | Read from `ImpactJFMatrix` cols Z–AI | Dynamic — sheet `Jet Fire Results` cols Z–AI |

---

### K.7 Toxic (TOXIC)

| Parameter | Value | v2 Reference | v6 Dynamic Status |
|-----------|-------|--------------|-------------------|
| `TOXIC_UNIT_MODE` | `'meters'` | Hardcoded | Confirmed (blob col 1 always in metres; not exposed in General) |
| `TOX_MIN_PROB` | `0.01` | Hardcoded in `_parse_tox_blob` | Dynamic — `General` J20 (`_GEN_TOX_RC`) |
| Distance–prob profile | per scenario | Read from `ImpactToxMatrix` col G (CSV blob) | Dynamic — sheet `Outdoor Toxic Results` col G |

---

### K.8 Core Sheet — Frequencies & Source Coordinates

| Parameter | Value | v2 Reference | v6 Dynamic Status |
|-----------|-------|--------------|-------------------|
| `_CORE_COL_KEY` | `0` (col A) | Hardcoded | Hardcoded `_CORE_COL_KEY = 0` |
| `_CORE_COL_X` | `3` (col D) | Hardcoded | Hardcoded `_CORE_COL_X = 3` |
| `_CORE_COL_Y` | `4` (col E) | Hardcoded | Hardcoded `_CORE_COL_Y = 4` |
| `_CORE_COL_SIZE` | `6` (col G) | Hardcoded | Hardcoded `_CORE_COL_SIZE = 6` |
| `_CORE_COL_TOXIC` (`P_TOXIC`) | `9` (col J) | Hardcoded | Hardcoded `_CORE_COL_TOXIC = 9` |
| `_CORE_COL_JF` (`P_JF`) | `10` (col K) | Hardcoded | Hardcoded `_CORE_COL_JF = 10` |
| `_CORE_COL_LPF` (`P_LPF`) | `11` (col L) | Hardcoded | Hardcoded `_CORE_COL_LPF = 11` |
| `_CORE_COL_EPF` (`P_EPF`) | `12` (col M) | Hardcoded | Hardcoded `_CORE_COL_EPF = 12` |
| `_CORE_COL_FB` (`P_FB`) | `13` (col N) | Hardcoded | Hardcoded `_CORE_COL_FB = 13` |
| `_CORE_COL_CVE` (`P_CVE`) | `14` (col O) | Hardcoded | Hardcoded `_CORE_COL_CVE = 14` |
| `_CORE_COL_BLV` (`P_BLV`) | `15` (col P) | Hardcoded | Hardcoded `_CORE_COL_BLV = 15` |
| `_CORE_COL_FF` (`P_FF`) | `16` (col Q) | Hardcoded | Hardcoded `_CORE_COL_FF = 16` |
| Source X, Y, freq, size | per scenario | Dynamic — read from `CoreExample` rows | Dynamic — Core `A4:R89` |

---

### K.9 Impact Event Configuration (IMPACT_CONFIG)

| Impact ID | Event | Formula | prob_idx | v2 Reference | v6 Dynamic Status |
|-----------|-------|---------|----------|--------------|-------------------|
| 16 | TOXIC | `toxic` | 9 | Hardcoded `EVENTS` list | Hardcoded `IMPACT_CONFIG[16]` |
| 17 | JF | `jf_ellipse` | 10 | Hardcoded | Hardcoded `IMPACT_CONFIG[17]` |
| 18 | LPF | `thermal` | 11 | Hardcoded | Hardcoded `IMPACT_CONFIG[18]` |
| 19 | EPF | `thermal` | 12 | Hardcoded | Hardcoded `IMPACT_CONFIG[19]` |
| 20 | FB | `thermal` | 13 | Hardcoded | Hardcoded `IMPACT_CONFIG[20]` |
| 21 | CVE | `explosion` (cve=True) | 14 | Hardcoded | Hardcoded `IMPACT_CONFIG[21]` |
| 22 | BLV | `explosion` | 15 | Hardcoded | Hardcoded `IMPACT_CONFIG[22]` |
| 23 | FF | `ff` | 16 | Hardcoded | Hardcoded `IMPACT_CONFIG[23]` |

---

### K.10 PageControl Layout (v6 only)

| Parameter | Value | v2 Reference | v6 Dynamic Status |
|-----------|-------|--------------|-------------------|
| `_PC_READ_RANGE` | `'J2:R500'` | N/A (no PageControl in v2) | Hardcoded (structural) |
| `_PC_EVENT_COL` | `0` | N/A | Hardcoded (structural) |
| `_PC_ID_COL` | `2` | N/A | Hardcoded (structural) |
| `_PC_SIZE_COLS['Total']` | `3` | N/A | Hardcoded (structural) |
| `_PC_SIZE_COLS['S']` | `4` | N/A | Hardcoded (structural) |
| `_PC_SIZE_COLS['M']` | `5` | N/A | Hardcoded (structural) |
| `_PC_SIZE_COLS['L']` | `6` | N/A | Hardcoded (structural) |
| `_PC_SIZE_COLS['XL']` | `7` | N/A | Hardcoded (structural) |
| `_PC_SIZE_COLS['INST']` | `8` | N/A | Hardcoded (structural) |

---

### K.11 File Paths & Sheet Names

| Parameter | Value | v2 Reference | v6 Dynamic Status |
|-----------|-------|--------------|-------------------|
| Input workbook | `KernelV0_v2_copy.xlsx` | Hardcoded `INPUT_FILE` | `MacroQRAV6 (version 1).xlsm` — from `.env` `WORKSPACE` |
| Core sheet name | `CoreExample` | Hardcoded `CORE_SHEET` | `Core` — hardcoded (structural) |
| Output directory | `output_v2/` | Hardcoded `OUTPUT_DIR` | `output/matrices/` relative to `WORKSPACE` |
| Export workbook | N/A | N/A | `MacroQRAV6_export_result.xlsx` — from `WORKSPACE` |
| Result sheet — impact | N/A | N/A | `Impact Matrix Result` — hardcoded `SHEET_IMPACT` |
| Result sheet — risk | N/A | N/A | `Risk Matrix Result` — hardcoded `SHEET_RISK` |
| Size labels | N/A | Implicit in v2 loop | `['Total', 'S', 'M', 'L', 'XL', 'INST']` — hardcoded `SIZES` |

---

## L. Kernel V0-4 Comparison (`KernelV0 (version 1) 4.xlsx`, from Downloads)

Compared against `KernelV0_v2_copy.xlsx` (the file `qra_engine_v2.py` currently reads) by
diffing sheet inventory, the `AB1`/`AL1` array-formula text (the single formula each effect
sheet fills down/across), row-1 column headers, and the per-sheet `AA`/`AK` test-harness
parameter block. Full raw dumps are not kept — this section is the synthesized result.

### L.1 Sheet inventory

| Sheet | New kernel (V0-4) | Current kernel (v2_copy) |
|-------|--------------------|---------------------------|
| ANCHOR, SumMatrix, WindMatrix, MaxMatrix, RiskMatrix0, ImpactMatrix0 | present | present |
| ImpactThermMatrix, ImpactToxMatrix, ImpactFFMatrix, ImpactExpMatrix, ImpactJFMatrix | present | present |
| `Core` / `CoreExample` | **absent** | `CoreExample` present |
| `Sheet1` | absent | present (unused scratch sheet) |

**The new kernel has no Core/CoreExample sheet at all** — no source X/Y or scenario
frequencies. It cannot be used as a drop-in `INPUT_FILE`/`CORE_FILE` replacement for
`qra_engine_v2.py` as-is; it only carries the effect-matrix formula/data sheets.

### L.2 `AB1` / `AL1` formula text — identical vs. changed

| Sheet | Formula vs. current kernel | Notes |
|-------|------------------------------|-------|
| `ImpactThermMatrix` | **Identical**, character-for-character | Same probit constants (`-36.38`, `2.56`, `4/3`, `1000`), same interpolation logic — matches `qra_engine_v2.py`'s `thermal` formula exactly |
| `ImpactFFMatrix` | **Identical** | Same step-function logic (`$AA$19`/`$AA$17`/`$AA$18` outside/inside/transition) — matches `ff` formula |
| `ImpactToxMatrix` | **Identical** | Same `MATCH`/interpolation against col `J`/`M` — matches `toxic` formula |
| `ImpactJFMatrix` | **Identical** | Same 8-direction ellipse formula (`aPower`/`bPower`/rotation) — matches `jf_ellipse` formula |
| `ImpactExpMatrix` | **Changed** — same interpolation/step-risk shape, but reads a different column range (see L.3) | `distV`/`impactV` now `$J:$N`/`$J$1:$N$1` instead of `$R:$V`/`$R$1:$V$1` |

Conclusion: the 4 formulas `qra_engine_v2.py` actually implements (thermal, ff, toxic,
jf_ellipse) are unchanged in the new kernel — safe to point the engine at new scenario data
for those. Explosion (CVE), which the engine currently defers, has a **real** layout change.

### L.3 `ImpactExpMatrix` column reassignment (CVE — currently deferred in v2)

| Field | Current kernel col | New kernel (V0-4) col |
|-------|---------------------|------------------------|
| Location Name / Area / S-L DIST | G / H / I | G / H / I (unchanged) |
| SIZE | J | **O** |
| Xign | **K** | **P** |
| Yign | **L** | **Q** |
| TimeExp / DistCM / DistExp / DistIgn / FlammMass | M–Q | R–V |
| Overpressure thresholds (5 values) | **R–V** = `[0.83, 0.34, 0.17, 0.1, 0.03]` | **J–N** = `[0.83, 0.34, 0.1, 0.17, 0.03]` (note: 3rd/4th also swapped) |

`read_exp()` in `qra_engine_v2.py` (lines ~431–449) hardcodes `K`/`L` for Xign/Yign and
`R,S,T,U,V` for thresholds/distances — **these must be changed to `P`/`Q` and `J,K,L,M,N`
respectively** before pointing that reader at kernel V0-4. Since CVE isn't wired into
`main()` yet, this is currently latent, not an active bug — flagging so it isn't missed
when CVE gets implemented.

### L.4 Row-1 column headers — Thermal / FF / Tox / JF

Identical column layout in both kernels (verified letter-by-letter): `B/C/D` = Path/Scenario/Weather,
`E` = pool diameter or flame length, `F–O` = the 10 kW/m² thresholds (thermal/JF),
`P–Y`/`Z–AI` = JF half-width/center columns, `E/F` = LFL/LFLF distance (FF), toxic blob in col `G`.
No changes here — `read_therm`, `read_ff`, `read_tox`, `read_jf` in `qra_engine_v2.py` all
match the new kernel without modification.

### L.5 Per-sheet `AA`/`AK` test-harness block

Each effect sheet has a manual single-scenario test area (cols `Z:AA` or `AJ:AK`, rows 1–21:
impact id, impact/risk flag, selected scenario, X/Y, probability, SX/SY, QX/QY, risk limits 1–5).
These are **user-editable scratch inputs** for spot-checking one cell, not workbook constants —
values differ between the two files simply because a different scenario was left selected:

- New kernel has `QX=1, QY=1` in every sheet vs. `QX=315, QY=317` in the current kernel — looks
  like the grid was shrunk to 1×1 for fast manual testing before being scaled back up; **not**
  a production grid-size change (confirmed `SX`/`SY` = `1.0698.../1.069425` match in most sheets).
- One anomaly: `ImpactFFMatrix` in the new kernel has `SX=0.338353..., SY=0.338142` — inconsistent
  with every other sheet in both files. Likely stale test input on that sheet, worth a sanity check
  next time that workbook is opened in Excel, but not a formula/reading-pattern issue.
- `ImpactExpMatrix` risk limits 3/4 (`AA19`/`AA20`) are `0`/`1` in the new kernel vs `1`/`None` in
  the current one — consistent with the same scenario-selection noise, not a spec change (the
  formula shape referencing `$AA$17–$AA$20` as risk_lo/risk_hi/mid-factor/high-factor is unchanged).

### L.6 Scenario counts (informational — not a structural issue)

| Sheet | New kernel rows | Current kernel rows |
|-------|------------------|-----------------------|
| ImpactThermMatrix | 53 | 317 |
| ImpactFFMatrix | 133 | 317 |
| ImpactExpMatrix | 24 | 317 |
| ImpactJFMatrix | 113 | 317 |
| ImpactToxMatrix | up to 10000 (blob-filled) | up to 10000 (blob-filled) |

Both readers stop at the first blank `col B`, so a smaller dataset doesn't require code
changes — this just reflects a smaller/different scenario set in the new kernel file.

### L.7 Bottom line

- **Thermal, Flash Fire, Toxic, Jet Fire: formulas and reading pattern are unchanged.** Safe to
  swap `INPUT_FILE`/`CORE_FILE` (once a Core sheet is supplied) to the new kernel for these four.
- **Explosion/CVE: formula shape unchanged, but column addresses moved.** `read_exp()` needs the
  column-letter update in L.3 before it can be used against kernel V0-4.
- **No Core/CoreExample sheet in the new kernel** — need a separate source for X/Y and
  frequencies (either add one to this file or keep reading `Core` from `MacroQRAV6`).

---

## M. Real Production Workbook Comparison (`MacroQRAV6 (version 1)-Win11-cWvAMQq4D.xlsm`, from Downloads)

Scope: only the **Results sheet columns** that `qra_engine_v2.py`'s `read_therm/read_ff/read_tox/read_exp/read_jf`
hardcode by letter — not `QX/QY/SX/SX` or other `General`-sheet parameters (those are out of
scope here; `qra_v6_engine.py` already reads them dynamically). Sheet name equivalence used:

| Kernel sheet | Production sheet |
|---|---|
| `ImpactThermMatrix` | `Early Pool Fire Results` / `Late Pool Fire Results` / `Fireball Results` |
| `ImpactToxMatrix` | `Outdoor Toxic Results` |
| `ImpactFFMatrix` | `Flash Fire Results` |
| `ImpactExpMatrix` | `CVE Results` (BLEVE has its own `BLEVE Results` sheet, not covered by `read_exp()`) |
| `ImpactJFMatrix` | `Jet Fire Results` |

Verified by dumping row-1 headers + sample rows directly from the workbook (`openpyxl`,
read-only). All target sheets exist under the exact names above (plus `Core`/`PageControl`/`General`).

### M.1 Thermal — Early/Late Pool Fire Results, Fireball Results

`read_therm()` reads `B,C,D,E` + thresholds at `F:O`. **Matches exactly.** Production headers:
`B=Path, C=Scenario, D=Weather, E=Pool/Fireball diameter [m], F:O=[1.6,5,7.3,9.5,12.5,16,20.9,25,30,35]`
— identical column letters and values to both kernel files. No changes needed.

(Production also carries a second, fully-labelled copy of the same 10 columns at `AA:AN`
— e.g. `AE1='Distance downwind to intensity level 1 (1.6 kW/m2) [m]'` — a human-readable
mirror of `F:O`. Not needed by the reader; just noting it exists so it isn't mistaken for
a second scenario block.)

### M.2 Flash Fire — Flash Fire Results

`read_ff()` reads `B,C,D,E(LFL),F(LFLF)`. **Matches exactly** — identical headers and column
letters to both kernel files.

### M.3 Toxic — Outdoor Toxic Results

`read_tox()` reads `B,C,D,G(blob)`. **Matches** — `B=SCENARIO_CODE, C=SCENARIO, D=WEATHER,
G=Toxic Dispersion Table` (blob), same as both kernels. `E/F` (`START TABLE ROW`/`TABLE HEIGHT`)
present but unused by the reader, as in the kernel.

### M.4 Jet Fire — Jet Fire Results

`read_jf()` reads `B,C,D,E(flame_len),F:O(dist_v),P:Y(half_v),Z:AI(cent_v)`. **Matches exactly**
— same 30-column ellipse-parameter layout as both kernels (values sparse per row when the
flame doesn't reach the outer thresholds, same as the kernel — not a discrepancy).

Production carries **two additional representations further right that the reader doesn't
use and doesn't need to**: a blob-style block at `AN:AT` (`RANDKEY`, `Frustum tip width`,
and two CSV-blob columns for half-width/center across all 10 thresholds — same blob idiom
as the toxic sheet), and a flat single-direction table at `AV:BI` (`Path/Scenario/Weather/
Flame length` + 10 labelled `Distance downwind to intensity level N` columns, using the text
`"Not reached at height of interest"` in place of blanks). These look like pre-rendered
report views of the same ellipse data — informational only, no action needed.

### M.5 Explosion/CVE — CVE Results — **major mismatch, `read_exp()` cannot be pointed at this sheet as-is**

This is the one real problem found. Column positions do **partially** match the *new* kernel
(V0-4) layout from Section L — `O=SIZE, P=Xign, Q=Yign` (not the current kernel's `J/K/L`) —
but the core assumption behind `read_exp()`, that there are 5 fixed threshold columns with
row-1 headers equal to the bar values, **does not hold**:

- `J1:N1` (where a threshold row should be) = `1, #N/A, #N/A, #N/A, #N/A` — broken/error
  formulas in this workbook, not real bar values (`0.83/0.34/0.17/0.1/0.03`).
- The real overpressure-vs-distance data lives in a **separate blob-table block at columns
  `Z:AG`**: `Z=SCENARIO_CODE, AA=SCENARIO, AB=WEATHER, AC=START TABLE ROW, AD=TABLE HEIGHT,
  AE='Explosion Table (Parameters)', AF='Explosion Table (OV Levels)', AG='Explosion Table
  (OV Distance)'` — the same **blob/CSV-per-row idiom used by `Outdoor Toxic Results` col G**,
  not a fixed 5-column layout at all.
- There's also a third, unrelated block at `AJ:AV` (`Type/L0/OVP0/L1/OVP1/X/1/K/K/OVP/L/
  OVPmin/OVPmax`) that looks like intermediate multi-energy-method calc data — not consumed
  by any current reader.

**Bottom line for CVE:** `read_exp()`'s column-letter approach (`K/L` for Xign/Yign, `R:V` or
`J:N` for 5 fixed threshold columns) doesn't match production in either kernel-derived form.

**Update — fixed 2026-07-13:** per explicit instruction, `read_exp()` was adapted to the new
column positions (`P/Q` for Xign/Yign, `J:N` for the 5 distance columns) **without** parsing
the `Z:AG` blob table — the decision was that the per-scenario distance data needed for
`compute_exp()`'s step function is already sitting in `J:N` under the broken row-1 headers,
so the blob table isn't required for the current formula. Row-1 `impact_v` (the bar threshold
labels) falls back to a hardcoded `[0.83, 0.34, 0.1, 0.17, 0.03]` when the headers are
non-numeric (`#N/A`) — confirmed this matches kernel V0-4's intact headers in the same column
order. `compute_exp()` itself only uses `min(distances)` (see the `MATCH(-1)`-on-ascending-array
quirk noted above), so `impact_v` only affects the informational CSV header, not the computed
matrix — the fallback is safe.

`INPUT_FILE` was also switched to `KernelV0 (version 1) 4.xlsx` (copied into the project root)
since it's the file with the new column layout; `CORE_FILE`/`CORE_SHEET` stayed on
`KernelV0_v2_copy.xlsx` / `CoreExample` because kernel V0-4 has no Core sheet of its own
(Section L.1) — scenario keys were verified to match across the two files (near-100% hit rate
per event, see `qra_engine_v2.py` run output). **This means `qra_engine_v2.py` no longer
reads CVE correctly from `KernelV0_v2_copy.xlsx` alone** — pointing `INPUT_FILE` back at the
old kernel would silently misread `ImpactExpMatrix` (verified: it would read `DistIgn`/
`FlammMass` as Xign/Yign and `SIZE/Xign/Yign/TimeExp/DistCM` as the 5 distances).

### M.6 BLEVE — BLEVE Results (no reader exists yet in `qra_engine_v2.py`)

Not read by any current v2 function, but worth recording for when it's implemented: production
`BLEVE Results` **does** use a simple fixed-column layout — `B/C/D` + 5 overpressure thresholds
at `J:N` = `[14.5038, 7.25189, 5.07632, 1.45038, 0.580151]` — **note: units are psi, not bar**
(differs from CVE's bar units). No Xign/Yign columns (BLEVE uses Core X/Y directly, per Section D3).
Because BLEVE's layout is simple and CVE's is a blob table, they **can't share one reader**
the way `qra_v6_engine.py`'s `read_explosion_results(is_cve=...)` currently assumes — that
function's dynamic header-match would fail on CVE (headers are `#N/A`) and its positional
fallback (`J:N`) would silently read BLEVE-shaped data from the wrong sheet.

### M.7 Summary — what's missing / not coinciding

| Event | Reader column letters vs. production | Verdict |
|---|---|---|
| Thermal (LPF/EPF/FB) | `B,C,D,E,F:O` | ✅ Matches |
| Flash Fire | `B,C,D,E,F` | ✅ Matches |
| Toxic | `B,C,D,G` | ✅ Matches |
| Jet Fire | `B,C,D,E,F:O,P:Y,Z:AI` | ✅ Matches |
| CVE (explosion) | `K/L` (Xign/Yign), `R:V`/`J:N` (thresholds) | ✅ Fixed — `read_exp()` now reads `P/Q` (Xign/Yign) and `J:N` (distances), with a hardcoded fallback for the broken `J1:N1` threshold labels. `INPUT_FILE` switched to kernel V0-4. |
| BLEVE | no reader | ⚠️ Not implemented; production layout (`B,C,D,J:N`, psi units) is simple and ready to read whenever it's added — but can't reuse a CVE reader. |

---

## N. `General`/`Core` Parameter Address Audit (for `qra_engine_v2.1.py`)

Scope: the parameters `qra_engine_v2.1.py` still hardcodes or reads from the wrong place —
`QX/QY/SX/SY`, per-event thresholds, JF direction params, and the `Core` frequency columns.
Checked the "pre-established" addresses documented in Section A–G (sourced from
`qra_v6_engine.py`'s `_GEN_*_RC`/`_CORE_COL_*` constants) directly against `General` and
`Core` in `MacroQRAV6 (version 1)-Win11-cWvAMQq4D.xlsm`. Row labels (col B/H on `General`,
row 3 on `Core`) were read alongside the values to remove ambiguity — several assumed
addresses turned out to be wrong or conceptually different from what the label says.

### N.1 `General` sheet — confirmed layout (row label → cell → value)

| Row label (as found) | Cell(s) | Value | Matches assumed address? |
|---|---|---|---|
| (Map grid qty, X) | `S5` / `T5` | `SX=1.0698412698412698` / `QX=315` | ✅ matches `_GEN_SX_RC`/`_GEN_QX_RC` |
| (Map grid qty, Y) | `S6` / `T6` | `SY=1.0709779179810726` / `QY=317` | ⚠️ **`QY` matches, but `SY` does NOT** — hardcoded/kernel value is `1.069425`, production `General!S6` is `1.0709779179810726` (~0.14% different). Real, not a typo. |
| `Thermal Levels` | `J17:S17` | `[1.6,5,7.3,9.5,12.5,16,20.9,25,30,35]` | ✅ matches `_GEN_THERM_RC`, same as hardcoded `KW_THRESHOLDS` |
| `Flammable concentration levels` | `J18:L18` | `[0,1,2]` | ✅ matches `_GEN_FF_RC` (FF_OUTSIDE/TRANSITION/INSIDE) |
| `Overpressure levels` | `J19:N19` | `[0.04,0.1,0.35,0.5,1]` | ⚠️ matches `_GEN_EXP_RC` address, **but this is NOT the same thing as the per-scenario bar thresholds in `CVE Results` row 1** (those are `[0.83,0.34,0.17,0.1,0.03]` — see Section M.5). These `General` values are the **risk step-function limits** (`limit1=0.04, limit2=0.1, prob1=0.35, prob2=0.5`) per `qra_v6_engine.py`'s `_explosion_risk_prob()` — a second, separate parameter set that happens to share the row-19 address pattern. `compute_exp()` in v2.1 doesn't use this yet (its risk branch still hardcodes `risk_lo=0.1, risk_hi=0.3`, which match neither this nor the per-sheet `AA17/AA18` test cells from Section D2/L.5). |
| `Toxic Levels` | `J20:L20` | `[0.01, 0.5, 0.9]` | ✅ first value matches `_GEN_TOX_RC` / `TOX_MIN_PROB` usage; the other two are extra report-threshold levels, not currently consumed |
| `Discharge directions (Default)` | `D21` | `8` | ❌ **This is `JF_N_DIRS`, not thermal exposure time.** `_GEN_TEXP_RC=(21,4)` in `qra_v6_engine.py` is wrong for this file. |
| `Exposure duration (Thermal effects)` | `D22` | `20` | ❌ **This is `THERM_EXP_TIME`, not `JF_INDEX`.** `_GEN_JF_INDEX_RC=(22,4)` is wrong for this file — **`D21`/`D22` are swapped relative to what `qra_v6_engine.py` assumes.** |
| *(none found)* | `D23`, `D24`, `D25` | — empty — | ❌ `_GEN_JF_OFFSET_RC`/`_GEN_JF_DIRS_RC`/`_GEN_JF_STEP_RC` point at empty cells. `General` has no per-run angle-offset/angle-step/direction-count triplet at all beyond the single `D21` default. |

### N.2 Where JF angle offset/step actually live

Not in `General`. `Core` row 3 headers (confirmed) are:
`A=KEY SCENARIO, B=Scenario code, C=Weather, D=X, E=Y, F=Scenario Description, G=Disch. Range,
H=Discharge directions, I=Offset Angle, J=Step Angle, K=Exposure Time, L=Basic Frequency (year)-1,
M=P Weather (PM), N=P_TOXIC, O=P_JF, P=P_LPF, Q=P_EPF, R=P_FB, S=P_CVE, ...`

So **JF direction count, angle offset, angle step, and thermal exposure time are stored
PER-SCENARIO in `Core` columns H/I/J/K** — sample row 4: `H=8, I=0, J=0, K=20`. This lines up
exactly with the per-sheet test-harness cells found in Section L.5 (kernel V0-4:
`AK9='directions'=8, AK10='angle offset'=0, AK11='angle steps'=0, AA8/AK8='Exp Time'=20`).
`D21` in `General` (`8`) looks like a **default/fallback**, not the live per-scenario value.

### N.3 `Core` sheet — frequency columns shifted vs. `_CORE_COL_*`

Production `Core` inserts 4 extra columns (`H,I,J,K` = discharge directions/offset/step/exposure
time) between `Y` and the frequency block that `_CORE_COL_TOXIC=9` etc. don't account for:

| Field | `_CORE_COL_*` assumed (0-based) | Production actual |
|---|---|---|
| X | `3` (col D) | ✅ `D` |
| Y | `4` (col E) | ✅ `E` |
| Size | `6` (col G, `Disch. Range`) | ✅ `G` |
| P_TOXIC | `9` (col J) | ❌ **col N** (0-based 13) |
| P_JF | `10` (col K) | ❌ **col O** (0-based 14) |
| P_LPF | `11` (col L) | ❌ **col P** (0-based 15) |
| P_EPF | `12` (col M) | ❌ **col Q** (0-based 16) |
| P_FB | `13` (col N) | ❌ **col R** (0-based 17) |
| P_CVE | `14` (col O) | ❌ **col S** (0-based 18) |
| P_BLV | `15` (col P) | ❌ **col T** (0-based 19) — confirmed via row 2 label `P_BLV` |
| P_FF | `16` (col Q) | ❌ **col U** (0-based 20) — confirmed via row 2 label `P_FF` |
| P_LEXP | `17` (col R) | ❌ **col V** (0-based 21) — confirmed via row 2 label `P_LEXP` |

Also note: `Core` data starts at **row 4** (rows 1-2 are numeric/category headers, row 3 is the
real header row) — same `min_row=4` that `qra_engine_v2.py`'s `load_core()` already uses, so
that part is fine; only the column offsets need updating.

### N.4 Open questions — resolved by the user, implemented in `qra_engine_v2.1.py`

1. **`SY` mismatch** — user confirmed `General!S6` is authoritative; no more hardcoding —
   `QX/QY/SX/SY` are read dynamically every run via `load_general_params()`. Also clarified
   that rows 17-20 ("Thermal/Flammable/Overpressure/Toxic Levels") must be read as a
   **variable-length run from col J to S**, stopping at the first blank — not a fixed count
   per row. Implemented as `_read_levels(ws, row, start_col='J', max_cols=10)`.
2. **CVE risk limits** — confirmed: use `General!J19:N19` (`EXP_THRESHOLDS`). User flagged
   that CVE is "more complex" than a straight port, so `compute_exp()` now uses these
   thresholds with an explicit code comment explaining the mismatch between the binary
   0/1 `impact` value (MATCH(-1) quirk) and the continuous-bar-pressure assumption the
   4-value step function was designed for — left as a known-degenerate placeholder, not a
   full fix, per instruction.
3. **JF direction params** — confirmed: `Core!H/I/J` (directions/offset/step) **per
   scenario**, falling back to `General!D21` (directions default) only when the Core cell
   is empty — not to fixed `General!D22:D25` constants (those don't exist for offset/step
   in this workbook). Implemented via `_scenario_val()` + `JF_ANGLE_OFFSET_DEFAULT`/
   `JF_ANGLE_STEP_DEFAULT` module fallbacks (both `0.0`, since `General` has no cell for them).
4. **Thermal/JF exposure time** — confirmed: `Core!K` per scenario, falling back to
   `General!D22` only when empty. Implemented; `main()` no longer has a single fixed
   `EXP_TIME` — thermal and JF each resolve their own per-scenario value.
5. ~~`Core!T/U/V`~~ — resolved: confirmed `T=P_BLV, U=P_FF, V=P_LEXP` (same 4-column shift),
   via the row-2 category labels directly above the row-3 headers.

**Verification:** ran `qra_engine_v2.1.py` end-to-end against the real Macro workbook
(`MacroQRAV6 (version 1)-Win11-cWvAMQq4D.xlsm`, copied into the project) for `Core`/`General`,
paired with kernel V0-4 for scenario/threshold data. Confirmed via direct cell checks that
`load_core()` now returns the exact `N:V` frequency values and `H:K` per-scenario JF/exposure
values found by manual inspection, and that `_scenario_val()` correctly overrides with the
Core value when present (tested against `P7301AB/L7/REF_FL16/H/1mDia`, `exp_time=21` vs. the
`General!D22` default of `20`) and falls back to the default otherwise. Full run completes
with exit code 0, no errors. Scenario-key overlap between kernel V0-4 and this particular
Macro file's `Core` sheet is low (~35-50%, 0/6 for CVE) — expected, since they're different
snapshots/exports, not the same production run; in real use both would come from the same
Macro run and match near-100%, same as `CoreExample` did with kernel V0-4 (Section L.1).

---

## O. Event/Sheet Mapping — `qra_engine_v2.1.py` now reads the Macro directly

### O.1 What changed

`INPUT_FILE` now equals `CORE_FILE`/`GENERAL_FILE` — kernel V0-4 is no longer used by
`qra_engine_v2.1.py` at all; everything (`General`, `Core`, and the Results sheets) comes
from `MacroQRAV6 (version 1)-Win11-cWvAMQq4D.xlsm`. New `EVENT_CONFIG` list documents the
kernel-sheet → production-sheet → frequency-column mapping, matching `qra_v6_engine.py`'s
`IMPACT_CONFIG` (Impact IDs 16-23) and the user's stated equivalence:

| Kernel sheet | Production sheet(s) | Core freq. column |
|---|---|---|
| `ImpactThermMatrix` | `Early Pool Fire Results` | `P_EPF` |
| | `Late Pool Fire Results` | `P_LPF` |
| | `Fireball Results` | `P_FB` |
| `ImpactFFMatrix` | `Flash Fire Results` | `P_FF` |
| `ImpactToxMatrix` | `Outdoor Toxic Results` | `P_TOXIC` |
| `ImpactExpMatrix` | `CVE Results` | `P_CVE` |
| `ImpactJFMatrix` | `Jet Fire Results` | `P_JF` |

Thermal is a **1-to-3 mapping** — the kernel's single `ImpactThermMatrix` collapsed three
distinct production events into one test sheet. `main()`'s THERMAL block is now a loop over
the 3 `EVENT_CONFIG` entries with `formula == 'thermal'`, writing to separate
`thermal_epf/`, `thermal_lpf/`, `thermal_fb/` (+ `_risk`) output folders — a shared scenario
key can legitimately appear in more than one of the three sheets (same leak, different fire
model), so they can't share one output folder without collisions.

Verified: full run against the Macro workbook completes with exit code 0. Per-event
written/skipped counts (this specific Core snapshot, frozen at Impact ID 23 — see O.2):
EPF 21/52, LPF 21/68, FB 0/30, FF 60/132, TOXIC 47/86, CVE 5 unique keys, JF 60/118.

### O.2 Still missing for "full" Macro reading

1. **BLEVE (`BLEVE Results`) has no reader.** Own column layout (`B,C,D,J:N`, **psi** not
   bar — Section M.6), and unlike CVE it uses **Core X/Y directly** (no Xign/Yign columns),
   so it can't reuse `compute_exp()` as-is. Needs its own `read_blv()` + a formula variant
   (structurally closer to `compute_thermal()`'s fixed-column interpolation than to CVE's
   ignition-point step function).
2. **`Core` is read once, statically — but it's actually a live, formula-driven table keyed
   by `Core!C2` (Impact ID).** This is the important one — see O.3.

### O.3 How `qra_v6_engine.py` iterates Core / PageControl (analysis, not yet implemented in v2.1)

Read `qra_v6_engine.py` end to end for this. Summary of the mechanism (lines ~1060-1200):

**`IMPACT_CONFIG`** (lines 96-105) is the authoritative event table — the same one `O.1`'s
mapping is built from, plus BLEVE and a `prob_idx` per event (0-based column offset into a
Core row, used to pick which frequency field `compute_event()` reads for that event):

```
16: Outdoor Toxic Results   TOXIC   formula=toxic
17: Jet Fire Results        JF      formula=jf_ellipse
18: Late Pool Fire Results  LPF     formula=thermal
19: Early Pool Fire Results EPF     formula=thermal
20: Fireball Results        FB      formula=thermal
21: CVE Results             CVE     formula=explosion (cve=True — uses Xign/Yign)
22: BLEVE Results           BLV     formula=explosion (uses Core X/Y, not Xign/Yign)
23: Flash Fire Results      FF      formula=ff
```
(`prob_idx` values there are 9-16, matching v6's **old** Core column assumption — our
`_CORE_FREQ_IDX` in v2.1 already corrects this to 13-21 for the real `N:V` layout; the
Impact-ID/sheet/formula mapping itself is unaffected by that column shift.)

**Per-event loop** (the actual Core iteration, requires `xlwings` / a live Excel session —
this is *why* v6 uses `xw` instead of `openpyxl` for `Core`):

```python
ws_core.range('C2').value = impact_id   # write the Impact ID into Core!C2
time.sleep(1.5); app.api.Calculate(); time.sleep(1.0)   # force full recalculation
size_filter = ws_core.range('G2').value   # v6's assumption — CORRECTED below, see O.4
scenarios = read_core_scenarios(ws_core)  # re-read Core!A4:R89 — CORRECTED below, see O.4
```

Then it reads the matching Results sheet (per `IMPACT_CONFIG[impact_id]['sheet']`), computes
impact/risk matrices, and stores them keyed by `impact_id` in `all_results`.

**`PageControl`** (`_PC_READ_RANGE = 'J2:R500'`) serves two purposes, both read *without*
needing a live Core recalc (it's a static control table, not formula-driven per Impact ID):
- `read_page_control()` → `{impact_id: True}` for every row where column `L` (offset 2 from
  `J`) parses as an integer **and** at least one of columns `M:R` (`Total/S/M/L/XL/INST`)
  contains an Excel-range-looking string (`_looks_like_range()`) — this is the "is this
  event turned on" flag.
- `read_directions()` → `{impact_id: {size: range_str}}` — the literal destination range
  (e.g. `$LD$318:$XF$634`) in the `Impact Matrix Result`/`Risk Matrix Result` sheets where
  that impact's QY×QX block gets bulk-written, per size filter. Only relevant once v2.1
  needs to reproduce the final combined-matrix write-back — not needed for the current
  per-scenario CSV output.

### O.4 Corrections from the user + verified against the real workbook

`qra_v6_engine.py`'s Core assumptions (`G2` filter cell, `A4:R89` read range) are **stale**
for the current production Core layout — corrected and verified directly against
`MacroQRAV6 (version 1)-Win11-cWvAMQq4D.xlsm`:

- **Size filter cell is `H2`, not `G2`.** `G2` is just the text label `'Size:'`; `H2` holds
  the live filter *value*. Confirmed: in this file, `H2 = 'M'` right now (not blank) — this
  Core snapshot is filtered to Medium-size scenarios only, on top of being frozen at
  `C2 = 23` (Flash Fire). Two independent live filters were stacked in the saved snapshot,
  which is *additional* evidence (beyond the per-event match-rate pattern in O.3) that Core
  really is a formula-driven, filterable view — not a static table. This settles O.3's open
  question: Core content does depend on the live `C2`/`H2` state, confirmed by the user.
- **Read range is `A4:W500`, not `A4:R89`.** `R89` was both too narrow a column range (misses
  the `N:V` frequency block entirely under the old assumption, and now also misses `W`) and
  too short a row range (this file only shows 60 rows *because `H2='M'` is filtering it* —
  with `H2` blank, row count can grow well past 89). Verified: `W2` header is `'Filter'` — a
  formula-helper column (`W4` duplicates the scenario key from col B), not new scenario data;
  read it for completeness but there's nothing to extract from it beyond A:V.
- **`H2` must be blank when the engine starts a run.** Per the user: leaving it empty returns
  *all* size variants (S/M/L/XL/INST/Total) in column G in one Core read per event, so size
  classification happens client-side (grouping by col G) instead of needing a second
  filter-and-recalculate pass per size. The engine should verify/clear `H2` before each
  `C2` write, not assume it's already empty.

**Bottom line, confirmed by the user:** Core genuinely is live/formula-driven, keyed by
`C2` (Impact ID) and `H2` (size filter, must be cleared). `qra_engine_v2.1.py`
(openpyxl-only, static reads) cannot reproduce this from a single saved `.xlsm` — it needs
`xlwings` + a live Excel session doing the same `C2` write + clear-`H2` + `Calculate()` +
re-read dance as v6, once per active Impact ID from PageControl. Not yet implemented —
next step, pending the user's go-ahead.

---

## P. Dynamic threshold-column matching (`read_therm()` / `read_jf()`)

**Why:** Results-sheet threshold headers (`F1:O1`, and for JF also `P1:Y1`/`Z1:AI1`) are
currently identical to `General!J17:S17`, but there's no live formula link between them —
Results headers come from the PHAST run, `General` is separate workbook config. They
*happen* to agree today (confirmed via the live Excel audit in Section O). Reading `F:O`
positionally, as `read_therm()`/`read_jf()` did before, assumed that agreement always
holds; if it ever didn't, the engine would silently misassign data (e.g. read the 7.3 kW/m²
column as if it were 5.0) with no error at all — worse than a crash.

**Fix:** added `_match_threshold_cols(ws, thresholds, col_start, col_end, ...)`, mirroring
`qra_v6_engine.py`'s function of the same name — for each `KW_THRESHOLDS` value, searches
the given column range for a header within `1e-4`, and returns `{threshold: col_idx}`.
Missing thresholds are logged via `_warn()` and left out of the map (`read_therm`/`read_jf`
then use `NaN` for that threshold's distance, exactly like a column that's blank in the
sheet — never a wrong column). JF needs 3 independent matches (`dist_v`/`half_v`/`cent_v`
blocks) since the same 10 threshold labels repeat across `F:O`, `P:Y`, `Z:AI` — matching the
whole row at once would make every block resolve to the first (`F:O`) occurrence.

**Verified:**
- Full run against the real Macro: identical per-event written/skipped counts to before the
  change (EPF 21/52, LPF 21/68, FB 0/30, FF 60/132, TOXIC 47/86, CVE 5, JF 60/118) — no
  regression, and "No structural warnings" printed (correct, since headers do currently
  match `General` exactly).
- Synthetic mismatch test: built a sheet with headers `[1.6, 5.0, 7.3]` and asked
  `_match_threshold_cols` to match `[1.6, 5.0, 7.3, 9.5]` — correctly matched the first
  three to their columns, left `9.5` out of the map, and logged exactly one warning
  (`"...threshold columns not found in F1:I1: [9.5]"`).

**Warning collector:** added `_warnings`/`_warn()` (mirrors `qra_v6_engine.py`'s pattern) —
every `_warn()` call prints immediately *and* is collected. `main()` now ends with a
`print_warning_summary()` block (per-event skip counts + every collected structural
warning) — the single place to check regardless of how the engine is invoked. This is
explicitly forward-looking: the stated end goal is running this engine from Excel and
writing matrices back into it, where scrolled-past per-line console warnings would be
easy to miss — a single end-of-run summary is required for that.
