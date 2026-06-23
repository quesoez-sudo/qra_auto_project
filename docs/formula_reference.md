# QRA Engine — Formula Reference

This document describes the four implemented event types in `qra_engine_v2.py`. For each event it covers:
1. The **Excel kernel formula** (from `KernelV0_v2_copy.xlsx`) and what every parameter means.
2. The **Python implementation** in the engine and the role of each argument.
3. The **physical/engineering formula** that both of the above are expressing.

Grid conventions shared by all events:
- `QX = 315` columns (west → east), `QY = 317` rows (north → south, row 1 = northmost)
- Cell-centre positions: `x = SX · (col − 0.5)`, `y = SY · (QY − row + 0.5)` in metres
- `SX = 1.0698 m/cell`, `SY = 1.0694 m/cell`

---

## 1. Thermal Radiation — `thermal`

**Applies to events:** LPF (pool fire), EPF (pressurised pool fire), FB (fireball)

---

### 1.1 Excel Formula — `ImpactThermMatrix!AB1`

```excel
IFERROR(
  LET(
    area,    ANCHOR:INDEX(ANCHOR_F, AA16, AA15),

    distM,   MAP(area, LAMBDA(n,
               SQRT((AA13*(COLUMN(n)-0.5) - AA5)^2
                  + (AA14*((ROWS(area)-ROW(n))+0.5) - AA6)^2))),

    distV,   INDEX(F:O, MATCH(AA3, B:B, 0),),
    distFV,  FILTER(distV, distV<>""),
    matchM,  MAP(distM, LAMBDA(n, MATCH(n, distV, -1))),
    impactV, F1:O1,

    impactM, MAP(matchM, distM, LAMBDA(y, d,
               IF(d > MAX(distV), 0,
                  IF(d < MIN(distV),
                     INDEX(impactV,, COLUMNS(distFV)),
                     (INDEX(impactV, y+1) - INDEX(impactV, y))
                     / (INDEX(distV, y+1)  - INDEX(distV, y))
                     * (d - INDEX(distV, y)) + INDEX(impactV, y))))),

    riskT,   MAP(impactM, LAMBDA(k,
               IFERROR(
                 AA7 * IF(k >= 35, 1,
                          0.5*(1 + ERF(
                            ((-36.38 + 2.56*LN((1000*k)^(4/3)*AA8)) - 5)
                            / SQRT(2)))),
                 0))),

    IF(AA2 = 0, impactM, riskT)
  ),
-1)
```

**Parameter reference:**

| Cell | Name | Description |
|------|------|-------------|
| `AA2` | mode | `0` = output kW/m² impact matrix; `1` = output risk matrix |
| `AA3` | scenario key | Identifier matched against col B of `ImpactThermMatrix` to select the active row |
| `AA5` | srcX | Source X coordinate in metres (from `Core` sheet via `INDEX/MATCH`) |
| `AA6` | srcY | Source Y coordinate in metres |
| `AA7` | frequency | Event probability (e.g. P_LPF) in events/year from `Core` |
| `AA8` | expTime | Exposure time in seconds (typically 20 s) |
| `AA13` | SX | Grid cell width in m/cell |
| `AA14` | SY | Grid cell height in m/cell |
| `AA15` | QX | Number of grid columns (315) |
| `AA16` | QY | Number of grid rows (317) |
| `F:O` | distV | 10-column table of radial distances (m) at which intensity drops to each threshold level, read from the matched scenario row |
| `F1:O1` | impactV | kW/m² threshold labels: `[1.6, 5.0, 7.3, 9.5, 12.5, 16.0, 20.9, 25.0, 30.0, 35.0]` |
| `ANCHOR` | — | Named range pointing to `ANCHOR!A1`; used to build a 317×315 spill area for the MAP array formula |

**What the formula does (step by step):**
1. Builds a 317×315 grid of Euclidean distances from the source (`distM`).
2. Looks up the active scenario row's 10 threshold distances (`distV`).
3. For each grid cell, linearly interpolates kW/m² from `distV`/`impactV` — returns 0 beyond the outermost ring and the maximum value inside the innermost ring.
4. If mode = 1, applies the Eisenberg thermal probit and multiplies by frequency to produce individual risk (events/year × lethality probability).

---

### 1.2 Python Implementation — `qra_engine_v2.compute_thermal`

```python
def compute_thermal(dist_m, distances, exp_time=20.0):
    dv = np.array([float(x) if x not in ('', None) else np.nan for x in distances])
    valid  = ~np.isnan(dv)
    dv_v   = dv[valid]               # distances at which kW drops to each threshold
    kw_v   = KW_THRESHOLDS[valid]    # corresponding kW/m² values

    impact = np.zeros_like(dist_m)

    # Inside all rings → max kW
    impact[dist_m <= dv_v[-1]] = kw_v[-1]

    # Linear interpolation between consecutive rings
    for i in range(len(dv_v) - 1):
        mask = (dist_m > dv_v[i+1]) & (dist_m <= dv_v[i])
        slope = (kw_v[i+1] - kw_v[i]) / (dv_v[i+1] - dv_v[i])
        impact[mask] = kw_v[i] + slope * (dist_m[mask] - dv_v[i])

    return impact   # kW/m²  (zero beyond outermost ring)
```

**Probit & risk conversion (`_thermal_prob` + `_to_risk`):**

```python
def _thermal_prob(kw_mat, exp_time):
    p = np.zeros_like(kw_mat)
    m = (kw_mat > 0) & (kw_mat < 35.0)
    Y = -36.38 + 2.56 * np.log((1000.0 * kw_mat[m])**(4/3) * exp_time)
    p[m] = 0.5 * (1.0 + erf((Y - 5.0) / sqrt(2)))
    p[kw_mat >= 35.0] = 1.0
    return np.clip(p, 0.0, 1.0)

# Risk matrix:
risk = frequency * _thermal_prob(impact, exp_time)
```

**Argument reference:**

| Argument | Type | Description |
|----------|------|-------------|
| `dist_m` | `(QY, QX)` ndarray | Euclidean distance from the source to every grid cell, in metres |
| `distances` | list of 10 floats/NaN | Radial distances (m) from `ImpactThermMatrix` cols F–O for the active scenario; NaN = threshold not reached |
| `exp_time` | float | Exposure time in seconds (default 20 s, matches `AA8`) |
| `KW_THRESHOLDS` | global array | `[1.6, 5.0, 7.3, 9.5, 12.5, 16.0, 20.9, 25.0, 30.0, 35.0]` kW/m² |
| `frequency` | float | Event probability in events/year (P_LPF from Core) |

---

### 1.3 Physical Formula

**Step 1 — Intensity at distance d (linear interpolation):**

The kernel provides radiation isopleth radii from consequence modelling (e.g. Phast, SAFETI). Between two consecutive isopleths at distances d₁ and d₂ with intensities I₁ and I₂:

$$I(d) = I_1 + \frac{I_2 - I_1}{d_2 - d_1} \cdot (d - d_1)$$

**Step 2 — Thermal probit (Eisenberg, 1975):**

$$Y = -36.38 + 2.56 \ln\!\left[(1000 \cdot I)^{4/3} \cdot t\right]$$

where:
- `I` = thermal radiation flux (kW/m²)
- `t` = exposure time (s), typically 20 s (person outdoors, delayed evacuation)
- `Y` = probit unit (dimensionless)

**Step 3 — Lethality probability:**

$$P_{lethal} = \Phi\!\left(\frac{Y - 5}{\sqrt{2}}\right) = \frac{1}{2}\left[1 + \text{erf}\!\left(\frac{Y-5}{\sqrt{2}}\right)\right]$$

where Φ is the standard normal CDF.

**Step 4 — Individual Risk:**

$$IR = f_{event} \times P_{lethal}(I, t)$$

where `f_event` is the event frequency (events/year). The result is the annual probability of fatality at each grid cell.

---

## 2. Flash Fire — `ff`

**Applies to event:** FF (flash fire from flammable cloud ignition)

---

### 2.1 Excel Formula — `ImpactFFMatrix`

The formula uses the same `LET + MAP` grid structure as the thermal sheet. Given a cell's distance `d` from the source and the two zone radii:

```excel
impactM = MAP(distM, LAMBDA(d,
              IF(d < LFL_r,  2,
              IF(d < LFLF_r, 1,
                             0))))
```

**Key parameters:**

| Parameter | Source column | Description |
|-----------|--------------|-------------|
| `LFL_r` | col E | Radius (m) of the Lower Flammability Limit cloud — inside this boundary the gas-air mixture is above LFL and will combust |
| `LFLF_r` | col F | Lethal Flash Fire radius (m) — outer zone where thermal exposure from the fire edge can cause fatality |
| `srcX`, `srcY` | Core via AA | Source location in metres |
| `frequency` | P_FF (Core) | Flash fire event probability (events/year) |

---

### 2.2 Python Implementation — `qra_engine_v2.compute_ff`

```python
def compute_ff(dist_m, lfl_r, lflf_r):
    impact = np.zeros_like(dist_m)
    impact[dist_m < lflf_r] = 1.0    # outer zone: partial hazard
    impact[dist_m < lfl_r]  = 2.0    # inner zone: engulfed in flame
    return impact  # integer field: 0 / 1 / 2
```

**Risk conversion:**
```python
risk = frequency * (impact == 2.0).astype(float)
# Only the LFL zone (value 2) contributes to individual risk
```

**Argument reference:**

| Argument | Description |
|----------|-------------|
| `dist_m` | `(QY, QX)` Euclidean distance from source in metres |
| `lfl_r` | LFL cloud radius in metres (from col E of `ImpactFFMatrix`) |
| `lflf_r` | Lethal flash fire radius in metres (from col F) |
| `frequency` | P_FF from Core sheet (events/year) |

---

### 2.3 Physical Formula

A flash fire is the near-instantaneous combustion of a flammable vapour cloud. Its hazard is defined by two concentric zones:

**Zone classification:**

$$\text{Impact}(d) = \begin{cases} 2 & d < r_{LFL} \quad \text{(engulfed — certain fatality)} \\ 1 & r_{LFL} \leq d < r_{LFLF} \quad \text{(edge exposure — conditional fatality)} \\ 0 & d \geq r_{LFLF} \quad \text{(safe)} \end{cases}$$

**Individual risk (conservative binary model):**

$$IR = f_{FF} \times \mathbf{1}[d < r_{LFL}]$$

The LFL radius comes from atmospheric dispersion modelling of the flammable cloud at the ignition moment. Persons inside the LFL boundary are assumed to have probability 1 of fatality (immersed in fire); persons in the outer zone are given probability 0 in the risk calculation (the `1` impact value is retained for consequence mapping only, not risk accumulation).

---

## 3. Toxic Dispersion — `toxic`

**Applies to event:** TOXIC (toxic gas release)

---

### 3.1 Excel Formula — `ImpactToxMatrix`

The toxic sheet stores the consequence result as a pre-computed distance–lethality CSV blob in **column G** of each scenario row (produced by a dispersion model such as Phast or SAFETI). The kernel interpolates directly from this table:

```excel
LET(
  blob,      INDEX(G:G, MATCH(AA3, B:B, 0)),
  distProb,  parsed 2-column table: [distance_km, lethality_prob],
  probM,     MAP(distM, LAMBDA(d,
               INTERP(d / 1000, distProb[distance], distProb[prob]))),
  riskT,     MAP(probM, LAMBDA(p, AA7 * p)),
  IF(AA2 = 0, probM, riskT)
)
```

**Key parameters:**

| Parameter | Source | Description |
|-----------|--------|-------------|
| `AA3` | control cell | Active scenario key, selects the row in `ImpactToxMatrix` |
| `AA5`, `AA6` | Core | Source coordinates in metres |
| `AA7` | Core (P_TOXIC) | Event frequency (events/year) |
| col G blob | each row | Pre-computed 5-column CSV table from dispersion model. Column 1 = distance from source in **metres** (positive values only), Column 4 = lethality probability at that distance (0–1). Rows are ordered from farthest to nearest. |

---

### 3.2 Python Implementation — `qra_engine_v2.compute_toxic`

```python
def _parse_tox_blob(blob_str, min_prob=0.01):
    # Parse 5-column CSV, keep rows where col1 >= 0 and col4 >= min_prob
    arr     = np.array(rows)
    keep    = (arr[:, 0] >= 0) & (arr[:, 3] >= min_prob)
    j_col   = filtered[:, 0]   # distance-like (km), descending
    m_col   = filtered[:, 3]   # lethality probability, descending
    return j_col, m_col

def compute_toxic(dist_m, blob_str, unit_mode='km'):
    j_col, m_col = _parse_tox_blob(blob_str)
    d = dist_m / 1000.0 if unit_mode == 'km' else dist_m

    # np.interp requires ascending x; reverse the descending arrays
    prob = np.interp(d, j_col[::-1], m_col[::-1],
                     left=1.0,   # closer than nearest point → P=1
                     right=0.0)  # farther than farthest point → P=0
    return np.clip(prob, 0.0, 1.0)
```

**Risk conversion:**
```python
risk = frequency * prob_matrix
# Toxic impact IS already lethality probability → multiply directly by f
```

**Argument reference:**

| Argument | Description |
|----------|-------------|
| `dist_m` | `(QY, QX)` Euclidean distance from source in metres |
| `blob_str` | Raw text from `ImpactToxMatrix` col G — 5-column CSV from the dispersion model |
| `unit_mode` | `'meters'` (default, confirmed) → compare `dist_m` directly against blob col 1, both in metres; `'km'` → divide by 1000 (incorrect for this kernel) |
| `min_prob` | Minimum lethality probability to include (filters noise rows from blob, default 0.01) |
| `frequency` | P_TOXIC from Core sheet (events/year) |

---

### 3.3 Physical Formula

Toxic consequence modelling follows a two-step chain: atmospheric dispersion → dose-response.

**Step 1 — Atmospheric dispersion:** A dedicated model (Phast/SAFETI) computes concentration `C(d)` at each downwind distance `d` for a given release rate, atmospheric stability, and wind speed. This step is performed **outside** the QRA engine; only its final distance–probability output is consumed.

**Step 2 — Dose-response (Probit or direct probability):** Using a substance-specific LC₅₀ or probit function, the concentration is converted to lethality probability. In this kernel the conversion is already embedded in the blob:

$$P_{lethal}(d) = \text{interpolate from blob}(d)$$

The blob provides a monotonically decreasing function: probability is 1 (or near 1) at zero distance and approaches 0 at the far field.

**Individual Risk:**

$$IR = f_{TOXIC} \times P_{lethal}(d)$$

Unlike thermal events, no additional probit calculation is needed at the grid stage because the dose-response is already resolved inside the dispersion model output.

---

## 4. Jet Fire — `jf`

**Applies to event:** JF (high-pressure jet fire)

---

### 4.1 Excel Formula — `ImpactJFMatrix`

The jet fire is non-symmetric: the flame elongates in the wind direction. The kernel models it as overlapping rotated ellipses — one ellipse per kW/m² threshold level, per wind direction — and assigns to each cell the maximum kW level whose ellipse it falls inside.

```excel
LET(
  -- For each threshold id (1..10) and each wind direction θ:
  armD,    distV[id] - centV[id],      -- forward semi-axis
  halfW,   halfWV[id],                 -- lateral semi-axis
  centXY,  centV[id] * [cos(θ), sin(θ)],  -- ellipse centre in cell units

  -- Vector from ellipse centre to each cell:
  dxGrid,  x_c - centXY.x,
  dyGrid,  y_c - centXY.y,

  -- Rotated quadratic form (= 1 on ellipse boundary):
  Q,       (dxGrid*cos(θ) + dyGrid*sin(θ))^2 / armD^2
         + (dxGrid*sin(θ) - dyGrid*cos(θ))^2 / halfW^2,

  insideM, Q <= 1,

  -- Maximum kW level across all (id, θ) combinations:
  impactM, MAX over all (id, θ) of: IF(insideM, kW[id], 0)
)
```

**Key parameters from `ImpactJFMatrix` row (cols F–AI):**

| Columns | Array name | Description |
|---------|-----------|-------------|
| F–O (10 cols) | `distV` | Forward-tip distance of the kW ellipse from the source, in **cell units** (not metres). One value per kW/m² threshold. |
| P–Y (10 cols) | `halfWV` | Lateral half-width of the kW ellipse in cell units |
| Z–AI (10 cols) | `centV` | Distance from the source to the **centre** of the ellipse in cell units |
| E | `flame_len` | Total flame length in metres (informational) |
| `n_dirs` = 8 | — | Number of equally-spaced wind directions (0°, 45°, 90°, …, 315°) |

---

### 4.2 Python Implementation — `qra_engine_v2.compute_jf`

```python
def compute_jf(x_c, y_c, dist_v, half_v, cent_v,
               n_dirs=8, angle_offset=0.0, exp_time=20.0):

    # Build angle list: 8 equally-spaced directions
    angles_rad = np.radians(angle_offset + np.arange(n_dirs) / n_dirs * 360.0)

    for id_idx in range(len(kw_v)):
        d   = dv_v[id_idx]    # forward tip distance (cell units)
        hw  = hv_v[id_idx]    # lateral half-width   (cell units)
        c   = cv_v[id_idx]    # centre distance from source (cell units)
        arm = d - c           # forward semi-axis of ellipse
        kw  = kw_v[id_idx]   # kW/m² for this level

        a = 1.0 / arm**2     # 1/a² in the quadratic form
        b = 1.0 / hw**2      # 1/b²

        for theta in angles_rad:
            ct, st = cos(theta), sin(theta)

            # Displacement from ellipse centre to each grid cell
            dx = x_c - c * ct    # (1, QX) broadcasts → (QY, QX)
            dy = y_c - c * st

            # Ellipse quadratic form Q ≤ 1 means "inside"
            Q = (dx*ct + dy*st)**2 * a + (dx*st - dy*ct)**2 * b
            inside_any |= (Q <= 1.0)

        impact = np.maximum(impact, inside_any * kw)

    # Mask cells beyond maximum forward reach
    impact[dist_cells > d_maxim] = 0.0
    return impact
```

**Risk conversion (same probit as thermal):**
```python
risk = frequency * _thermal_prob(impact, exp_time)
# kW/m² → probit → lethality probability → × f_JF
```

**Argument reference:**

| Argument | Description |
|----------|-------------|
| `x_c` | `(1, QX)` array — X displacement in **cell units** from source to each column |
| `y_c` | `(QY, 1)` array — Y displacement in **cell units** from source to each row |
| `dist_v` | 10-element list — forward tip distance of each kW ellipse in cell units (NaN = threshold not reached) |
| `half_v` | 10-element list — lateral half-width of each kW ellipse in cell units |
| `cent_v` | 10-element list — centre offset from source to each ellipse in cell units |
| `n_dirs` | Number of wind directions (default 8, uniformly spaced at 45° intervals) |
| `angle_offset` | Starting angle in degrees (default 0° = East) |
| `exp_time` | Exposure time in seconds for the probit calculation (default 20 s) |
| `frequency` | P_JF from Core sheet (events/year) |

> **Note on cell units:** The JF ellipse parameters (`dist_v`, `half_v`, `cent_v`) are stored in the kernel as multiples of `SX`/`SY` cell widths, not metres. The displacement arrays `x_c` and `y_c` are also in cell units, so the ellipse equation is dimensionally consistent without unit conversion.

---

### 4.3 Physical Formula

A high-pressure jet fire projects an elongated, wind-stabilised flame. The thermal radiation field is non-circular; it follows the flame envelope. The standard engineering model represents constant-intensity contours as **ellipses rotated to the wind direction**.

**Ellipse geometry per wind direction θ:**

For a flame pointing in direction θ with:
- `c` = distance from release point to the flame's centroid
- `a = d − c` = forward semi-axis (from centroid to flame tip)
- `b = hw` = lateral semi-axis (half-width at widest point)

A grid cell at displacement `(x, y)` from the source (in cell units) is inside the `I`-kW/m² contour if:

$$\frac{\left[(x - c\cos\theta)\cos\theta + (y - c\sin\theta)\sin\theta\right]^2}{(d-c)^2} + \frac{\left[(x - c\cos\theta)\sin\theta - (y - c\sin\theta)\cos\theta\right]^2}{hw^2} \leq 1$$

**Multi-direction superposition:**

Since the wind direction at the time of the event is uncertain, all 8 directions are given equal weight by assigning the maximum kW level across all directions:

$$I_{cell} = \max_{\theta \in \{0°, 45°, \ldots, 315°\}} \max_{id} \left[ kW_{id} \cdot \mathbf{1}_{cell \in E_{id,\theta}} \right]$$

**Lethality and individual risk (same model as thermal):**

$$Y = -36.38 + 2.56 \ln\!\left[(1000 \cdot I)^{4/3} \cdot t\right]$$

$$IR = f_{JF} \times \Phi\!\left(\frac{Y-5}{\sqrt{2}}\right)$$

The jet fire radiation probit is identical to the pool fire probit (Eisenberg model) because both ultimately impose thermal flux on the person; the difference lies in the spatial shape of the intensity field, not the dose-response function.

---

## Summary Table

| Event | Impact output | Risk conversion | Source of geometry |
|-------|--------------|-----------------|-------------------|
| Thermal (LPF/EPF/FB) | kW/m² via linear interp from 10 radial isopleths | Eisenberg probit → × frequency | Circular isopleths from consequence model |
| Flash Fire (FF) | 0 / 1 / 2 zone indicator | 1 inside LFL × frequency | Two radii (LFL, LFLF) from dispersion model |
| Toxic | Lethality probability (0–1) via interp from blob | Direct × frequency (no extra probit) | Distance–probability table from dispersion model |
| Jet Fire (JF) | kW/m² from maximum over 8 rotated ellipses | Eisenberg probit → × frequency | Ellipse params (tip, half-width, centre) per kW level from flame model |

All individual risk outputs share the same unit: **annual probability of fatality** (dimensionless, typically 10⁻⁶ to 10⁻³ range) and are accumulated across scenarios to build the total risk matrices.
