# ImpactThermMatrix — AB1 Formula Analysis

## Raw Formula (cell AB1, array spilling dynamically based on QY×QX)

```excel
IFERROR(
  LET(
    area,     ANCHOR:INDEX(ANCHOR_F, $AA$14, $AA$13),
    distM,    MAP(area, LAMBDA(n,
                SQRT(
                  ($AA$11*(COLUMN(n)-0.5) - $AA$5)^2 +
                  ($AA$12*((ROWS(area)-ROW(n))+0.5) - $AA$6)^2
                )
              )),
    distV,    INDEX($G:$P, MATCH($AA$3,$B:$B,0), ),
    matchM,   MAP(distM, LAMBDA(n, MATCH(n, distV, -1))),
    impactV,  $G$1:$P$1,
    impactM,  MAP(matchM, distM, LAMBDA(y, d,
                IF(d > MAX(distV), 0,
                  IF(d < MIN(distV), $P$1,
                    (INDEX(impactV, y+1) - INDEX(impactV, y)) /
                    (INDEX(distV, y+1) - INDEX(distV, y)) *
                    (d - INDEX(distV, y)) + INDEX(impactV, y)))
              )),
    riskT,    MAP(impactM, LAMBDA(k,
                IFERROR(
                  $AA$7 * IF(k >= 35, 1,
                    0.5 * (1 + ERF(
                      ((-36.38 + 2.56 * LN((1000*k)^(4/3) * $AA$8)) - 5) / SQRT(2)
                    ))
                  ), 0)
              )),
    IF($AA$2=0, impactM, riskT)
  ),
-1)
```

---

## Formula Step-by-Step Logic

### Step 1: `area` — Build the spatial grid

```
area = ANCHOR:INDEX(ANCHOR_F, $AA$14, $AA$13)
```

Creates a **QY × QX** matrix using the `ANCHOR` named range as origin.

### Step 2: `distM` — Distance matrix

```
distM = MAP(area, LAMBDA(n,
  SQRT(($AA$11*(COLUMN(n)-0.5) - $AA$5)^2 + ($AA$12*((ROWS(area)-ROW(n))+0.5) - $AA$6)^2)
))
```

For each grid cell, the Euclidean distance (m) from the cell centre to the source point (X, Y).

- `n` = iterator variable representing each cell in the grid
- The distance is what gets compared against the radiation-distance profile

### Step 3: `distV` — Look up scenario distance profile

```
distV = INDEX($G:$P, MATCH($AA$3, $B:$B, 0), )
```

- Finds the row in column B matching the scenario in `$AA$3`
- Returns the **entire row from columns G to P** — these are the **distances (m)** at which each thermal radiation threshold is reached
- Result: a 1×10 array of distances (decreasing — G is closest/highest radiation, P is farthest/lowest)

**Important:** Values can be numeric or "Not reached at height of interest" (meaning that radiation level is never reached).

### Step 4: `matchM` — Find interpolation bracket

```
matchM = MAP(distM, LAMBDA(n, MATCH(n, distV, -1)))
```

- `n` = the distance from one grid cell to the source (from `distM`)
- `MATCH(n, distV, -1)` = finds the position in `distV` where `n` falls (descending search)
- Returns the index of the interval bracket for linear interpolation

### Step 5: `impactV` — Radiation threshold values

```
impactV = $G$1:$P$1
```

The **row 1 headers** of columns G through P. These are the thermal radiation levels in kW/m²:

| Column | G | H | I | J | K | L | M | N | O | P |
|--------|---|---|---|---|---|---|---|---|---|---|
| Value (kW/m²) | 1.6 | 5 | 7.3 | 9.5 | 12.5 | 16 | 20.9 | 25 | 30 | 35 |

### Step 6: `impactM` — Interpolate radiation at each cell

```
impactM = MAP(matchM, distM, LAMBDA(y, d,
  IF(d > MAX(distV), 0,
    IF(d < MIN(distV), $P$1,
      (INDEX(impactV,y+1) - INDEX(impactV,y)) /
      (INDEX(distV,y+1) - INDEX(distV,y)) *
      (d - INDEX(distV,y)) + INDEX(impactV,y)
    ))
))
```

- `y` = the bracket index from `matchM` (which interval the cell falls in)
- `d` = the actual distance from `distM` (in metres)
- Logic:
  - If `d > MAX(distV)` → cell is beyond the farthest radiation contour → **radiation = 0**
  - If `d < MIN(distV)` → cell is closer than the nearest contour → **radiation = $P$1 = 35 kW/m²** (maximum)
  - Otherwise → **linear interpolation** between the two bracketing radiation values

The result is a QY×QX matrix of thermal radiation values (kW/m²) at each grid cell.

### Step 7: `riskT` — Convert radiation to fatality probability (Probit)

```
riskT = MAP(impactM, LAMBDA(k,
  IFERROR(
    $AA$7 * IF(k >= 35, 1,
      0.5 * (1 + ERF(
        ((-36.38 + 2.56 * LN((1000*k)^(4/3) * $AA$8)) - 5) / SQRT(2)
      ))
    ), 0)
))
```

- `k` = thermal radiation value (kW/m²) at that grid cell (from `impactM`)
- If `k >= 35` → fatality probability = 1 (certain death)
- Otherwise applies the **thermal radiation probit equation**:
  - Probit = `-36.38 + 2.56 × LN((1000 × k)^(4/3) × exposure_time)`
  - The probit is converted to probability via: `0.5 × (1 + ERF((Probit - 5) / √2))`
  - `$AA$8` = **Exposure Time** (seconds) — a key parameter in the dose calculation
- The probability is multiplied by the scenario **Probability** (`$AA$7`)
- Errors (e.g. LN of negative) default to 0

### Step 8: Output selection

```
IF($AA$2=0, impactM, riskT)
```

- `$AA$2 = 0` → returns **raw thermal radiation values** (kW/m² matrix)
- `$AA$2 = 1` → returns **risk matrix** (probability × fatality probability)

### Error handling

Outer `IFERROR(..., -1)` — any unhandled error defaults to -1.

---

## Variables Inventory

### Cell References (Parameters in columns Z/AA)

| Cell | Label (col Z) | Current Value | Role in Formula | Description |
|------|---------------|---------------|-----------------|-------------|
| `$AA$2` | Impact/Risk | 1 | Output toggle | 0 = radiation values, 1 = risk (fatality prob × event prob) |
| `$AA$3` | Scenario | D7204_gas/ST1/ROG_FL30/NA/5.5mNoche | MATCH key | Scenario identifier to look up in column B |
| `$AA$5` | X | 30 | Distance calc | X-coordinate of the source point [m] |
| `$AA$6` | Y | 106.4125 | Distance calc | Y-coordinate of the source point [m] |
| `$AA$7` | Probability | 0.001 | Risk calc | Scenario event probability/frequency |
| `$AA$8` | Exp. Time | 20 | Probit equation | Exposure time [s] — used in thermal dose calculation |
| `$AA$11` | SX | 1.0698 | Distance calc | Cell size in X direction [m/cell] |
| `$AA$12` | SY | 1.0694 | Distance calc | Cell size in Y direction [m/cell] |
| `$AA$13` | QX | 1 | Grid size | Number of columns in the output matrix |
| `$AA$14` | QY | 1 | Grid size | Number of rows in the output matrix |

### Named Ranges

| Name | Definition | Role in Formula |
|------|-----------|-----------------|
| `ANCHOR` | `ANCHOR!$A$1` | Origin cell for the spatial grid |
| `ANCHOR_F` | `ANCHOR!$1:$1048576` | Full ANCHOR sheet (used with INDEX to define grid extent) |
| `THERMAT_F` | `OFFSET(ImpactThermMatrix!$AB$1,,, $AA$14, $AA$13)` | External reference to this formula's output (used by SumMatrix) |

### Data Columns (scenario table, rows 2–121)

| Column | Header | Role in Formula | Description |
|--------|--------|-----------------|-------------|
| B | ScenarioWeather | Match key — filtered against `$AA$3` | Composite key: Scenario + Weather |
| C | Path | Not used in formula | Folder path of the scenario |
| D | Scenario | Not used in formula | Scenario name (without weather) |
| E | Weather | Not used in formula | Weather condition (Dia/Noche) |
| F | Flame length [m] | Not used in formula | Flame length from PHAST |
| G | 1.6 (kW/m²) | `distV` column 1 | Distance [m] where 1.6 kW/m² radiation is reached |
| H | 5 (kW/m²) | `distV` column 2 | Distance [m] where 5 kW/m² radiation is reached |
| I | 7.3 (kW/m²) | `distV` column 3 | Distance [m] where 7.3 kW/m² radiation is reached |
| J | 9.5 (kW/m²) | `distV` column 4 | Distance [m] where 9.5 kW/m² radiation is reached |
| K | 12.5 (kW/m²) | `distV` column 5 | Distance [m] where 12.5 kW/m² radiation is reached |
| L | 16 (kW/m²) | `distV` column 6 | Distance [m] where 16 kW/m² radiation is reached |
| M | 20.9 (kW/m²) | `distV` column 7 | Distance [m] where 20.9 kW/m² radiation is reached |
| N | 25 (kW/m²) | `distV` column 8 | Distance [m] where 25 kW/m² radiation is reached |
| O | 30 (kW/m²) | `distV` column 9 | Distance [m] where 30 kW/m² radiation is reached |
| P | 35 (kW/m²) | `distV` column 10 | Distance [m] where 35 kW/m² radiation is reached |
| Q | Frustum tip width | Not used in formula | Flame geometry parameter |
| R | OUTPUT | Not used in formula | CSV string of computed output distances |

**Note:** Columns G–P contain **distances in descending order** (G = farthest = lowest radiation, P = closest = highest radiation). Values may be numeric or "Not reached at height of interest" when that radiation level is never achieved.

### Internal (LET) Variables

| Variable | Type | Size | Description |
|----------|------|------|-------------|
| `area` | Range/Array | QY × QX | Spatial grid built from ANCHOR |
| `distM` | Array | QY × QX | Euclidean distance from each cell to source point |
| `distV` | Array | 1 × 10 | Distances at which each radiation threshold is reached (from matched scenario row) |
| `matchM` | Array | QY × QX | Interpolation bracket index for each cell |
| `impactV` | Array | 1 × 10 | Radiation threshold values: [1.6, 5, 7.3, 9.5, 12.5, 16, 20.9, 25, 30, 35] kW/m² |
| `impactM` | Array | QY × QX | Interpolated thermal radiation at each cell (kW/m²) |
| `riskT` | Array | QY × QX | Fatality probability × event probability at each cell |

---

## Probit Equation Detail

The thermal radiation probit model converts radiation intensity + exposure time into fatality probability:

$$\text{Probit} = -36.38 + 2.56 \times \ln\left((1000 \times k)^{4/3} \times t_{exp}\right)$$

Where:
- $k$ = thermal radiation intensity [kW/m²]
- $t_{exp}$ = exposure time [s] (from `$AA$8`)

Conversion to probability:

$$P_{fatality} = 0.5 \times \left(1 + \text{erf}\left(\frac{\text{Probit} - 5}{\sqrt{2}}\right)\right)$$

If $k \geq 35$ kW/m², probability is set to 1 (bypasses probit).

---

## Data Flow Diagram

```
$AA$3 (Scenario) ────► MATCH in col B ──► INDEX($G:$P) ──► distV [10 distances]
                                                                    │
ANCHOR + $AA$13/$AA$14 ──► area (grid)                              │
         │                                                          │
$AA$11, $AA$12 (SX, SY)                                            │
$AA$5, $AA$6 (X, Y) ────► distM (distances)                        │
                                │                                   │
                                ▼                                   ▼
                          matchM = MATCH(distM, distV, -1)
                                │
                                ▼
                     $G$1:$P$1 (impactV = radiation thresholds)
                                │
                                ▼
                     impactM = linear_interpolation(matchM, distM, distV, impactV)
                                │
                    ┌───────────┴───────────┐
                    │                       │
              $AA$2 = 0                $AA$2 = 1
                    │                       │
              OUTPUT: impactM         $AA$7 (Probability)
              (kW/m² radiation)       $AA$8 (Exp. Time)
                                            │
                                            ▼
                                      riskT = Prob × Probit(k, t_exp)
                                            │
                                      OUTPUT: riskT
```

---

## Key Difference from ImpactFFMatrix

| Aspect | FF (Flash Fire) | Thermal |
|--------|-----------------|---------|
| Distance data | 2 columns (LFL, LFL Fraction) | 10 columns (G–P), one per radiation level |
| Impact model | Binary/zone classification (0/1/2) | Linear interpolation → continuous kW/m² |
| Risk conversion | Simple: Prob × (inside LFL) | Probit equation with exposure time |
| Extra parameter | — | `$AA$8` Exposure Time [s] |
| Data values can be | Always numeric | Numeric or "Not reached at height of interest" |
