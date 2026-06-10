# ImpactExpMatrix — AB1 Formula Analysis

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
    distV,    INDEX($J:$N, MATCH($AA$3, $B:$B, 0), ),
    matchM,   MAP(distM, LAMBDA(n, MATCH(n, distV, -1))),
    impactV,  $J$1:$N$1,
    impactM,  MAP(matchM, distM, LAMBDA(y, d,
                IF(d > MAX(distV), 0,
                  IF(d < MIN(distV), 1,
                    (INDEX(impactV, y+1) - INDEX(impactV, y)) /
                    (INDEX(distV, y+1) - INDEX(distV, y)) *
                    (d - INDEX(distV, y)) + INDEX(impactV, y)))
              )),
    riskCVE,  MAP(impactM, LAMBDA(v,
                IFERROR(
                  $AA$7 * IF(v < $AA$15, 0,
                    IF(AND(v >= $AA$15, v < $AA$16), $AA$17, $AA$18)
                  ), 0)
              )),
    IF($AA$2=0, impactM, riskCVE)
  ),
0)
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

### Step 3: `distV` — Look up scenario overpressure-distance profile

```
distV = INDEX($J:$N, MATCH($AA$3, $B:$B, 0), )
```

- Finds the row in column B matching the scenario in `$AA$3`
- Returns the **entire row from columns J to N** — these are the **distances (m)** at which each overpressure level is reached
- Result: a 1×5 array of distances (decreasing — J is farthest/lowest overpressure, N is closest/highest)

### Step 4: `matchM` — Find interpolation bracket

```
matchM = MAP(distM, LAMBDA(n, MATCH(n, distV, -1)))
```

- `n` = distance from one grid cell to the source (from `distM`)
- `MATCH(n, distV, -1)` = finds position in `distV` where `n` falls (descending search)
- Returns the bracket index for linear interpolation

### Step 5: `impactV` — Overpressure threshold values

```
impactV = $J$1:$N$1
```

The **row 1 headers** of columns J through N. These are the overpressure levels in **bar**:

| Column | J | K | L | M | N |
|--------|---|---|---|---|---|
| Value (bar) | 0.04 | 0.1 | 0.35 | 0.5 | 1 |

### Step 6: `impactM` — Interpolate overpressure at each cell

```
impactM = MAP(matchM, distM, LAMBDA(y, d,
  IF(d > MAX(distV), 0,
    IF(d < MIN(distV), 1,
      (INDEX(impactV, y+1) - INDEX(impactV, y)) /
      (INDEX(distV, y+1) - INDEX(distV, y)) *
      (d - INDEX(distV, y)) + INDEX(impactV, y)
    ))
))
```

- `y` = bracket index from `matchM`
- `d` = actual distance from `distM` (in metres)
- Logic:
  - If `d > MAX(distV)` → cell is beyond the farthest overpressure contour → **overpressure = 0**
  - If `d < MIN(distV)` → cell is closer than the nearest contour → **overpressure = 1 bar** (maximum in table, capped at $N$1)
  - Otherwise → **linear interpolation** between the two bracketing overpressure values

The result is a QY×QX matrix of **overpressure values (bar)** at each grid cell.

### Step 7: `riskCVE` — Convert overpressure to risk via thresholds

```
riskCVE = MAP(impactM, LAMBDA(v,
  IFERROR(
    $AA$7 * IF(v < $AA$15, 0,
      IF(AND(v >= $AA$15, v < $AA$16), $AA$17, $AA$18)
    ), 0)
))
```

- `v` = overpressure value (bar) at that grid cell (from `impactM`)
- Applies a **step function** with two thresholds:
  - If `v < limitOV1` (0.1 bar) → **fatality factor = 0** (no significant harm)
  - If `limitOV1 ≤ v < limitOV2` (0.1 ≤ v < 0.3 bar) → **fatality factor = limitF1 = 0** (injury zone, no fatality)
  - If `v ≥ limitOV2` (≥ 0.3 bar) → **fatality factor = limitF2 = 1** (fatality zone)
- The fatality factor is multiplied by the scenario **Probability** (`$AA$7`)
- Errors default to 0

**Summary of threshold logic:**

| Condition | Overpressure range | Fatality factor | Meaning |
|-----------|-------------------|-----------------|---------|
| `v < limitOV1` | < 0.1 bar | 0 | Below harm threshold |
| `limitOV1 ≤ v < limitOV2` | 0.1 – 0.3 bar | limitF1 = 0 | Injury zone (no fatality by default) |
| `v ≥ limitOV2` | ≥ 0.3 bar | limitF2 = 1 | Fatality zone |

### Step 8: Output selection

```
IF($AA$2=0, impactM, riskCVE)
```

- `$AA$2 = 0` → returns **raw overpressure values** (bar matrix)
- `$AA$2 = 1` → returns **risk matrix** (probability × fatality factor)

### Error handling

Outer `IFERROR(..., 0)` — any unhandled error defaults to 0.

---

## Variables Inventory

### Cell References (Parameters in columns Z/AA)

| Cell | Label (col Z) | Current Value | Role in Formula | Description |
|------|---------------|---------------|-----------------|-------------|
| `$AA$2` | Impact/Risk | 0 | Output toggle | 0 = overpressure values, 1 = risk |
| `$AA$3` | Scenario | D7301/150mm/REF_FL16/H/1mNoche | MATCH key | Scenario identifier to look up in column B |
| `$AA$5` | X | 10 | Distance calc | X-coordinate of the source point [m] |
| `$AA$6` | Y | 50 | Distance calc | Y-coordinate of the source point [m] |
| `$AA$7` | Probability | 0.01 | Risk calc | Scenario event probability/frequency |
| `$AA$11` | SX | 1.0698 | Distance calc | Cell size in X direction [m/cell] |
| `$AA$12` | SY | 1.0694 | Distance calc | Cell size in Y direction [m/cell] |
| `$AA$13` | QX | 1 | Grid size | Number of columns in the output matrix |
| `$AA$14` | QY | 1 | Grid size | Number of rows in the output matrix |
| `$AA$15` | limitOV1 | 0.1 | Risk threshold | Overpressure threshold 1 (bar) — below this = no harm |
| `$AA$16` | limitOV2 | 0.3 | Risk threshold | Overpressure threshold 2 (bar) — above this = fatality |
| `$AA$17` | limitF1 | 0 | Fatality factor | Factor for intermediate zone (limitOV1 ≤ v < limitOV2) |
| `$AA$18` | limitF2 | 1 | Fatality factor | Factor for fatality zone (v ≥ limitOV2) |

### Named Ranges

| Name | Definition | Role in Formula |
|------|-----------|-----------------|
| `ANCHOR` | `ANCHOR!$A$1` | Origin cell for the spatial grid |
| `ANCHOR_F` | `ANCHOR!$1:$1048576` | Full ANCHOR sheet (used with INDEX to define grid extent) |
| `EXPMAT_F` | `OFFSET(ImpactExpMatrix!$AB$1,,, $AA$14, $AA$13)` | External reference to this formula's output (used by SumMatrix) |

### Data Columns (scenario table, rows 2–67)

| Column | Header | Role in Formula | Description |
|--------|--------|-----------------|-------------|
| B | ScenarioWeather | Match key — used in MATCH against `$AA$3` | Composite key: Scenario + Weather |
| C | Scenario | Not used in formula | Scenario name (without weather) |
| D | Weather | Not used in formula | Weather condition |
| E | Explosion flammable mass [kg] | Not used in formula | Mass of flammable material in the explosion |
| F | Ignition time [s] | Not used in formula | Time to ignition |
| G | Ignition source [m] | Not used in formula | Distance to ignition source |
| H | Cloud centre [m] | Not used in formula | Distance to centre of the flammable cloud |
| I | Explosion centre [m] | Not used in formula | Distance to centre of the explosion |
| **J** | **0.04** (bar) | `distV` column 1, `impactV[1]` | Distance [m] where 0.04 bar overpressure is reached |
| **K** | **0.1** (bar) | `distV` column 2, `impactV[2]` | Distance [m] where 0.1 bar overpressure is reached |
| **L** | **0.35** (bar) | `distV` column 3, `impactV[3]` | Distance [m] where 0.35 bar overpressure is reached |
| **M** | **0.5** (bar) | `distV` column 4, `impactV[4]` | Distance [m] where 0.5 bar overpressure is reached |
| **N** | **1** (bar) | `distV` column 5, `impactV[5]` | Distance [m] where 1 bar overpressure is reached |
| O | 0.04 (bar) | Not used in formula | Distance for secondary set (side-on?) at 0.04 bar |
| P | 0.1 (bar) | Not used in formula | Distance for secondary set at 0.1 bar |
| Q | 0.35 (bar) | Not used in formula | Distance for secondary set at 0.35 bar |
| R | 0.5 (bar) | Not used in formula | Distance for secondary set at 0.5 bar |
| S | 1 (bar) | Not used in formula | Distance for secondary set at 1 bar |
| T | Location Name | Not used in formula | Name of the ignition location/receptor |
| U | Delayed Probability (P2) | Not used in formula | Probability of delayed ignition at this location |
| V | X | Not used in formula | X-coordinate of the ignition/receptor location [m] |
| W | Y | Not used in formula | Y-coordinate of the ignition/receptor location [m] |
| X | S-L DIST | Not used in formula | Source-to-location distance [m] |
| Y | Hole Range | Not used in formula | Hole size category (e.g. "M" for medium) |

**Note on columns J–N vs O–S:** Both sets have the same overpressure headers (0.04, 0.1, 0.35, 0.5, 1 bar). The formula **only uses J:N**. The O:S columns appear to be a secondary set of distances (possibly side-on vs reflected overpressure, or another blast model). They are not used in the AB1 calculation.

**Note on duplicate scenario rows:** The same scenario (column B) can appear in **multiple rows** — once per ignition location (column T). The formula uses `MATCH($AA$3, $B:$B, 0)` which returns the **first** match only.

### Internal (LET) Variables

| Variable | Type | Size | Description |
|----------|------|------|-------------|
| `area` | Range/Array | QY × QX | Spatial grid built from ANCHOR |
| `distM` | Array | QY × QX | Euclidean distance from each cell to source point |
| `distV` | Array | 1 × 5 | Distances at which each overpressure threshold is reached (from matched scenario row) |
| `matchM` | Array | QY × QX | Interpolation bracket index for each cell |
| `impactV` | Array | 1 × 5 | Overpressure threshold values: [0.04, 0.1, 0.35, 0.5, 1] bar |
| `impactM` | Array | QY × QX | Interpolated overpressure at each cell (bar) |
| `riskCVE` | Array | QY × QX | Risk = event probability × fatality factor |

---

## Data Flow Diagram

```
$AA$3 (Scenario) ────► MATCH in col B ──► INDEX($J:$N) ──► distV [5 distances]
                                                                  │
ANCHOR + $AA$13/$AA$14 ──► area (grid)                            │
         │                                                        │
$AA$11, $AA$12 (SX, SY)                                          │
$AA$5, $AA$6 (X, Y) ────► distM (distances)                      │
                                │                                 │
                                ▼                                 ▼
                          matchM = MATCH(distM, distV, -1)
                                │
                                ▼
                     $J$1:$N$1 (impactV = overpressure thresholds)
                                │
                                ▼
                     impactM = linear_interpolation(matchM, distM, distV, impactV)
                                │
                    ┌───────────┴───────────┐
                    │                       │
              $AA$2 = 0                $AA$2 = 1
                    │                       │
              OUTPUT: impactM         $AA$7 (Probability)
              (bar overpressure)      $AA$15 (limitOV1 = 0.1)
                                      $AA$16 (limitOV2 = 0.3)
                                      $AA$17 (limitF1 = 0)
                                      $AA$18 (limitF2 = 1)
                                            │
                                            ▼
                                      riskCVE = Prob × step_function(impactM)
                                            │
                                      OUTPUT: riskCVE
```

---

## Key Differences from Other Impact Sheets

| Aspect | Explosion | FF | Thermal | Toxic |
|--------|-----------|----|---------| ------|
| Distance data | 5 columns (J–N), overpressure contours | 2 values (LFL, LFL Frac) | 10 columns (G–P), radiation contours | Variable-length table (J column) |
| Impact model | Linear interpolation → continuous bar value | Zone classification (0/1/2) | Linear interpolation → kW/m² | Linear interpolation → probability |
| Risk conversion | Step function with 2 thresholds | Simple binary (inside LFL) | Probit equation | Direct probability × event prob |
| Extra parameters | `limitOV1/OV2/F1/F2` thresholds | `limit1/2/3` zone values | Exposure time | `limit1` filter threshold |
| Multiple rows per scenario | **Yes** (one per ignition location) | No | No | No |
| Has X/Y per row | **Yes** (columns V, W) | Added to H, I | Needs to be added | Needs to be added |
| Distance direction | Decreasing (J=far, N=near) | N/A (thresholds) | Decreasing (G=far, P=near) | Decreasing (top=far, bottom=near) |
