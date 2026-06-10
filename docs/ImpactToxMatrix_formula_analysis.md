# ImpactToxMatrix — AB1 Formula Analysis

## Raw Formulas

### Cell AB1 — Main impact/risk matrix (array, spills dynamically based on QY×QX)

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
    matchM,   MAP(distM, LAMBDA(n, MATCH(n, $J:$J, -1))),
    impactM,  MAP(matchM, distM, LAMBDA(y, d,
                IF(d > $J$1, 0,
                  IF(d < MIN($J:$J), 1,
                    (INDEX($M:$M, y+1) - INDEX($M:$M, y)) /
                    (INDEX($J:$J, y+1) - INDEX($J:$J, y)) *
                    (d - INDEX($J:$J, y)) + INDEX($M:$M, y)))
              )),
    riskTOX,  MAP(impactM, LAMBDA(t, IFERROR($AA$7 * t, 0))),
    IF($AA$2=0, impactM, riskTOX)
  ),
0)
```

### Cell AA23 — Fetch scenario raw toxic dispersion table

```excel
AA23 = INDEX($G:$G, MATCH($AA$3, $B:$B, 0))
```

Looks up column G for the matched scenario. Column G contains a **multi-line CSV text blob** — the full toxic dispersion profile for that scenario.

### Cell J1 — Parse and filter the toxic table (array, spills to J1:N77)

```excel
LET(
  detailT,  TEXTSPLIT($AA$23, ",", CHAR(10)) * 1,
  FILTER(detailT,
    (INDEX(detailT,,1) >= 0) * (INDEX(detailT,,4) >= $AA$15)
  )
)
```

- `TEXTSPLIT($AA$23, ",", CHAR(10))` → splits the CSV blob into a 2D table (comma = column separator, newline = row separator)
- `* 1` → coerces text to numeric
- `FILTER(...)` → keeps only rows where:
  - Column 1 (distance) >= 0
  - Column 4 (probability/lethality) >= `$AA$15` (limit1 = 0.01)
- Result: a filtered numeric table written to cells J1:N (variable height)

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
- Same coordinate system as the other impact sheets

### Step 3: `matchM` — Find interpolation bracket in distance column

```
matchM = MAP(distM, LAMBDA(n, MATCH(n, $J:$J, -1)))
```

- `n` = the distance from one grid cell to the source (from `distM`)
- `MATCH(n, $J:$J, -1)` = finds the position in column J (descending distances) where `n` falls
- Column J contains distances in **descending order** (farthest to nearest from the source)
- Returns the row index bracket for linear interpolation

### Step 4: `impactM` — Interpolate fatality probability at each cell

```
impactM = MAP(matchM, distM, LAMBDA(y, d,
  IF(d > $J$1, 0,
    IF(d < MIN($J:$J), 1,
      (INDEX($M:$M, y+1) - INDEX($M:$M, y)) /
      (INDEX($J:$J, y+1) - INDEX($J:$J, y)) *
      (d - INDEX($J:$J, y)) + INDEX($M:$M, y)
    ))
))
```

- `y` = bracket index from `matchM`
- `d` = distance from `distM` (in metres)
- Logic:
  - If `d > $J$1` (farther than the maximum distance in the table) → **probability = 0**
  - If `d < MIN($J:$J)` (closer than the minimum distance) → **probability = 1** (certain fatality)
  - Otherwise → **linear interpolation** between bracketing rows using column J (distance) and column M (probability)

The result is a QY×QX matrix of **fatality probability values** (0 to 1) at each grid cell.

### Step 5: `riskTOX` — Convert to risk

```
riskTOX = MAP(impactM, LAMBDA(t, IFERROR($AA$7 * t, 0)))
```

- `t` = fatality probability at that grid cell (from `impactM`)
- Risk = `Probability × fatality_probability`
- Much simpler than Thermal — no probit, because the toxic dispersion table already provides probability directly

### Step 6: Output selection

```
IF($AA$2=0, impactM, riskTOX)
```

- `$AA$2 = 0` → returns **raw fatality probability** (0 to 1 matrix)
- `$AA$2 = 1` → returns **risk matrix** (event probability × fatality probability)

### Error handling

Outer `IFERROR(..., 0)` — any unhandled error defaults to 0.

---

## Variables Inventory

### Cell References (Parameters in columns Z/AA)

| Cell | Label (col Z) | Current Value | Role in Formula | Description |
|------|---------------|---------------|-----------------|-------------|
| `$AA$2` | Impact/Risk | 0 | Output toggle | 0 = fatality probability, 1 = risk |
| `$AA$3` | Scenario | D7204_gas/ST1/ROG_FL30/NA/5.5mDia | MATCH key | Scenario identifier to look up in column B |
| `$AA$5` | X | 20 | Distance calc | X-coordinate of the source point [m] |
| `$AA$6` | Y | 320 | Distance calc | Y-coordinate of the source point [m] |
| `$AA$7` | Probability | 0.001 | Risk calc | Scenario event probability/frequency |
| `$AA$11` | SX | 0.3384 | Distance calc | Cell size in X direction [m/cell] |
| `$AA$12` | SY | 0.3381 | Distance calc | Cell size in Y direction [m/cell] |
| `$AA$13` | QX | 1 | Grid size | Number of columns in the output matrix |
| `$AA$14` | QY | 1 | Grid size | Number of rows in the output matrix |
| `$AA$15` | limit1 | 0.01 | Table filter | Minimum probability threshold — rows with column 4 < this are excluded from J:N table |
| `$AA$23` | Toxic Dispersion Table | *(computed from AA23 formula)* | Intermediate | Raw CSV text blob for the matched scenario |

### Named Ranges

| Name | Definition | Role in Formula |
|------|-----------|-----------------|
| `ANCHOR` | `ANCHOR!$A$1` | Origin cell for the spatial grid |
| `ANCHOR_F` | `ANCHOR!$1:$1048576` | Full ANCHOR sheet (used with INDEX to define grid extent) |
| `TOXMAT_F` | `OFFSET(ImpactToxMatrix!$AB$1,,, $AA$14, $AA$13)` | External reference to this formula's output (used by SumMatrix) |

### Data Columns (scenario table)

| Column | Header | Role in Formula | Description |
|--------|--------|-----------------|-------------|
| B | SCENARIO_CODE | Match key — used in MATCH against `$AA$3` | Composite key: Scenario + Weather |
| C | SCENARIO | Not used in AB1 formula | Scenario name |
| D | WEATHER | Not used in AB1 formula | Weather condition |
| E | START TABLE ROW | Not used in AB1 formula | Row index for the dispersion table |
| F | TABLE HEIGHT | Not used in AB1 formula | Number of rows in the dispersion table |
| G | Toxic Dispersion Table | Used via AA23 lookup | Multi-line CSV blob containing the full distance-vs-probability profile |
| H | *(no header — formula: SEQUENCE(10000))* | Row numbering helper | Sequential integers 1–10000 |

### Computed Table (columns J–N, dynamic array from J1 formula)

These columns are **computed by the J1 formula** from the raw CSV in column G. They do **NOT** have explicit headers in row 1 — the J1 cell itself is the first data row.

| Column | Row 1 sample value | Role in AB1 Formula | Description |
|--------|-------------------|---------------------|-------------|
| J | 161.603 | Distance for interpolation (`$J:$J`) | **Distance [m]** — Distance from source [m] (descending order) |
| K | 122737000000 | Not used in AB1 formula | **Toxic dose** — *(large numeric values, possibly Ct dose or concentration)* |
| L | 2.70466 | Not used in AB1 formula | **Probit number** — *(coefficient, possibly probit or dose exponent)* |
| M | 0.0108567 | Probability for interpolation (`$M:$M`) | **Probability of fatality** — Fatality/lethality probability at that distance |
| N | 0.772165 | Not used in AB1 formula | **Integrated probability of fatality** — *(another probability or metric)* |

> **Note:** Only columns J and M are used in the AB1 interpolation. Columns K, L, and N are computed but not referenced. Please provide the column meanings for the complete documentation.

### Internal (LET) Variables

| Variable | Type | Size | Description |
|----------|------|------|-------------|
| `area` | Range/Array | QY × QX | Spatial grid built from ANCHOR |
| `distM` | Array | QY × QX | Euclidean distance from each cell to source point |
| `matchM` | Array | QY × QX | Interpolation bracket index for each cell in column J |
| `impactM` | Array | QY × QX | Interpolated fatality probability at each cell (0 to 1) |
| `riskTOX` | Array | QY × QX | Risk = event probability × fatality probability |

---

## Data Flow Diagram

```
$AA$3 (Scenario) ──► MATCH in col B ──► INDEX($G:$G) ──► AA23 (raw CSV blob)
                                                              │
                                                              ▼
                                                J1 formula: TEXTSPLIT + FILTER
                                                              │
                                                              ▼
                                              J:N computed table (distance, ?, ?, prob, ?)
                                                              │
                                              ┌───────────────┴──────────────┐
                                              │ J column (distances)         │ M column (probability)
                                              │                              │
ANCHOR + $AA$13/$AA$14 ──► area (grid)        │                              │
         │                                    │                              │
$AA$11, $AA$12 (SX, SY)                      │                              │
$AA$5, $AA$6 (X, Y) ────► distM              │                              │
                              │               │                              │
                              ▼               ▼                              ▼
                        matchM = MATCH(distM, $J:$J, -1)
                              │
                              ▼
                    impactM = linear_interpolation(y, d, $J:$J, $M:$M)
                              │
                  ┌───────────┴───────────┐
                  │                       │
            $AA$2 = 0                $AA$2 = 1
                  │                       │
            OUTPUT: impactM         $AA$7 (Probability)
            (fatality prob 0–1)           │
                                          ▼
                                    riskTOX = Prob × impactM
                                          │
                                    OUTPUT: riskTOX
```

---

## Key Differences from Other Impact Sheets

| Aspect | Toxic | FF | Thermal |
|--------|-------|----|---------| 
| Data source | CSV blob in column G, parsed at runtime | Direct numeric columns F, G | Direct numeric columns G–P |
| Distance profile | Variable-length table (J column, descending) | 2 fixed values (LFL, LFL Frac) | 10 fixed values (G–P) |
| Impact model | Linear interpolation on probability directly | Zone classification (0/1/2) | Linear interpolation on radiation |
| Risk conversion | Simple: Prob × probability | Simple: Prob × (zone==2) | Probit equation with exposure time |
| Extra parameters | `$AA$15` filter threshold | `limit1/2/3` zone values | `$AA$8` exposure time |
| Intermediate formula | J1 (parses CSV), AA23 (fetches CSV) | None | None |
