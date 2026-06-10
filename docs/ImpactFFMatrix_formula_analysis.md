# ImpactFFMatrix — AB1 Formula Analysis

## Raw Formula (cell AB1, array spilling to AB1:AK10)

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
    distV,    CHOOSECOLS(FILTER($B:$G, $B:$B=$AA$3), COLUMNS($B$1:$F$1), COLUMNS($B$1:$G$1)),
    impactM,  MAP(distM, LAMBDA(n,
                IF(INDEX(distV,1,1) > n, $AA$17,
                  IF(INDEX(distV,1,2) < n, $AA$15, $AA$16))
              )),
    riskFF,   MAP(impactM, LAMBDA(f,
                IFERROR($AA$7 * IF(f=2, 1, 0), 0)
              )),
    IF($AA$2=0, impactM, riskFF)
  ),
0)
```

---

## Formula Step-by-Step Logic

### Step 1: `area` — Build the spatial grid

```
area = ANCHOR:INDEX(ANCHOR_F, $AA$14, $AA$13)
```

Creates a **QY × QX** matrix (rows × columns) using the `ANCHOR` named range as origin. This grid represents the physical area being evaluated.

### Step 2: `distM` — Distance matrix

```
distM = MAP(area, LAMBDA(n,
  SQRT(
    ($AA$11 * (COLUMN(n) - 0.5) - $AA$5)^2 +
    ($AA$12 * ((ROWS(area) - ROW(n)) + 0.5) - $AA$6)^2
  )
))
```

For each cell in the grid, calculates the **Euclidean distance** from the cell center to the source/release point. The coordinate system:

- **X-coord of cell** = `SX * (column_index - 0.5)`
- **Y-coord of cell** = `SY * ((total_rows - row_index) + 0.5)` (Y increases upward)
- Then subtracts the source point (X, Y) and computes Euclidean distance.

### Step 3: `distV` — Look up scenario distances

```
distV = CHOOSECOLS(FILTER($B:$G, $B:$B=$AA$3), COLUMNS($B$1:$F$1), COLUMNS($B$1:$G$1))
```

- `FILTER($B:$G, $B:$B=$AA$3)` → finds the row in the data table matching the current scenario.
- `CHOOSECOLS(..., 5, 6)` → extracts columns F and G from that row (since `COLUMNS($B$1:$F$1)=5` and `COLUMNS($B$1:$G$1)=6`).
- **Result**: a 1×2 array = `[LFL_distance, LFL_Fraction_distance]`

### Step 4: `impactM` — Classify each cell by zone

```
impactM = MAP(distM, LAMBDA(n,
  IF(INDEX(distV,1,1) > n, $AA$17,      ← cell is INSIDE LFL cloud → limit3 = 2
    IF(INDEX(distV,1,2) < n, $AA$15,    ← cell is OUTSIDE LFL Fraction → limit1 = 0
      $AA$16))                           ← cell is BETWEEN LFL and LFL Fraction → limit2 = 1
))
```

Here, `n` is a single scalar value from `distM`: the distance (in meters) from one matrix cell to the source point `(X,Y)`. The `MAP` function evaluates this for every cell in the matrix.

- `INDEX(distV,1,1)` is the scenario LFL distance (column F)
- `INDEX(distV,1,2)` is the scenario LFL Fraction distance (column G)
- The test is done as:
  - if `n < LFL` -> inside cloud -> `limit3`
  - else if `n > LFL Fraction` -> outside -> `limit1`
  - else -> between both boundaries -> `limit2`

Produces an integer classification matrix:
| Zone | Condition | Value | Meaning |
|------|-----------|-------|---------|
| Inside LFL cloud | `distance < LFL` | 2 (limit3) | Fully inside flammable envelope |
| Transition zone | `LFL ≤ distance ≤ LFL Fraction` | 1 (limit2) | Between LFL and LFL Fraction boundary |
| Outside | `distance > LFL Fraction` | 0 (limit1) | No flash fire impact |

### Step 5: `riskFF` — Convert impact to risk

```
riskFF = MAP(impactM, LAMBDA(f,
  IFERROR($AA$7 * IF(f=2, 1, 0), 0)
))
```

- Only cells with impact value = **2** (inside LFL cloud) get risk = `Probability × 1`.
- All other cells (zones 1 and 0) get risk = **0**.
- The transition zone (value=1) does **NOT** contribute to risk in the current formula.

### Step 6: Output selection

```
IF($AA$2=0, impactM, riskFF)
```

- `$AA$2 = 0` → returns the **raw impact classification** (0/1/2 matrix)
- `$AA$2 = 1` → returns the **risk matrix** (probability-weighted)

### Error handling

The entire formula is wrapped in `IFERROR(..., 0)` — any error defaults to 0.

---

## Variables Inventory

### Cell References (Parameters in columns Z/AA)

| Cell | Label (col Z) | Current Value | Role in Formula | Description |
|------|---------------|---------------|-----------------|-------------|
| `$AA$2` | Impact/Risk | 1 | Output toggle | 0 = return impact classification, 1 = return risk |
| `$AA$3` | Scenario | D7203/L7/HYD_FL14/H/1mNoche | FILTER match key | Scenario identifier to look up in column B |
| `$AA$5` | X | 40 | Distance calc | X-coordinate of the source/release point [m] |
| `$AA$6` | Y | 300 | Distance calc | Y-coordinate of the source/release point [m] |
| `$AA$7` | Probability | 1 | Risk calc | Scenario probability/frequency — multiplied by impact to get risk |
| `$AA$11` | SX | 0.3384 | Distance calc | Cell size in X direction [m/cell] |
| `$AA$12` | SY | 0.3381 | Distance calc | Cell size in Y direction [m/cell] |
| `$AA$13` | QX | 10 | Grid size | Number of columns in the output matrix |
| `$AA$14` | QY | 10 | Grid size | Number of rows in the output matrix |
| `$AA$15` | limit1 | 0 | Impact classification | Value assigned to cells OUTSIDE the LFL Fraction boundary |
| `$AA$16` | limit2 | 1 | Impact classification | Value assigned to cells in the TRANSITION zone (between LFL and LFL Fraction) |
| `$AA$17` | limit3 | 2 | Impact classification | Value assigned to cells INSIDE the LFL cloud |

### Named Ranges

| Name | Definition | Role in Formula |
|------|-----------|-----------------|
| `ANCHOR` | `ANCHOR!$A$1` | Origin cell for the spatial grid |
| `ANCHOR_F` | `ANCHOR!$1:$1048576` | Full ANCHOR sheet (used with INDEX to define grid extent) |
| `FFMAT_F` | `OFFSET(ImpactFFMatrix!$AB$1,,, $AA$14, $AA$13)` | External reference to this formula's output (used by SumMatrix) |

### Data Columns (scenario table, rows 2–133)

| Column | Header | Role in Formula |
|--------|--------|-----------------|
| B | ScenarioWeather | Match key — filtered against `$AA$3` |
| F | Distance downwind to LFL [m] | `distV[1,1]` — LFL cloud boundary distance (from PHAST) |
| G | Distance downwind to LFL Fraction [m] | `distV[1,2]` — LFL Fraction boundary distance (from PHAST) |

### Internal (LET) Variables

| Variable | Type | Size | Description |
|----------|------|------|-------------|
| `area` | Range/Array | QY × QX (10×10) | Spatial grid built from ANCHOR |
| `distM` | Array | QY × QX | Euclidean distance from each cell to source point |
| `distV` | Array | 1 × 2 | `[LFL_distance, LFL_Fraction_distance]` for the matched scenario |
| `impactM` | Array | QY × QX | Impact zone classification (0, 1, or 2) |
| `riskFF` | Array | QY × QX | Risk values (Probability × 1 for zone 2, else 0) |

---

## Scenario Iteration (current state)

- The formula evaluates **one scenario at a time** (the one specified in `$AA$3`).
- There are **132 scenarios** in the data table (rows 2–133).
- The goal is to iterate over all scenarios, computing a risk matrix for each, then **summing all matrices** to produce the total Flash Fire risk matrix.
- This summation is handled externally in the `SumMatrix` sheet, which references `FFMAT_F` and adds it element-wise to the cumulative risk matrix (`RISKMAT0`).

---

## Data Flow Diagram

```
$AA$3 (Scenario) ─────► FILTER($B:$G) ──► distV [LFL, LFL_Frac]
                                                     │
ANCHOR + $AA$13/$AA$14 ──► area (grid)               │
         │                                           │
$AA$11, $AA$12 (SX, SY)                             │
$AA$5, $AA$6 (X, Y) ────► distM (distances)         │
                                │                    │
                                ▼                    ▼
                           impactM = classify(distM, distV, limit1/2/3)
                                │
                    ┌───────────┴───────────┐
                    │                       │
              $AA$2 = 0                $AA$2 = 1
                    │                       │
              OUTPUT: impactM         $AA$7 (Probability)
              (categories 0/1/2)            │
                                            ▼
                                      riskFF = Prob × (impactM==2)
                                            │
                                      OUTPUT: riskFF
```
