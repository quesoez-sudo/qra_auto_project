# ImpactJFMatrix — AM1 Formula Analysis

## Raw Formula (cell AM1, array spilling over AM1:MO317 = QY × QX)

```excel
IFERROR(
  LET(
    area,     ANCHOR:INDEX(ANCHOR_F, $AL$14, $AL$13),

    rId,      MATCH($AL$3, $B:$B, 0),
    distV,    INDEX($F:$O, rId,),
    halfWV,   INDEX($P:$Y, rId,),
    centerV,  INDEX($Z:$AI, rId,),
    impactV,  $F$1:$O$1,
    impactids, SEQUENCE(COUNT(distV)),

    X, MAP(area, LAMBDA(n, ROUND(($AL$11*(COLUMN(n)-0.5) - $AL$5) / $AL$11, 0))),
    Y, MAP(area, LAMBDA(n, ROUND(($AL$12*((ROWS(area)-ROW(n))+0.5) - $AL$6) / $AL$12, 0))),

    dirs,   SEQUENCE($AL$25, 1, 0),
    angles, $AL$24 + IF($AL$26=0, dirs/$AL$25*360, dirs*$AL$26),
    ctV,    COS(RADIANS(angles)),
    stV,    SIN(RADIANS(angles)),

    MAP(X, Y,
      LAMBDA(x, y,
        MAX(
          MAP(impactids,
            LAMBDA(id,
              LET(
                a,    INDEX(distV,,id) - INDEX(centerV,,id),
                b,    INDEX(halfWV,,id),
                c,    INDEX(centerV,,id),
                imp,  INDEX(impactV,,id),
                eVals,
                  (((x - c*ctV)*ctV + (y - c*stV)*stV)^2 / a^2) +
                  (((x - c*ctV)*stV - (y - c*stV)*ctV)^2 / b^2),
                MAX(--(eVals<=1)) * imp
              )
            )
          )
        )
      )
    )
  ),
-333)
```

---

## Formula Step-by-Step Logic

### Step 1: `area` — Build the spatial grid

```
area = ANCHOR:INDEX(ANCHOR_F, $AL$14, $AL$13)
```

Creates a **QY × QX** matrix using the `ANCHOR` named range as origin (same pattern as other impact matrices).

### Step 2: Look up scenario data row

```
rId    = MATCH($AL$3, $B:$B, 0)
distV  = INDEX($F:$O, rId,)    — far-tip distances for 10 thresholds
halfWV = INDEX($P:$Y, rId,)    — semi-minor axis (half-width) for 10 thresholds
centerV = INDEX($Z:$AI, rId,)  — center distance from source for 10 thresholds
impactV = $F$1:$O$1            — kW/m² threshold values [1.6, 5, 7.3, ..., 35]
```

Each scenario row has **three sets of 10 distances**, one per kW/m² threshold:
- **distV**: How far the kW/m² contour extends from the source along the jet axis (far tip)
- **halfWV**: The perpendicular half-width of the contour (semi-minor axis b)
- **centerV**: Where the centre of the ellipse sits along the jet axis from the source

### Step 3: Compute cell coordinates (cell units, relative to source)

```
X = ROUND((SX*(COLUMN(n)-0.5) - X_source) / SX, 0)
Y = ROUND((SY*(ROWS(area)-ROW(n)+0.5) - Y_source) / SY, 0)
```

- Each cell's X and Y offset from the source is computed in **cell units** (dividing meters by SX/SY).
- ROUND gives integer cell indices (Y increases upward in the Excel convention).
- Source is at X=0, Y=0 in these relative coordinates.

### Step 4: Generate direction angles

```
dirs   = SEQUENCE($AL$25, 1, 0)        — [0, 1, 2, ..., n_dirs-1]
angles = $AL$24 + IF($AL$26=0, dirs/$AL$25*360, dirs*$AL$26)
ctV    = COS(RADIANS(angles))          — cosines, one per direction
stV    = SIN(RADIANS(angles))          — sines, one per direction
```

Default parameters: `n_dirs=8`, `angle_offset=0`, `angle_step=0` (= equal spacing → 0°, 45°, 90°, ..., 315°).

When `$AL$26 = 0`: angles are evenly spaced at `360/n_dirs` degrees.  
When `$AL$26 ≠ 0`: angles are `angle_offset + d × angle_step` for each d in [0, n_dirs-1].

### Step 5: Evaluate ellipse membership for each threshold and direction

For each threshold `id` and for each of the `n_dirs` directions simultaneously:

```
a   = distV[id] - centerV[id]       — semi-major axis (from center to far tip)
b   = halfWV[id]                    — semi-minor axis (perpendicular half-width)
c   = centerV[id]                   — center distance from source along jet axis
imp = impactV[id]                   — kW/m² value for this threshold

eVals = ((x - c*ct)*ct + (y - c*st)*st)^2 / a^2
      + ((x - c*ct)*st - (y - c*st)*ct)^2 / b^2
```

This is the **standard axis-aligned ellipse equation rotated into the jet direction**:

- `(x*ct + y*st - c)` = projection of cell position onto jet axis, minus center distance  
- `(x*st - y*ct)`     = perpendicular component from jet axis

Simplified form of the ellipse condition:

```
(proj_along / a)² + (proj_perp / b)² ≤ 1
```

where:
- `proj_along = x*ct + y*st - c`
- `proj_perp  = x*st - y*ct`

The ellipse is centred at `c` from source along the jet and extends from `c-a` to `c+a`
along the jet axis, with half-width `b` in the perpendicular direction.

**Key:** `eVals` is a vector over all `n_dirs` directions simultaneously.

```
MAX(--(eVals<=1)) * imp
```

Returns `imp` if the cell is inside the ellipse for **any** direction, else 0.

### Step 6: Cell impact = maximum threshold kW/m² across all zones

```
MAX(MAP(impactids, LAMBDA(id, ...)))
```

For each cell (x, y), the outer MAX gives the **highest kW/m² threshold** for which the
cell falls inside at least one directional ellipse. Since higher-kW/m² zones are nested
within lower-kW/m² zones, this yields the "worst-case" thermal intensity for the cell
from all eight fire directions.

### Error handling

Outer `IFERROR(..., -333)` — any unhandled error (e.g. divide-by-zero if a=0) returns −333.

---

## Physical Geometry

For a single jet direction θ:

```
     Source (0,0)
        │
        │   c → (center of ellipse along jet)
        │
        ●────────── jet axis (ct, st)
       ╱|╲
      ╱ | ╲   b = semi-minor axis (half-width perpendicular to jet)
     ╱  |  ╲
    ╱   |   ╲  a = dist - center (semi-major axis along jet)
   ──────────── far tip at distance = dist from source
```

The zone extends from `c - a` to `c + a` along the jet direction and `±b` perpendicular.
Because `c - a = 2c - dist` can be negative, the zone may extend slightly **behind** the
source — consistent with radiation from the full flame body including the near-source region.

---

## Variables Inventory

### Cell References (Parameters in columns AK/AL)

| Cell | Label (col AK) | Default | Role in Formula | Description |
|------|----------------|---------|-----------------|-------------|
| `$AL$1` | impact id | 19 | — | Impact configuration ID |
| `$AL$2` | Impact/Risk | 0 | — | Output mode toggle |
| `$AL$3` | Scenario | (varies) | MATCH key | Scenario key (col B) to look up |
| `$AL$5` | X | (source X m) | Grid origin | X-coordinate of source [m] |
| `$AL$6` | Y | (source Y m) | Grid origin | Y-coordinate of source [m] |
| `$AL$7` | Probability | (varies) | — | Scenario frequency |
| `$AL$11` | SX | 1.0698 | Cell spacing | Metres per cell, X direction |
| `$AL$12` | SY | 1.069425 | Cell spacing | Metres per cell, Y direction |
| `$AL$13` | QX | 315 | Grid size | Number of grid columns |
| `$AL$14` | QY | 317.46... | Grid size | Number of grid rows |
| `$AL$23` | index | 8 | — | (internal index count) |
| `$AL$24` | angle offset | 0 | Direction angles | Starting angle [degrees] |
| `$AL$25` | directions | 8 | Direction angles | Number of equally-spaced jet directions |
| `$AL$26` | angle directions | 0 | Direction angles | 0 = equal spacing; >0 = fixed step [deg] |

### Data Columns (scenario table, rows 2–120+)

| Cols | Header row value | Role | Description |
|------|-----------------|------|-------------|
| B | "Path" | Match key | Full ScenarioWeather key (e.g. `P7201/L4/…/Dia`) |
| C | "Scenario" | — | Scenario name without weather suffix |
| D | "Weather" | — | Weather condition (Dia / Noche) |
| E | "Flame length [m]" | — | Flame length from PHAST [m] |
| F–O (cols 6–15) | 1.6, 5, 7.3, 9.5, 12.5, 16, 20.9, 25, 30, 35 | `distV` | **Far-tip distance** from source to outer edge of kW/m² zone, in **cell units** |
| P–Y (cols 16–25) | 1.6, 5, 7.3, 9.5, 12.5, 16, 20.9, 25, 30, 35 | `halfWV` | **Semi-minor axis** (perpendicular half-width) of zone, in **cell units** |
| Z–AI (cols 26–35) | 1.6, 5, 7.3, 9.5, 12.5, 16, 20.9, 25, 30, 35 | `centerV` | **Center distance** from source to ellipse centre along jet axis, in **cell units** |

**Units note:** All three distance sets (distV, halfWV, centerV) are stored in **cell units**
(not metres). This is consistent with the formula computing X, Y in cell units (dividing
by SX/SY). The ellipse equation is dimensionally consistent only when all spatial quantities
share the same unit.

### Internal (LET) Variables

| Variable | Type | Size | Description |
|----------|------|------|-------------|
| `area` | Range | QY × QX | Spatial grid over ANCHOR sheet |
| `rId` | Scalar | 1 | Row index in ImpactJFMatrix matching `$AL$3` |
| `distV` | Array | 1 × 10 | Far-tip distances for 10 kW/m² thresholds |
| `halfWV` | Array | 1 × 10 | Semi-minor axes for 10 kW/m² thresholds |
| `centerV` | Array | 1 × 10 | Centre distances for 10 kW/m² thresholds |
| `impactV` | Array | 1 × 10 | kW/m² threshold values [1.6, 5, 7.3, …, 35] |
| `impactids` | Array | 1 × n_valid | Indices 1..n for valid (non-None) thresholds |
| `X` | Array | QY × QX | Cell X offsets from source (cell units, rounded integers) |
| `Y` | Array | QY × QX | Cell Y offsets from source (cell units, Y-up, rounded integers) |
| `dirs` | Array | n_dirs | Direction indices [0, 1, …, n_dirs-1] |
| `angles` | Array | n_dirs | Direction angles in degrees |
| `ctV` | Array | n_dirs | cos(angle) for each direction |
| `stV` | Array | n_dirs | sin(angle) for each direction |
| `eVals` | Array | n_dirs | Ellipse equation values for each direction (per threshold, per cell) |

---

## Key Difference from ImpactThermMatrix (old JF model)

| Aspect | ImpactThermMatrix (old JF) | ImpactJFMatrix (new JF) |
|--------|---------------------------|------------------------|
| Zone geometry | Circular (Euclidean distance) | Directional ellipses (N directions) |
| Distance data | 1 set of 10 distances (outer radius) | 3 sets of 10 (tip dist, half-width, center) |
| Output | Probability (via probit) | kW/m² thermal intensity level |
| Directions | Single circular zone | 8 equally-spaced jet directions |
| Parameters | Exposure time, probit constants | n_dirs, angle_offset, angle_step |

---

## Python Implementation Notes

In the Python engine, the formula is implemented as `formula_jf()`:

```python
x_rel = (XX - sx) / SX     # cell-unit X offset from source  (QY×QX)
y_rel = (YY - sy) / SY     # cell-unit Y offset from source  (QY×QX)

# For each of 8 directions:
angles = angle_offset + arange(n_dirs) * 360.0 / n_dirs
ct, st = cos(radians(angles)), sin(radians(angles))

# For each threshold id:
a = dist[id] - center[id]         # semi-major axis (cell units)
b = halfW[id]                     # semi-minor axis (cell units)
c = center[id]                    # center distance (cell units)

# Vectorised over grid and directions simultaneously:
proj_along = x_rel[...,None]*ct + y_rel[...,None]*st - c    # (QY,QX,n_dirs)
proj_perp  = x_rel[...,None]*st - y_rel[...,None]*ct        # (QY,QX,n_dirs)
eVals = (proj_along/a)**2 + (proj_perp/b)**2                # (QY,QX,n_dirs)

inside_any = any(eVals <= 1, axis=2)    # (QY,QX) bool
result = where(inside_any, maximum(result, imp_kw), result)
```

Final `result` is the maximum kW/m² threshold inside at least one ellipse at each cell.
The ROUND step from Excel is omitted for sub-cell accuracy.

---

## Direction Parameters (read from General sheet in qra_v6_engine.py)

| General sheet location | Variable | Default | Meaning |
|------------------------|----------|---------|---------|
| D22 | index | 8 | (internal sequence count) |
| D23 | angle offset | 0° | Starting angle of first jet direction |
| D24 | directions | 8 | Number of equally-spaced directions |
| D25 | angle directions | 0 | 0 = equal spacing; >0 = fixed angular step [deg] |

In `qra_engine.py` (KernelV0 engine) these are hardcoded constants matching the defaults.
