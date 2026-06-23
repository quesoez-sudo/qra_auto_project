# QRA Learning Guide — From Physics to Code

This document is a study companion to the QRA engine. It is designed to teach you **why** each formula exists — the engineering physics, the historical reasoning, and the decisions that shaped how we model risk. By the end you should be able to read any result the engine produces and understand what it really means about the world.

For the exact Excel and Python implementations see `formula_reference.md`. This guide focuses on understanding, not syntax.

---

## Part 1 — What is QRA and Why Does It Exist?

### The fundamental question

Every industrial plant with flammable or toxic materials poses a question to its neighbours and regulators: *How dangerous is this, exactly?* A Quantitative Risk Assessment (QRA) is the systematic engineering answer.

The goal is not to say "safe" or "unsafe" — it is to compute a number, the **individual risk** at each point in space, so that risk can be compared against tolerability criteria, optimised, and communicated.

### The risk equation

Risk is defined as:

```
Risk = Frequency × Probability of Harm
```

In this engine:

| Term | What it means | Units |
|------|--------------|-------|
| **Frequency** | How often does a specific event (fire, explosion, toxic release) occur? | events per year |
| **Probability of harm** | Given the event occurred, what is the chance a person at location X dies? | dimensionless (0 to 1) |
| **Individual Risk (IR)** | The annual probability that a specific person at a specific location dies | fatalities / (person · year) |

The engine computes IR at every cell of a 317×315 grid, producing a **risk map** of the facility footprint.

### What makes this "quantitative"

Before QRA, safety was assessed qualitatively: "this is high hazard, let's keep a 200m buffer". QRA replaces the buffer rule with a calculation. You can compare the result against criteria that regulators publish:

- **1×10⁻⁶ per year** — upper tolerable limit for the public (UK HSE, Dutch RIVM)
- **1×10⁻⁵ per year** — upper tolerable limit for workers
- **1×10⁻⁸ per year** — broadly acceptable (no further action needed)

A result of 3×10⁻⁶ tells you something specific: a person who lives at that location for a lifetime (~70 years) has roughly a 1-in-5000 extra chance of dying due to the facility. That is comparable to other societal risks and can be compared and justified.

### Events, scenarios, and sizes

A **scenario** in this engine is a specific combination of:
- Which equipment item is involved (e.g. `P7201AB`)
- What type of failure (e.g. `L7` = 7mm hole leak, `150mm` = 150mm bore failure, `SUC` = suction side)
- Which substance is released (e.g. `NAV_FL5` = naphtha flash 5, `HVN_FL50` = heavy naphtha flash 50)
- Day or night atmospheric conditions (`Dia` / `Noche`)

Each scenario is run through one or more **event types** (pool fire, flash fire, jet fire, toxic cloud) depending on the substance properties. The event type determines which physical model is applied and which formula computes lethality.

The engine runs 6 **size filters** (Total, S, M, L, XL, INST) to partition scenarios by hole diameter. This lets you understand what fraction of the risk comes from pinhole leaks vs catastrophic failures.

---

## Part 2 — The Grid and How Space Is Discretized

The risk map has **315 columns (west→east)** and **317 rows (north→south)**. Each cell is approximately 1.07 m × 1.07 m, making the total area roughly 337 m × 339 m.

### Cell-centre coordinates

To convert a grid cell (row `r`, column `c`) to real-world coordinates in metres:

```
x_centre = SX × (c − 0.5)        SX = 1.0698 m/cell
y_centre = SY × (QY − r + 0.5)   SY = 1.0694 m/cell, QY = 317
```

Row 1 is the **northern edge** of the map (highest y). Column 1 is the **western edge** (lowest x). This matches the way map data is typically stored (north-up raster).

### Distance from a source

For any source at real-world coordinates `(src_x, src_y)` in metres, the Euclidean distance to each cell centre is:

```
dist(c, r) = sqrt((x_centre − src_x)² + (y_centre − src_y)²)
```

This single precomputed 317×315 array of distances is the starting point for every event type except Jet Fire (which needs directional displacement).

---

## Part 3 — Thermal Radiation Events

### 3.1 The physical phenomenon

Pool fires, pressurised pool fires, and fireballs all kill the same way: **heat radiation**. A burning liquid or vapour cloud radiates energy outward. The amount of energy a person absorbs depends on:

1. **Intensity** `I` — the power per unit area of radiation arriving at the person's location, measured in kW/m²
2. **Exposure time** `t` — how long the person is exposed before the fire burns out or they can escape

Neither alone is sufficient: a brief exposure to very intense radiation and a long exposure to mild radiation can both be fatal. This led engineers to define a **thermal dose**.

### 3.2 The thermal dose — why the 4/3 power?

The thermal dose is:

```
D = I^(4/3) × t      [W/m² units, so I must be in W/m², not kW/m²]
```

The 4/3 exponent is not arbitrary. It comes from the physiology of skin burns. The criterion for a first-degree burn is when the skin surface reaches a critical temperature. Heat conduction into skin tissue follows a partial differential equation whose solution, integrated over time, shows that the **lethal dose scales as I^(4/3) × t**. Eberlin and later Eisenberg derived this from experiments on human skin and burn wound literature from the 1960s–1970s.

In practice, this means:
- Doubling the intensity for half the time gives **more** than equivalent harm (since 2^(4/3) ≈ 2.52, not 2)
- Very brief exposures to high intensity are disproportionately harmful compared to mild but prolonged exposures

### 3.3 The Eisenberg probit — from dose to probability

The dose `D` tells us how severe the exposure is. We then need to convert it to a **probability of death**. The tool for this is the **probit function**, invented by statistician Chester Bliss in 1934 for pesticide toxicity studies and later applied to industrial hazards.

A probit is a transformation of probability:

```
P(death) = Φ((Y − 5) / √2)
```

where Φ is the standard normal cumulative distribution function (CDF) and `Y` is the probit unit. The number 5 shifts the function so that Y=5 corresponds to P=50% (median lethal dose).

For thermal radiation, the probit was calibrated by Eisenberg et al. (1975) using data from nuclear weapon tests and industrial fire casualties. The result:

```
Y = −36.38 + 2.56 × ln(D)
  = −36.38 + 2.56 × ln((1000 × I)^(4/3) × t)
```

(The 1000 converts kW/m² to W/m², since the original calibration used SI units.)

**What the constants mean:**

| Constant | Meaning |
|----------|---------|
| −36.38 | Offset that places the 50% lethal point at around I = 10 kW/m² for t = 20 s |
| 2.56 | Slope: steepness of the dose-response curve. Higher values = sharper transition from "mostly survive" to "mostly die" |
| ln | Natural log, because the underlying distribution of individual susceptibility is log-normal |

### 3.4 Computing the result step by step

**Step 1** — The consequence modelling software (PHAST, SAFETI, or similar) runs a fire simulation and outputs 10 **isopleths**: radial distances from the source at which the intensity drops to each of the 10 threshold levels `[1.6, 5.0, 7.3, 9.5, 12.5, 16.0, 20.9, 25.0, 30.0, 35.0]` kW/m². These distances are stored in the `ImpactThermMatrix` sheet.

The isopleth geometry is circular because pool fires and fireballs radiate roughly symmetrically in all horizontal directions.

**Step 2** — For each grid cell, use the precomputed Euclidean distance `d` and **linearly interpolate** between the nearest two isopleths to find `I(d)`.

```
If d < d_innermost_ring  →  I = 35 kW/m² (maximum, certain death)
If d > d_outermost_ring  →  I = 0 (safe)
Otherwise:  I = I_inner + (I_outer − I_inner) / (d_outer − d_inner) × (d − d_inner)
```

**Step 3** — Apply the Eisenberg probit to convert `I` (kW/m²) and `t = 20 s` to lethality probability `P`.

**Step 4** — Multiply by frequency: `IR = f × P`.

### 3.5 Why 20 seconds for exposure time?

The 20-second exposure time is a convention from risk guidelines (e.g. UK HSE guidance). It represents:
- A person outdoors who is not immediately near an escape route
- Time to recognise the hazard and begin moving (reaction time ~5 s) + partial escape time

For people immediately adjacent to the fire source, 20 s is conservative. For people far enough away that the radiation is mild, the duration of the fire matters more, but the probability is already low.

---

## Part 4 — Flash Fire

### 4.1 The physical phenomenon

A **flash fire** is what happens when a large flammable vapour cloud ignites but the deflagration (flame front) is too slow to generate significant overpressure. Unlike an explosion, the primary hazard is **thermal exposure** as the flame front sweeps through the cloud. People inside the cloud die; people outside it survive.

Flash fires occur when:
1. A significant amount of flammable material is released and dispersed as a vapour cloud
2. The cloud reaches an ignition source
3. The cloud concentration at the ignition point is between the Lower Flammability Limit (LFL) and Upper Flammability Limit (UFL)

### 4.2 The two-zone model

The cloud is not uniformly lethal. Engineers divide the hazard into two concentric zones:

**Zone 1 — LFL zone (radius r_LFL):** The region where the gas-air mixture is above the LFL. Inside this zone:
- The mixture is flammable and will burn
- Any person engulfed in the flame is assumed to die with probability ≈ 1.0
- The LFL radius is computed by the dispersion model for the specific release scenario

**Zone 2 — LFLF zone (radius r_LFLF):** An outer zone representing edge effects — thermal radiation from the cloud boundary can still cause harm even outside the direct flame:
- This zone is retained in the impact matrix (value = 1) for consequence mapping
- In the risk calculation, only zone 1 (value = 2) contributes to individual risk

```
Impact(d) = 2   if d < r_LFL     → certain death zone
Impact(d) = 1   if d < r_LFLF    → edge exposure zone (mapped but not counted in IR)
Impact(d) = 0   otherwise         → safe
```

### 4.3 Why the binary model?

Real flash fires kill people inside the cloud with near-certainty. The physical time scale is fast (seconds), there is nowhere to go, and the temperature inside a combusting cloud is lethal. The binary model (in = dead, out = alive) is a well-validated simplification used across the industry. More complex models exist (e.g., probability varies with distance through the cloud) but the binary approach matches historical casualty data well for risk calculations.

**Individual risk:**

```
IR = f_FF × 1(d < r_LFL)
```

Only the people physically inside the LFL cloud at ignition are counted in the risk calculation.

---

## Part 5 — Toxic Dispersion

### 5.1 Why toxic events are different

For thermal events, we compute the hazard field ourselves (distance → kW/m² via isopleths → probit). For toxic events, the dose-response at each distance is already pre-computed and stored in the blob — we only need to **interpolate**.

This is because toxic dispersion depends on:
- Atmospheric stability (stable vs unstable atmosphere, day vs night)
- Wind speed (1 m/s vs 5 m/s changes the cloud shape dramatically)
- Substance properties (molecular weight, toxicity threshold)
- Source term (release rate, temperature, liquid fraction)

These are complex enough that a dedicated dispersion model (PHAST, SAFETI, ALOHA) is run separately and its output — a table of `(distance, lethality_probability)` pairs — is embedded in the workbook.

### 5.2 What the blob contains

Each scenario row in `ImpactToxMatrix` has a 5-column CSV blob in column G. The 5 columns are:

| Column | Name | Meaning |
|--------|------|---------|
| 1 | Distance | Radial distance from source in **metres** |
| 2 | Toxic dose | Time-integrated concentration (mg·min/m³ or similar substance-specific unit) |
| 3 | Probit number | Y value from the dose-response probit |
| 4 | Probability of fatality | P = Φ((Y−5)/√2), already computed |
| 5 | Integrated probability | Convolution of spatial probability over the cloud (for reporting) |

The blob rows are ordered **farthest to nearest** (decreasing distance, increasing probability).

We use only columns 1 and 4. The dispersion model has already done the dose-response conversion for us.

### 5.3 The dispersion model chain (what happens before the blob)

Even though we don't run this step, it is important to understand:

1. **Source term:** The release rate of the toxic substance (kg/s), its temperature and phase, the hole size and pressure
2. **Atmospheric stability class:** Pasquill-Gifford classes A (very unstable, sunny day) through F (very stable, calm night). The `Dia` (day) and `Noche` (night) scenario suffixes correspond to different stability classes — typically D for day, F for night
3. **Gaussian or heavy-gas dispersion model:** Computes the cloud concentration `C(x, y)` at each downwind distance
4. **Dose-response (toxic probit):** Substance-specific probit functions convert concentration × time to lethality probability. For example, for H₂S: `Y = −31.42 + 3.008 × ln(C² × t)`, for Cl₂: `Y = −8.29 + 0.92 × ln(C² × t)`. The constants differ by substance and come from animal and human toxicity data
5. **Resulting blob:** Distance vs P(death) table, embedded in the workbook

### 5.4 Why we filter rows with P < 1%

The kernel's `FILTER(table, col4 >= AA17)` step removes rows where the lethality probability is below 1% (AA17 = 0.01). This is done for two reasons:

1. **Numerics:** At large distances, probabilities become extremely small and are essentially zero for risk purposes. Including them would add floating-point noise without affecting any risk contour that matters
2. **Convention:** The 1% threshold is a common definition of the **toxic hazard zone boundary** in European QRA practice (Netherlands Purple Book, UK HSE guidance)

### 5.5 The interpolation

For a grid cell at distance `d` from the source:

```
If d > max_distance_in_table  →  P = 0     (beyond the 1% hazard zone)
If d < min_distance_in_table  →  P = 1.0   (so close it's certain death)
Otherwise: P = linear interpolation between the two surrounding table rows
```

Since the table is sorted farthest-to-nearest (descending), this is a **descending interpolation**. NumPy's `interp` requires ascending x — hence the `[::-1]` reversal in Python.

**Individual risk:** `IR = f_TOXIC × P(d)`

No additional probit is needed because `P` is already the lethality probability.

---

## Part 6 — Jet Fire

### 6.1 Why jet fires are different from pool fires

A pool fire sits on the ground and radiates roughly symmetrically. A **jet fire** is a high-pressure flame that projects in the direction of the release orifice. The fluid is moving at high velocity when it ignites, and momentum carries the flame outward as a roughly conical or ellipsoidal shape.

Key differences from a pool fire:
- The flame is **directional**, not circular
- The reach is **much longer** relative to the release rate (a 10-bar gas leak produces a jet flame 50 m long, not just a pool at the base)
- The thermal radiation field is **elongated** in the direction of the jet

### 6.2 The ellipse model for radiation isopleths

Consequence models (like PHAST Jet fire module) compute the flame shape and from it derive radiation isopleths. Instead of circles, these isopleths are **ellipses** whose major axis aligns with the wind/jet direction.

Each kW/m² threshold level `id` (10 levels from 1.6 to 35 kW/m²) has its own ellipse, characterised by three numbers stored in `ImpactJFMatrix`:

| Parameter | Symbol | Meaning |
|-----------|--------|---------|
| Forward tip distance | `dist_v[id]` | How far the flame ellipse extends **ahead** of the release point (in cell units) |
| Lateral half-width | `half_v[id]` | How wide the ellipse is at its widest point, perpendicular to the jet direction (cell units) |
| Centre offset | `cent_v[id]` | Distance from the release point to the **geometric centre** of the ellipse |

The semi-axes of the ellipse are then:
```
arm (forward) = dist_v[id] − cent_v[id]   (half the length)
hw  (lateral) = half_v[id]                (half the width)
```

### 6.3 The rotated ellipse equation

A grid cell at position `(x, y)` in cell units relative to the source is inside the kW/m² ellipse for direction θ if:

```
[(dx·cosθ + dy·sinθ)² / arm²] + [(dx·sinθ − dy·cosθ)² / hw²] ≤ 1
```

where:
```
dx = x_cell − cent_v[id] × cosθ    (displacement from ellipse centre, not from source)
dy = y_cell − cent_v[id] × sinθ
```

This is the standard equation of a rotated ellipse in matrix form. The rotation matrix `[[cosθ, sinθ], [−sinθ, cosθ]]` transforms the cell coordinates into the ellipse's own reference frame (forward/lateral), and then the ellipse test is simply the sum of normalised squared displacements.

**Intuition:** Imagine the ellipse as a rugby ball. The long axis points in direction θ. We rotate the coordinate system to align with the ball, then check whether the point is inside using the standard ellipse inequality.

### 6.4 Handling uncertain wind direction — the "flower of ignition"

The direction of a jet fire at the moment of ignition is uncertain. Wind direction changes over time, and we don't know which direction it will be when the fire occurs. QRA handles this by considering **8 equally spaced wind directions** (every 45°: 0°, 45°, 90°, 135°, 180°, 225°, 270°, 315°) and assigning each equal probability.

For each grid cell, the engine computes the **maximum kW/m² level** across all 8 directions:

```
Impact(cell) = max over all θ, max over all id: IF(cell inside ellipse(id, θ)) → kW[id]
```

When you visualise the result, you see a **"flower" pattern** — 8 overlapping ellipses radiating outward from the source, one for each wind direction. The petals of the flower cover a ring around the source. Very close cells (inside the smallest ellipse in all directions) have the maximum kW/m². Cells at the tips of the petals only receive radiation when the wind points exactly toward them.

### 6.5 Risk calculation — same probit as thermal

Once we have the kW/m² value at each cell (the maximum over all jet directions), the risk calculation is **identical to pool fire thermal**:

```
Y      = −36.38 + 2.56 × ln((1000 × I)^(4/3) × t)
P      = Φ((Y−5)/√2)
IR     = f_JF × P
```

The dose-response is the same because a person standing in the radiation field of a jet fire experiences the same physics (thermal flux) as one near a pool fire. The only difference is the geometry — how the intensity `I` varies with position.

---

## Part 7 — Accumulation: How Total Risk Is Built

### 7.1 One scenario, one matrix

For a single scenario (one release event on one piece of equipment), the engine produces a 317×315 **risk matrix** — each cell contains the annual probability of fatality assuming this specific scenario occurs at the specified rate.

The matrix for scenario `s` is:
```
Risk_s(cell) = f_s × P_lethal_s(cell)
```

### 7.2 Summing across scenarios

Total Individual Risk at each cell is the **sum** across all contributing scenarios:

```
IR_total(cell) = Σ_s  Risk_s(cell)
               = Σ_s  f_s × P_lethal_s(cell)
```

This is valid under the assumption that events are independent and rare (which they are, since each f_s is on the order of 10⁻⁵ to 10⁻⁷ per year — multiple simultaneous fires are negligible in probability).

### 7.3 The size filters — S, M, L, XL, INST, Total

Each piece of equipment has multiple hole-size scenarios. The engine runs the accumulation 6 times:

| Filter name | Scenario selection | Purpose |
|-------------|-------------------|---------|
| **Total** | All scenarios | Overall combined risk |
| **S** | Small holes (pin-holes, ~1mm) | Risk from minor leaks |
| **M** | Medium holes (~5mm) | |
| **L** | Large holes (~25mm) | |
| **XL** | Very large holes (>50mm) | |
| **INST** | Instantaneous / catastrophic releases | Full pipe bore or vessel rupture |

Comparing Total vs INST tells you whether the facility's risk is dominated by rare large events or frequent small ones. This distinction matters for mitigation: improving maintenance reduces S/M risk, while structural integrity management targets INST.

### 7.4 Day and night split

Many scenarios appear in both `Dia` (day) and `Noche` (night) variants. This captures:
- Different atmospheric stability (day = more turbulent, faster dispersion; night = stable, slower dispersion and denser cloud)
- Different number of exposed persons (operational staff changes)
- Different background ignition probability (certain ignition sources only operate during day)

The day/night frequencies are split from the overall scenario frequency, typically 60%/40% or according to operational data.

---

## Part 8 — Reading the Results

### 8.1 What the numbers mean

A risk matrix cell value of `5×10⁻⁶` means: "If a person lived at this exact location for 70 years, they would have an additional 0.035% chance of dying due to a hazardous event at this facility, over and above all other causes of death."

For context: in most countries, the background risk of dying in a given year from all causes combined is roughly `10⁻²` (1%). Industrial risk criteria are set **four to six orders of magnitude below** this, because the risk is involuntary and specifically imposed by the facility.

### 8.2 Typical values from this engine

From the test runs on this facility:

| Event type | Non-zero region size | Typical IR at source vicinity |
|------------|---------------------|------------------------------|
| Thermal (pool fire) | Circular, ~50–300m radius | ~10⁻⁵ to 10⁻⁷ |
| Flash fire | Circular, ~50–500m radius | ~10⁻⁵ to 10⁻⁶ |
| Jet fire | Elliptical "flower", ~50–300m reach | ~10⁻⁵ to 10⁻⁷ |
| Toxic | Small spot or large cloud (depends on substance) | ~10⁻⁶ to 10⁻⁴ at source proximity |

Note that some toxic scenarios (large releases, `150mm` bore, night conditions) produce non-zero values across the entire 337m×337m grid — the cloud reaches everywhere. This is physically real: a 150mm pipe rupture of a highly toxic gas in stable night conditions can produce lethal concentrations hundreds of metres downwind.

### 8.3 The event frequency hierarchy

Typical event frequencies from loss-of-containment databases (EGIG, TNO Purple Book):

| Leak size | Typical frequency |
|-----------|-------------------|
| Pin-hole (< 2mm) | ~10⁻³ / year / km pipe |
| Small hole (2–10mm) | ~5×10⁻⁵ / year |
| Large hole (> 50mm) | ~5×10⁻⁶ / year |
| Catastrophic rupture | ~10⁻⁷ to 10⁻⁶ / year |

The full QRA frequency for any specific event (`P_LPF`, `P_JF`, etc.) in the Core sheet is the product of: leak frequency × conditional probability of ignition × conditional probability of that specific fire type (jet vs pool) given ignition.

---

## Part 9 — A Worked Example

Let's trace the calculation for **one cell** to make it concrete. Take scenario `P7201AB/L7/NAV_FL5/H/1mDia` (LPF thermal event, small hole on pump P7201AB, naphtha fuel, day conditions).

Suppose the source is at `src_x = 150.0 m, src_y = 200.0 m` and the cell of interest is at grid column 150, row 200.

**Step 1 — Cell position in metres:**
```
x = 1.0698 × (150 − 0.5) = 1.0698 × 149.5 = 159.94 m
y = 1.0694 × (317 − 200 + 0.5) = 1.0694 × 117.5 = 125.65 m
```

**Step 2 — Distance from source:**
```
dist = sqrt((159.94 − 150.0)² + (125.65 − 200.0)²)
     = sqrt(9.94² + 74.35²)
     = sqrt(98.8 + 5527.9)
     = sqrt(5626.7)
     = 75.0 m
```

**Step 3 — kW/m² by interpolation:**

Suppose the thermal isopleth distances from the scenario row are:
```
 35 kW/m²  at  d = 10 m
 25 kW/m²  at  d = 20 m
 16 kW/m²  at  d = 30 m
  9.5 kW/m²  at  d = 45 m
  5.0 kW/m²  at  d = 65 m
  1.6 kW/m²  at  d = 90 m  (outermost ring)
```

Our cell is at d = 75 m, which falls between the 5.0 kW/m² ring (65 m) and the 1.6 kW/m² ring (90 m):
```
I = 5.0 + (1.6 − 5.0) / (90 − 65) × (75 − 65)
  = 5.0 + (−3.4 / 25) × 10
  = 5.0 − 1.36
  = 3.64 kW/m²
```

**Step 4 — Eisenberg probit:**
```
Y = −36.38 + 2.56 × ln((1000 × 3.64)^(4/3) × 20)
  = −36.38 + 2.56 × ln(3640^1.333 × 20)
  = −36.38 + 2.56 × ln(258320 × 20)     [3640^1.333 ≈ 258320]
  = −36.38 + 2.56 × ln(5166400)
  = −36.38 + 2.56 × 15.46
  = −36.38 + 39.58
  = 3.20
```

**Step 5 — Lethality probability:**
```
P = Φ((3.20 − 5) / √2) = Φ(−1.273) = 0.102   (about 10.2% chance of death)
```

**Step 6 — Individual risk:**
```
f_LPF = 2 × 10⁻⁵ per year  (example frequency)
IR = 2 × 10⁻⁵ × 0.102 = 2.04 × 10⁻⁶ per year
```

A person living at that cell would have an additional 2×10⁻⁶ annual probability of death from this single scenario. After summing contributions from all thermal, flash fire, toxic, and jet fire scenarios, the total IR at that cell might be, say, 8×10⁻⁶ per year — just above the typical tolerability threshold.

---

## Part 10 — Why This Tool Matters

### The connection to decisions

The risk matrices produced by this engine are not just numbers — they are the basis for:

1. **Land-use planning:** Setting exclusion zones around the facility boundary. No sensitive land use (houses, schools, hospitals) inside the 10⁻⁶ per year contour
2. **Layout decisions:** Comparing two equipment layouts to see which produces lower off-site risk
3. **Mitigation cost-benefit:** If installing better isolation valves costs €100k and reduces IR by 2×10⁻⁷, is that money well spent? QRA makes this comparison possible
4. **Regulatory compliance:** Demonstrating to authorities that risk is ALARP (As Low As Reasonably Practicable)

### What the engine does not model

Understanding limitations is as important as understanding capabilities:

| Not modelled | Why it matters |
|-------------|---------------|
| BLEVE (Boiling Liquid Expanding Vapour Explosion) | No BLEVE scenarios in current `ImpactExpMatrix` |
| Late explosion | Vapour cloud drift and delayed ignition — not yet implemented |
| Cascade effects | A fire that causes a secondary vessel failure |
| Occupied building shielding | A person inside a building has different exposure than one outdoors |
| Population density | IR is per-person; you need population data to compute societal risk (FN curves) |

---

## Quick Reference Card

```
EVENT TYPE     HAZARD MEASURE   DOSE-RESPONSE      GEOMETRY
──────────────────────────────────────────────────────────────────
Pool Fire      I (kW/m²)        Eisenberg probit   Circular isopleths
Flash Fire     Zone (0/1/2)     Binary (zone 2=1)  Two concentric circles
Toxic          P_lethal (0-1)   Pre-embedded       Radial from source
Jet Fire       I (kW/m²)        Eisenberg probit   8 rotated ellipses

PROBIT:  Y = −36.38 + 2.56 × ln((1000·I)^(4/3) × t)
CDF:     P = 0.5 × (1 + erf((Y−5)/√2))
RISK:    IR = frequency × P    [events/year → fatalities/person/year]

GRID:  315 cols × 317 rows  |  SX = 1.0698 m/cell  |  SY = 1.0694 m/cell
COORD: x = SX×(col−0.5)    |  y = SY×(QY−row+0.5) |  row 1 = north
```

---

*This guide covers the specific models implemented in `qra_engine_v2.py`. For the exact formulas and parameter tables see `formula_reference.md`. For model limitations and open items see `docs/open_questions.md`.*
