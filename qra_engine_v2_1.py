"""
QRA Engine v2.1
===============
Reads directly from the production MacroQRAV6 workbook — General (grid/thresholds),
Core (scenarios, X/Y, frequencies, per-scenario JF/exposure params), and the actual
Results sheets (not the kernel's ImpactXxxMatrix test sheets). Branched from
qra_engine_v2.py, which still reads kernel V0-4 + KernelV0_v2_copy.xlsx/CoreExample —
see that file for the kernel-pinned version.

Covered event types (see EVENT_CONFIG for the sheet/frequency mapping):
  Thermal    — Early Pool Fire / Late Pool Fire / Fireball Results — circular distance, thermal probit
  Flash Fire — Flash Fire Results    — step function on LFL / LFLF radius
  Toxic      — Outdoor Toxic Results — CSV-blob probability interpolation
  Jet Fire   — Jet Fire Results      — ELLIPTICAL zones, per-scenario direction count
  Explosion/CVE — CVE Results        — step function, ignition-point coords (Xign/Yign)

Known gap: BLEVE Results has no reader yet (own column layout, psi units, uses Core
X/Y not Xign/Yign — can't reuse compute_exp()). Also: Core is read once per run via
openpyxl (a static snapshot) — qra_v6_engine.py sets Core!C2 to each Impact ID and
forces a live recalculation via xlwings before reading, since Core's scenario set is
formula-driven per Impact ID. This engine does not yet do that — see parameters_joint.md.

Output: output_v2.1/<event_type>/<scenario>_<weather>.csv
        output_v2.1/<event_type>_summary.csv   (key stats per scenario)
"""

import math
import csv
import re
import time
import numpy as np
from pathlib import Path
from scipy.special import erf
from contextlib import contextmanager
import openpyxl
import xlwings as xw
import ctypes

# ── Config ────────────────────────────────────────────────────────────────────
# Everything now comes from the same production workbook — General/Core/Results sheets
# are all read from this one file (kernel V0-4 is no longer used by v2.1).
INPUT_FILE    = "MacroQRAV6 (version 1)-Win11-cWvAMQq4D.xlsm"   # Results sheets (see EVENT_CONFIG)
CORE_FILE     = INPUT_FILE                                      # Core sheet: X/Y, per-scenario JF/thermal params, frequencies
CORE_SHEET    = "Core"
GENERAL_FILE  = INPUT_FILE                                      # General sheet: QX/QY/SX/SY + per-event levels
GENERAL_SHEET = "General"
OUTPUT_DIR = Path("output_v2.1")

# Kernel ImpactXxxMatrix -> production Results-sheet equivalence. One kernel "formula
# family" can map to several production events — e.g. thermal covers three separate
# sheets (Early Pool Fire / Late Pool Fire / Fireball), each its own sheet + frequency
# column. impact_id matches qra_v6_engine.py's IMPACT_CONFIG / PageControl!L exactly —
# it's what gets written to Core!C2 to select that event's live scenario set.
#   ImpactThermMatrix -> Early/Late Pool Fire Results, Fireball Results
#   ImpactToxMatrix   -> Outdoor Toxic Results
#   ImpactFFMatrix    -> Flash Fire Results
#   ImpactExpMatrix   -> CVE Results          (BLEVE Results has no reader yet — see below)
#   ImpactJFMatrix    -> Jet Fire Results
EVENT_CONFIG = [
    {'impact_id': 16, 'label': 'TOXIC', 'formula': 'toxic',   'sheet': 'Outdoor Toxic Results',   'freq_col': 'P_TOXIC'},
    {'impact_id': 17, 'label': 'JF',    'formula': 'jf',      'sheet': 'Jet Fire Results',        'freq_col': 'P_JF'},
    {'impact_id': 18, 'label': 'LPF',   'formula': 'thermal', 'sheet': 'Late Pool Fire Results',  'freq_col': 'P_LPF'},
    {'impact_id': 19, 'label': 'EPF',   'formula': 'thermal', 'sheet': 'Early Pool Fire Results', 'freq_col': 'P_EPF'},
    {'impact_id': 20, 'label': 'FB',    'formula': 'thermal', 'sheet': 'Fireball Results',        'freq_col': 'P_FB'},
    {'impact_id': 21, 'label': 'CVE',   'formula': 'exp',     'sheet': 'CVE Results',             'freq_col': 'P_CVE'},
    {'impact_id': 22, 'label': 'BLV',   'formula': None,      'sheet': 'BLEVE Results',           'freq_col': 'P_BLV'},  # no reader yet
    {'impact_id': 23, 'label': 'FF',    'formula': 'ff',      'sheet': 'Flash Fire Results',       'freq_col': 'P_FF'},
]

SIZES = ['Total', 'S', 'M', 'L', 'XL', 'INST']
_SIZE_NORM = {sz.upper(): sz for sz in SIZES if sz != 'Total'}   # case-insensitive lookup

SHEET_IMPACT = 'CalcBUImpact'
SHEET_RISK   = 'CalcBURisk'
EXPORT_FILE  = "MacroQRAV6_v2.1_export_result.xlsx"   # never overwrites the original .xlsm

# Everything below is a fallback default only — load_general_params() overwrites these
# module globals at the top of main() with values read from the General sheet. They're
# kept as literals so compute_*() functions still work if called standalone/in tests.
QX = 315
QY = 317
SX = 1.0698412698412698   # m / cell
SY = 1.069425             # m / cell

# kW/m² threshold labels — overwritten from General!J17:S17 ("Thermal Levels") at runtime
KW_THRESHOLDS = np.array([1.6, 5.0, 7.3, 9.5, 12.5, 16.0, 20.9, 25.0, 30.0, 35.0])

# FF zone values [outside, transition, inside] — overwritten from General!J18:L18
# ("Flammable concentration levels") at runtime
FF_LEVELS = [0.0, 1.0, 2.0]

# CVE risk step-function values [limit1, limit2, prob_mid, prob_high, ...] — overwritten
# from General!J19:N19 ("Overpressure levels") at runtime. NOT the same thing as the
# per-scenario bar thresholds in ImpactExpMatrix/CVE Results row 1 — see compute_exp().
EXP_THRESHOLDS = [0.04, 0.1, 0.35, 0.5, 1.0]

# Toxic minimum-probability filter — overwritten from General!J20 ("Toxic Levels", first
# value) at runtime
TOX_MIN_PROB = 0.01

# JF direction parameters are PER-SCENARIO in Core!H (directions) / I (offset) / J (step);
# these are only the fallback defaults used when a scenario's Core cell is empty.
# JF_N_DIRS_DEFAULT is overwritten from General!D21 ("Discharge directions (Default)") —
# there's no General cell for angle offset/step, so 0.0 (uniform, no rotation) is used.
JF_N_DIRS_DEFAULT       = 8
JF_ANGLE_OFFSET_DEFAULT = 0.0   # degrees
JF_ANGLE_STEP_DEFAULT   = 0.0   # degrees — 0 = uniform 360/N_dirs

# Thermal/JF exposure time is PER-SCENARIO in Core!K ("Exposure Time"); this is the
# fallback used when empty — overwritten from General!D22 ("Exposure duration") at runtime.
THERM_EXP_TIME_DEFAULT = 20.0

# Toxic unit mode:
#   'meters' — blob col1 is in metres (confirmed: kernel distM is in m, blob col1 is same unit)
#   'km'     — divide distM by 1000 before matching (incorrect — kept for reference only)
TOXIC_UNIT_MODE = 'meters'

# ── Warning / error collector ────────────────────────────────────────────────
# The goal is for this engine to eventually run from Excel and write matrices back
# into it — so warnings can't just scroll past in a console. Every _warn() call is
# both printed immediately (for live console runs) and collected, so main() can print
# one consolidated summary at the very end regardless of how it's invoked.
_warnings = []

def _warn(msg):
    _warnings.append(msg)
    print(f"  [WARN] {msg}")


# ── UI freeze (button path only) ─────────────────────────────────────────────
@contextmanager
def _frozen(app):
    """Freeze screen updates / calculation / events on `app` for the duration of
    the with-block; restores the ORIGINAL values (not hardcoded defaults) in a
    finally, so a caller that already had non-default settings isn't clobbered."""
    screen_updating = app.screen_updating
    calculation = app.calculation
    enable_events = app.api.EnableEvents
    app.screen_updating = False
    app.calculation = 'manual'
    app.api.EnableEvents = False
    try:
        yield
    finally:
        app.screen_updating = screen_updating
        app.calculation = calculation
        app.api.EnableEvents = enable_events


def _progress(app, text):
    """Application.StatusBar progress text — called once per event during
    _run_pipeline(). Cleared with app.api.StatusBar = False by the caller once
    the pipeline finishes (success or failure)."""
    app.api.StatusBar = text
    print(f"  {text}")


def _format_warning_summary(warnings_list):
    """Pure formatting — no Excel/Windows dependency, easy to unit test."""
    if not warnings_list:
        return "Sin advertencias."
    lines = [f"{i}. {w}" for i, w in enumerate(warnings_list, 1)]
    return f"{len(warnings_list)} advertencia(s):\n" + "\n".join(lines)


def _show_msgbox(text, title="QRA Engine v2.1"):
    """Native Windows message box — VBA's MsgBox isn't COM-exposed, so this is
    called directly from Python rather than routed back through VBA."""
    MB_ICONINFORMATION = 0x40
    ctypes.windll.user32.MessageBoxW(0, text, title, MB_ICONINFORMATION)


# ── Pipeline time report ──────────────────────────────────────────────────────
_timings = {}


@contextmanager
def _timed(key):
    """Accumulates wall-clock seconds spent inside the with-block into
    _timings[key]. Used to time-bucket _run_pipeline() into its four stages:
    reading, calculation, iteration, writing. Still records the elapsed time
    if the block raises (finally), so a failed run still reports where the
    time went up to the point of failure."""
    t0 = time.perf_counter()
    try:
        yield
    finally:
        _timings[key] = _timings.get(key, 0.0) + (time.perf_counter() - t0)


def _format_timing_report(timings):
    """Pure formatting — no Excel dependency, easy to unit test."""
    order = ['reading', 'calculation', 'iteration', 'writing']
    total = sum(timings.get(k, 0.0) for k in order)
    lines = [f"  {k.capitalize():<12}: {timings.get(k, 0.0):6.1f} s" for k in order]
    return "Tiempos de ejecución:\n" + "\n".join(lines) + f"\n  {'Total':<12}: {total:6.1f} s"


# ── Grid arrays ───────────────────────────────────────────────────────────────
# Pre-compute once; all formula functions take (dist_m, x_c, y_c) as arguments.

def _build_grids(src_x=None, src_y=None, qx=None, qy=None, sx=None, sy=None):
    """
    Returns:
        dist_m  (QY, QX)  Euclidean distance from source, metres
        x_c     (1,  QX)  x displacement in cell-width units from source  (for JF)
        y_c     (QY,  1)  y displacement in cell-height units from source (for JF)

    qx/qy/sx/sy default to the current QX/QY/SX/SY module globals — resolved at call
    time (not definition time) so main()'s dynamic General-sheet values apply.
    """
    qx = QX if qx is None else qx
    qy = QY if qy is None else qy
    sx = SX if sx is None else sx
    sy = SY if sy is None else sy

    cols = np.arange(1, qx + 1)
    rows = np.arange(1, qy + 1)

    # Cell-centre positions in metres
    x_m = sx * (cols - 0.5)                    # (QX,)
    y_m = sy * (qy - rows + 0.5)               # (QY,)  row 1 = northmost

    dist_m = np.sqrt(
        (x_m[None, :] - src_x) ** 2 +
        (y_m[:, None] - src_y) ** 2
    )                                           # (QY, QX)

    # Cell units from source — used by the JF ellipse formula
    src_col = src_x / sx
    src_row = src_y / sy
    x_c = ((cols - 0.5) - src_col)[None, :]   # (1, QX)
    y_c = ((qy - rows + 0.5) - src_row)[:, None]  # (QY, 1)

    return dist_m, x_c, y_c


# ── Thermal probit ────────────────────────────────────────────────────────────
def _thermal_prob(kw_mat, exp_time):
    """kW/m² array -> lethality probability array (same shape)."""
    p = np.zeros_like(kw_mat, dtype=float)
    m = (kw_mat > 0) & (kw_mat < 35.0)
    Y = -36.38 + 2.56 * np.log((1000.0 * kw_mat[m]) ** (4.0 / 3.0) * exp_time)
    p[m] = 0.5 * (1.0 + erf((Y - 5.0) / math.sqrt(2.0)))
    p[kw_mat >= 35.0] = 1.0
    return np.clip(p, 0.0, 1.0)


# ── THERMAL ───────────────────────────────────────────────────────────────────
def compute_thermal(dist_m, distances, exp_time=20.0, return_risk=False, frequency=1.0):
    """
    dist_m    : (QY, QX) distances in metres
    distances : array-like length 10 — each entry is the radial distance (m) at
                which kW/m² drops to KW_THRESHOLDS[i]. NaN / '' = threshold not reached.
    exp_time  : exposure seconds (from $AA$8)
    Returns   : kW/m² impact matrix (QY, QX), or risk matrix if return_risk=True
    """
    dv = np.array([float(x) if x not in ('', None) else np.nan for x in distances])
    valid = ~np.isnan(dv)
    dv_v = dv[valid]
    kw_v = KW_THRESHOLDS[valid]

    impact = np.zeros_like(dist_m, dtype=float)
    if len(dv_v) == 0:
        return impact

    # Inside every threshold -> max kW
    impact[dist_m <= dv_v[-1]] = kw_v[-1]

    # Linear interpolation between consecutive threshold distances
    for i in range(len(dv_v) - 1):
        mask = (dist_m > dv_v[i + 1]) & (dist_m <= dv_v[i])
        slope = (kw_v[i + 1] - kw_v[i]) / (dv_v[i + 1] - dv_v[i])
        impact[mask] = kw_v[i] + slope * (dist_m[mask] - dv_v[i])

    # Beyond outermost threshold -> 0 (already zero)

    if return_risk:
        return frequency * _thermal_prob(impact, exp_time)
    return impact


# ── FLASH FIRE ────────────────────────────────────────────────────────────────
def compute_ff(dist_m, lfl_r, lflf_r, return_risk=False, frequency=1.0):
    """
    lfl_r  : LFL radius (m)   — col E
    lflf_r : LFLF radius (m)  — col F
    Impact: FF_LEVELS[2] inside LFL, [1] between LFL-LFLF, [0] outside
            (values read dynamically from General!J18:L18 — usually [0, 1, 2])
    Risk  : frequency x (1 if inside LFL, else 0)
    """
    outside, transition, inside = FF_LEVELS[0], FF_LEVELS[1], FF_LEVELS[2]
    impact = np.full_like(dist_m, outside, dtype=float)
    impact[dist_m < lflf_r] = transition
    impact[dist_m < lfl_r]  = inside
    if return_risk:
        return frequency * (impact == inside).astype(float)
    return impact


# ── TOXIC ─────────────────────────────────────────────────────────────────────
def _parse_tox_blob(blob_str, min_prob=None):
    """
    Parse the 5-column CSV blob from ImpactToxMatrix col G.
    min_prob defaults to the TOX_MIN_PROB module global (General!J20) when not given.

    Columns in the blob:
      col1: distance-like value (hypothesis: km) — descending, negative = upwind
      col2: concentration or dispersion parameter
      col3: log-like value
      col4: lethality probability (0-1)
      col5: secondary metric

    Filter: col1 >= 0  AND  col4 >= min_prob
    Returns (j_col, m_col) — sorted descending by col1 (distance).
    """
    if min_prob is None:
        min_prob = TOX_MIN_PROB

    rows = []
    for line in blob_str.strip().split('\n'):
        parts = line.split(',')
        if len(parts) < 5:
            continue
        try:
            vals = [float(p.strip()) for p in parts]
            rows.append(vals)
        except ValueError:
            continue

    if not rows:
        return np.array([]), np.array([])

    arr = np.array(rows)                         # (N, 5)
    keep = (arr[:, 0] >= 0) & (arr[:, 3] >= min_prob)
    filtered = arr[keep]

    if len(filtered) == 0:
        return np.array([]), np.array([])

    j_col = filtered[:, 0]   # col1 — distance-like, descending
    m_col = filtered[:, 3]   # col4 — lethality probability

    return j_col, m_col


def compute_toxic(dist_m, blob_str, return_risk=False, frequency=1.0,
                  unit_mode=TOXIC_UNIT_MODE):
    """
    dist_m   : (QY, QX) distances in metres
    blob_str : raw CSV blob from col G
    unit_mode: 'meters' -> compare distM directly to blob col1
               'km'     -> divide distM by 1000 before comparing (col1 treated as km)

    NOTE: The blob's col1 unit is ambiguous in this standalone kernel file.
    With unit_mode='km' the formula gives physically meaningful output when
    toxic clouds extend to hundreds of metres. Verify with real Core coords.
    """
    j_col, m_col = _parse_tox_blob(blob_str)
    if len(j_col) == 0:
        return np.zeros_like(dist_m, dtype=float)

    d = dist_m / 1000.0 if unit_mode == 'km' else dist_m

    j_max = j_col[0]    # largest (= farthest)
    j_min = j_col[-1]   # smallest positive

    # np.interp needs ascending x-coords; flip the descending j/m arrays
    j_asc = j_col[::-1]
    m_asc = m_col[::-1]

    out = np.interp(d, j_asc, m_asc, left=1.0, right=0.0)
    out = np.clip(out, 0.0, 1.0)

    if return_risk:
        return frequency * out
    return out


# ── JET FIRE ──────────────────────────────────────────────────────────────────
def compute_jf(x_c, y_c,
               dist_v, half_v, cent_v,
               n_dirs=JF_N_DIRS_DEFAULT,
               angle_offset=JF_ANGLE_OFFSET_DEFAULT,
               angle_step=JF_ANGLE_STEP_DEFAULT,
               exp_time=20.0,
               return_risk=False,
               frequency=1.0):
    """
    n_dirs/angle_offset/angle_step/exp_time are PER-SCENARIO in production (Core!H/I/J/K) —
    main() always passes them explicitly per scenario; the defaults here only apply to
    standalone/test calls.
    Jet-fire elliptical impact formula.

    x_c, y_c : cell-unit displacement from source, shapes (1, QX) and (QY, 1)
    dist_v    : forward tip distances in cell units  (10 values, NaN if unused)
    half_v    : lateral half-width distances in cell units
    cent_v    : centre-from-source distances in cell units

    Algorithm (mirrors ImpactJFMatrix AL1 array formula):
      For each kW/m² threshold id (with valid data):
        ellipse centre  = c · (cos θ, sin θ)      [cell units]
        forward semi-ax = distV[id] - centerV[id]
        lateral semi-ax = halfWV[id]
        ellipse test    = (dx·cosθ + dy·sinθ)² / (d-c)² +
                          (dx·sinθ - dy·cosθ)² / halfW²  ≤ 1
        where dx = x - c·cosθ,  dy = y - c·sinθ
      impact[cell] = MAX(kW) over all (id, direction) where cell is inside ellipse
      Final: zero out cells beyond distV[0] (max forward reach) from source.

    Returns kW/m² impact matrix (QY, QX), or risk matrix if return_risk=True.
    """
    dv = np.array([float(x) if x not in ('', None) else np.nan for x in dist_v])
    hv = np.array([float(x) if x not in ('', None) else np.nan for x in half_v])
    cv = np.array([float(x) if x not in ('', None) else np.nan for x in cent_v])

    valid = ~(np.isnan(dv) | np.isnan(hv) | np.isnan(cv) |
              (dv == 0) | (hv == 0))
    dv_v = dv[valid]
    hv_v = hv[valid]
    cv_v = cv[valid]
    kw_v = KW_THRESHOLDS[valid]

    qy, qx = x_c.shape[1], y_c.shape[0]   # note: x_c (1,QX), y_c (QY,1)
    qy, qx = y_c.shape[0], x_c.shape[1]

    impact = np.zeros((qy, qx), dtype=float)

    if len(dv_v) == 0:
        return impact

    # Distance from source in cell units (for dmaxim outer mask)
    dist_cells = np.sqrt(x_c ** 2 + y_c ** 2)   # broadcast -> (QY, QX)
    d_maxim = dv_v[0]

    # Build angle list
    dirs = np.arange(n_dirs)
    if angle_step == 0:
        angles_deg = angle_offset + dirs / n_dirs * 360.0
    else:
        angles_deg = angle_offset + dirs * angle_step
    angles_rad = np.radians(angles_deg)

    for id_idx in range(len(kw_v)):
        d  = dv_v[id_idx]   # forward tip
        hw = hv_v[id_idx]   # lateral half-width
        c  = cv_v[id_idx]   # centre distance from source
        kw = kw_v[id_idx]

        arm = d - c          # forward semi-axis of ellipse
        if arm <= 0 or hw <= 0:
            continue

        a = 1.0 / arm ** 2   # aPower
        b = 1.0 / hw  ** 2   # bPower

        inside_any = np.zeros((qy, qx), dtype=bool)

        for theta in angles_rad:
            ct = math.cos(theta)
            st = math.sin(theta)

            # Vector from ellipse centre to each cell, in cell units
            dx = x_c - c * ct   # (1, QX) broadcasts with y_c (QY, 1) -> (QY, QX)
            dy = y_c - c * st

            # Ellipse quadratic form (= 1 on ellipse boundary)
            evals = (dx * ct + dy * st) ** 2 * a + \
                    (dx * st - dy * ct) ** 2 * b

            inside_any |= (evals <= 1.0)

        impact = np.maximum(impact, inside_any * kw)

    # Cells beyond max forward reach -> 0
    impact[dist_cells > d_maxim] = 0.0

    if return_risk:
        return frequency * _thermal_prob(impact, exp_time)
    return impact


# ── EXPLOSION (CVE) ───────────────────────────────────────────────────────────
def compute_exp(dist_m, distances, impact_v, return_risk=False, frequency=1.0):
    """
    Explosion (CVE) impact/risk formula (mirrors ImpactExpMatrix AB1).

    dist_m    : (QY, QX) distances in metres from ignition point (Xign, Yign)
    distances : 5 values (cols J-N per scenario row) — ascending distances (m) at
                which overpressure equals the corresponding impact_v thresholds.
                Zero or NaN entries are excluded.
    impact_v  : 5 threshold values read from row 1 cols J-N of the sheet (from PHAST;
                falls back to hardcoded values if the sheet's headers are broken).
                Paired with distances for interpolation intent.

    Effective behaviour (step function):
      The Excel formula calls MATCH(d, distV, -1) where distV is in ASCENDING order,
      but match_type=-1 expects DESCENDING. When given ascending distV, MATCH(-1)
      treats the first (smallest) element as the "maximum" of the array. Any d >=
      min(distV) causes MATCH to return #N/A, which the outer IFERROR catches → 0.
      Only d < min(distV) is handled by the inner IF branch → 1.0.
      Result: impact = 1.0 inside the first distance threshold, 0 everywhere else.

    Risk: uses EXP_THRESHOLDS = [limit1, limit2, prob_mid, prob_high, ...] read from
    General!J19:N19 ("Overpressure levels"), same shape as qra_v6_engine.py's
    _explosion_risk_prob(): risk = 0 below limit1, prob_mid in [limit1, limit2),
    prob_high >= limit2.

    NOTE (CVE is more complex than this — flagged, not fixed here): `impact` above is
    a BINARY 0/1 step value (the MATCH(-1) quirk), not a continuous bar-pressure value
    like v6's formula_explosion produces. Feeding a 0/1 input through bar-unit limits
    (limit1=0.04, limit2=0.1) only distinguishes "outside -> 0" from "inside -> prob_high"
    — the prob_mid band can never be hit since impact is never between limit1 and limit2.
    A real fix means either (a) making compute_exp produce an actual bar-pressure impact
    via proper interpolation (abandoning the MATCH(-1) mirror), or (b) accepting the
    binary impact and picking a deliberate two-value risk mapping instead of reusing
    General's 4-value step function as-is. Left as this degenerate-but-defined behaviour
    for now — revisit before CVE risk numbers are relied on.
    """
    dv = np.array([float(x) if x not in ('', None) else np.nan for x in distances])
    valid = ~np.isnan(dv) & (dv > 0)
    dv_v = dv[valid]

    if len(dv_v) == 0:
        return np.zeros_like(dist_m, dtype=float)

    # Step function: 1.0 inside the minimum (first) distance threshold, 0 outside.
    # This mirrors the MATCH(-1)-on-ascending-array failure in the Excel formula.
    threshold = float(np.min(dv_v))
    impact = np.zeros_like(dist_m, dtype=float)
    impact[dist_m < threshold] = 1.0

    if return_risk:
        thr = EXP_THRESHOLDS
        limit1, limit2 = thr[0], thr[1]
        prob_mid  = thr[2] if len(thr) > 2 else 0.0
        prob_high = thr[3] if len(thr) > 3 else 1.0
        risk = np.where(impact >= limit2, prob_high,
               np.where(impact >= limit1, prob_mid, 0.0))
        return frequency * risk
    return impact


# ── Sheet readers ─────────────────────────────────────────────────────────────
def _col(letter):
    """Column letter(s) -> 1-based column index."""
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n

def _v(ws, row, col_letter):
    return ws.cell(row=row, column=_col(col_letter)).value

def _flt(val):
    """Convert cell value to float or NaN."""
    if val is None or val == '':
        return np.nan
    try:
        return float(val)
    except (TypeError, ValueError):
        return np.nan


def _match_threshold_cols(ws, thresholds, col_start, col_end, row=1, sheet_name=''):
    """Match row-1 numeric headers in [col_start, col_end] to threshold values (tol 1e-4).

    Mirrors qra_v6_engine.py's _match_threshold_cols(): threshold columns in Results
    sheets are NOT guaranteed to sit at a fixed position just because they currently do —
    there's no live formula link between General's threshold array and a Results sheet's
    row-1 headers (Results headers come from the PHAST run, General is separate config).
    Reading a fixed column range blindly would silently misassign data if they ever drift
    out of sync; matching by header value instead means a mismatch produces a warning +
    a skipped (None) column, never a wrong assignment.

    Returns dict {threshold_value: 1-based absolute column index}. Thresholds with no
    matching header are absent from the dict (and logged via _warn()).
    """
    tol = 1e-4
    header_vals = {}
    for c in range(col_start, col_end + 1):
        v = ws.cell(row=row, column=c).value
        if v is None:
            continue
        try:
            header_vals[c] = float(v)
        except (TypeError, ValueError):
            continue

    matches = {}
    unmatched = []
    for t in thresholds:
        found = False
        for c, hv in header_vals.items():
            if abs(hv - t) < tol:
                matches[t] = c
                found = True
                break
        if not found:
            unmatched.append(t)

    if unmatched:
        rng = f"{ws.cell(row=row, column=col_start).coordinate}:{ws.cell(row=row, column=col_end).coordinate}"
        _warn(f"{sheet_name or ws.title}: threshold columns not found in {rng}: {unmatched}")
    return matches


def read_therm(ws):
    thresholds = list(KW_THRESHOLDS)
    col_map = _match_threshold_cols(ws, thresholds, _col('F'), _col('O'), sheet_name=ws.title)

    scenarios = []
    for r in range(2, ws.max_row + 1):
        b = _v(ws, r, 'B')
        if not b:
            break
        distances = [
            _flt(ws.cell(row=r, column=col_map[t]).value) if t in col_map else np.nan
            for t in thresholds
        ]
        scenarios.append({
            'path_code': str(b),
            'scenario':  str(_v(ws, r, 'C')),
            'weather':   str(_v(ws, r, 'D')),
            'distances': distances,
        })
    return scenarios


def read_ff(ws):
    scenarios = []
    for r in range(2, ws.max_row + 1):
        b = _v(ws, r, 'B')
        if not b:
            break
        scenarios.append({
            'path_code': str(b),
            'scenario':  str(_v(ws, r, 'C')),
            'weather':   str(_v(ws, r, 'D')),
            'lfl_r':     _flt(_v(ws, r, 'E')),
            'lflf_r':    _flt(_v(ws, r, 'F')),
        })
    return scenarios


def read_tox(ws):
    scenarios = []
    for r in range(2, 5000):    # sheet has SEQUENCE(10000) filling H:H — stop at empty B
        b = _v(ws, r, 'B')
        if not b:
            break
        blob = _v(ws, r, 'G')
        scenarios.append({
            'path_code': str(b),
            'scenario':  str(_v(ws, r, 'C')),
            'weather':   str(_v(ws, r, 'D')),
            'blob':      str(blob) if blob else '',
        })
    return scenarios


# Production CVE Results sheets have been seen with broken row-1 threshold headers
# (J1=1, K1:N1=#N/A instead of the real bar values). Fall back to the known PHAST
# values — same column order confirmed against kernel V0-4's intact J:N headers.
_EXP_IMPACT_V_FALLBACK = [0.83, 0.34, 0.1, 0.17, 0.03]

def read_exp(ws):
    # Row 1 cols J-N: impact threshold values (from PHAST — not hardcoded).
    impact_v = [_flt(_v(ws, 1, c)) for c in ['J', 'K', 'L', 'M', 'N']]
    if any(math.isnan(v) for v in impact_v):
        impact_v = list(_EXP_IMPACT_V_FALLBACK)
    scenarios = []
    for r in range(2, ws.max_row + 1):
        b = _v(ws, r, 'B')
        if not b:
            break
        scenarios.append({
            'path_code': str(b),
            'scenario':  str(_v(ws, r, 'C')),
            'weather':   str(_v(ws, r, 'D')),
            'location':  str(_v(ws, r, 'G')),
            'xign':      _flt(_v(ws, r, 'P')),
            'yign':      _flt(_v(ws, r, 'Q')),
            'distances': [_flt(_v(ws, r, c)) for c in ['J', 'K', 'L', 'M', 'N']],
            'impact_v':  impact_v,
        })
    return scenarios


def read_jf(ws):
    thresholds = list(KW_THRESHOLDS)
    # 3 separate blocks share the same threshold label set (1.6..35 repeated) — must be
    # matched independently per block, not as one row-wide search (the first match in the
    # row would otherwise win for all three).
    dist_map = _match_threshold_cols(ws, thresholds, _col('F'), _col('O'),
                                      sheet_name=f"{ws.title} (dist_v F:O)")
    half_map = _match_threshold_cols(ws, thresholds, _col('P'), _col('Y'),
                                      sheet_name=f"{ws.title} (half_v P:Y)")
    cent_map = _match_threshold_cols(ws, thresholds, _col('Z'), _col('AI'),
                                      sheet_name=f"{ws.title} (cent_v Z:AI)")

    scenarios = []
    for r in range(2, ws.max_row + 1):
        b = _v(ws, r, 'B')
        if not b:
            break
        dist_v = [_flt(ws.cell(row=r, column=dist_map[t]).value) if t in dist_map else np.nan
                  for t in thresholds]
        half_v = [_flt(ws.cell(row=r, column=half_map[t]).value) if t in half_map else np.nan
                  for t in thresholds]
        cent_v = [_flt(ws.cell(row=r, column=cent_map[t]).value) if t in cent_map else np.nan
                  for t in thresholds]
        scenarios.append({
            'path_code': str(b),
            'scenario':  str(_v(ws, r, 'C')),
            'weather':   str(_v(ws, r, 'D')),
            'flame_len': _flt(_v(ws, r, 'E')),
            'dist_v':    dist_v,
            'half_v':    half_v,
            'cent_v':    cent_v,
        })
    return scenarios


# ── Live (xlwings) sheet readers ─────────────────────────────────────────────
# Mirror the openpyxl readers above field-for-field, but read from an
# ALREADY-OPEN workbook via bulk xlwings range reads (fast — one COM round
# trip per rectangular block instead of one per cell).

def _bulk_read(ws, row_start, row_end, col_start, col_end):
    """Bulk-read an xlwings rectangular range as an always-2D list-of-rows.
    xlwings' Range.value shape depends on the range's own dimensions (scalar
    for 1x1, flat list for 1xN/Nx1, nested list for NxM) — this normalizes
    all three cases so callers can always index data[row_offset][col_offset]."""
    data = ws.range((row_start, col_start), (row_end, col_end)).value
    single_row = (row_end == row_start)
    single_col = (col_end == col_start)
    if single_row and single_col:
        return [[data]]
    if single_row:
        return [data]
    if single_col:
        return [[v] for v in data]
    return data


def _match_threshold_cols_live(header_row, thresholds, col_start, sheet_name=''):
    """Live-read equivalent of _match_threshold_cols(): header_row is a flat
    list of row-1 values starting at absolute column col_start. Returns
    {threshold_value: absolute_1based_column_index} — same contract as the
    openpyxl version, so callers index bulk-read data rows the same way in
    both the static and live readers."""
    tol = 1e-4
    header_vals = {}
    for offset, v in enumerate(header_row):
        if v is None:
            continue
        try:
            header_vals[col_start + offset] = float(v)
        except (TypeError, ValueError):
            continue

    matches = {}
    unmatched = []
    for t in thresholds:
        found = False
        for c, hv in header_vals.items():
            if abs(hv - t) < tol:
                matches[t] = c
                found = True
                break
        if not found:
            unmatched.append(t)

    if unmatched:
        _warn(f"{sheet_name}: threshold columns not found (live): {unmatched}")
    return matches


def read_therm_live(ws):
    """Live xlwings equivalent of read_therm() — identical output shape."""
    thresholds = list(KW_THRESHOLDS)
    col_f, col_o = _col('F'), _col('O')
    header_row = _bulk_read(ws, 1, 1, col_f, col_o)[0]
    col_map = _match_threshold_cols_live(header_row, thresholds, col_f, sheet_name=ws.name)

    col_b = _col('B')
    last_row = ws.used_range.last_cell.row
    rows = _bulk_read(ws, 2, last_row, col_b, col_o)

    scenarios = []
    for row in rows:
        b = row[_col('B') - col_b]
        if not b:
            break
        distances = [
            _flt(row[col_map[t] - col_b]) if t in col_map else np.nan
            for t in thresholds
        ]
        scenarios.append({
            'path_code': str(b),
            'scenario':  str(row[_col('C') - col_b]),
            'weather':   str(row[_col('D') - col_b]),
            'distances': distances,
        })
    return scenarios


def read_ff_live(ws):
    """Live xlwings equivalent of read_ff() — identical output shape."""
    col_b = _col('B')
    last_row = ws.used_range.last_cell.row
    rows = _bulk_read(ws, 2, last_row, col_b, _col('F'))

    scenarios = []
    for row in rows:
        b = row[_col('B') - col_b]
        if not b:
            break
        scenarios.append({
            'path_code': str(b),
            'scenario':  str(row[_col('C') - col_b]),
            'weather':   str(row[_col('D') - col_b]),
            'lfl_r':     _flt(row[_col('E') - col_b]),
            'lflf_r':    _flt(row[_col('F') - col_b]),
        })
    return scenarios


def read_tox_live(ws):
    """Live xlwings equivalent of read_tox() — identical output shape.
    Capped at row 5000 like read_tox(), since col H on this sheet is
    formula-filled (SEQUENCE(10000)) far beyond the real scenario data."""
    col_b = _col('B')
    last_row = min(ws.used_range.last_cell.row, 5000)
    rows = _bulk_read(ws, 2, last_row, col_b, _col('G'))

    scenarios = []
    for row in rows:
        b = row[_col('B') - col_b]
        if not b:
            break
        blob = row[_col('G') - col_b]
        scenarios.append({
            'path_code': str(b),
            'scenario':  str(row[_col('C') - col_b]),
            'weather':   str(row[_col('D') - col_b]),
            'blob':      str(blob) if blob else '',
        })
    return scenarios


def read_exp_live(ws):
    """Live xlwings equivalent of read_exp() — identical output shape."""
    col_j, col_n = _col('J'), _col('N')
    header = _bulk_read(ws, 1, 1, col_j, col_n)[0]
    impact_v = [_flt(v) for v in header]
    if any(math.isnan(v) for v in impact_v):
        impact_v = list(_EXP_IMPACT_V_FALLBACK)

    col_b = _col('B')
    last_row = ws.used_range.last_cell.row
    rows = _bulk_read(ws, 2, last_row, col_b, _col('Q'))

    scenarios = []
    for row in rows:
        b = row[_col('B') - col_b]
        if not b:
            break
        scenarios.append({
            'path_code': str(b),
            'scenario':  str(row[_col('C') - col_b]),
            'weather':   str(row[_col('D') - col_b]),
            'location':  str(row[_col('G') - col_b]),
            'xign':      _flt(row[_col('P') - col_b]),
            'yign':      _flt(row[_col('Q') - col_b]),
            'distances': [_flt(row[_col(c) - col_b]) for c in ['J', 'K', 'L', 'M', 'N']],
            'impact_v':  impact_v,
        })
    return scenarios


def read_jf_live(ws):
    """Live xlwings equivalent of read_jf() — identical output shape. Three
    separate threshold blocks (dist_v/half_v/cent_v) are matched independently,
    same as read_jf()'s dist_map/half_map/cent_map."""
    thresholds = list(KW_THRESHOLDS)

    col_f = _col('F')
    dist_header = _bulk_read(ws, 1, 1, col_f, _col('O'))[0]
    dist_map = _match_threshold_cols_live(dist_header, thresholds, col_f,
                                           sheet_name=f"{ws.name} (dist_v F:O)")

    col_p = _col('P')
    half_header = _bulk_read(ws, 1, 1, col_p, _col('Y'))[0]
    half_map = _match_threshold_cols_live(half_header, thresholds, col_p,
                                           sheet_name=f"{ws.name} (half_v P:Y)")

    col_z = _col('Z')
    cent_header = _bulk_read(ws, 1, 1, col_z, _col('AI'))[0]
    cent_map = _match_threshold_cols_live(cent_header, thresholds, col_z,
                                           sheet_name=f"{ws.name} (cent_v Z:AI)")

    col_b = _col('B')
    last_row = ws.used_range.last_cell.row
    rows = _bulk_read(ws, 2, last_row, col_b, _col('AI'))

    scenarios = []
    for row in rows:
        b = row[_col('B') - col_b]
        if not b:
            break
        dist_v = [_flt(row[dist_map[t] - col_b]) if t in dist_map else np.nan for t in thresholds]
        half_v = [_flt(row[half_map[t] - col_b]) if t in half_map else np.nan for t in thresholds]
        cent_v = [_flt(row[cent_map[t] - col_b]) if t in cent_map else np.nan for t in thresholds]
        scenarios.append({
            'path_code': str(b),
            'scenario':  str(row[_col('C') - col_b]),
            'weather':   str(row[_col('D') - col_b]),
            'flame_len': _flt(row[_col('E') - col_b]),
            'dist_v':    dist_v,
            'half_v':    half_v,
            'cent_v':    cent_v,
        })
    return scenarios


# ── Output helpers ────────────────────────────────────────────────────────────
def _safe_name(s):
    """Turn a scenario string into a safe filename stem."""
    return s.replace('/', '_').replace('\\', '_').replace(' ', '').replace('(', '').replace(')', '')


def _write_matrix_csv(path, matrix, header_lines, integer=False):
    """Write header comment lines + QYxQX numerical matrix to a CSV file."""
    with open(path, 'w', newline='') as f:
        w = csv.writer(f)
        for line in header_lines:
            w.writerow(['#' + line])
        if integer:
            for row in matrix:
                w.writerow([int(v) for v in row])
        else:
            for row in matrix:
                w.writerow([float(v) for v in row])


def _write_summary_csv(path, rows, fieldnames):
    with open(path, 'w', newline='') as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        w.writerows(rows)


def _radial_range(matrix, src_col, src_row, sx, sy, threshold):
    """Largest Euclidean distance (m) at which any cell has value >= threshold."""
    mask = matrix >= threshold
    if not mask.any():
        return 0.0
    row_idx, col_idx = np.where(mask)
    dx = sx * (col_idx + 1 - 0.5) - src_col
    dy = sy * (QY - row_idx - 0.5) - src_row
    return float(np.sqrt(dx**2 + dy**2).max())


def _to_risk(impact_mat, freq, event, exp_time=20.0):
    """Convert an impact matrix to a risk matrix (freq × lethality_probability)."""
    if event in ('thermal', 'jf'):
        return freq * _thermal_prob(impact_mat, exp_time)
    elif event == 'ff':
        # Only inside-LFL zone (FF_LEVELS[2]) contributes to individual risk
        return freq * (impact_mat >= FF_LEVELS[2]).astype(float)
    elif event == 'toxic':
        # Toxic impact is already a lethality probability
        return freq * impact_mat
    return np.zeros_like(impact_mat)


# ── General-sheet loader ──────────────────────────────────────────────────────
def _read_levels(ws, row, start_col='J', max_cols=10):
    """Read consecutive non-blank values starting at start_col (General rows 17-20 pattern)."""
    vals = []
    start = _col(start_col)
    for i in range(max_cols):
        v = ws.cell(row=row, column=start + i).value
        if v is None:
            break
        vals.append(float(v))
    return vals


def load_general_params(path=GENERAL_FILE, sheet_name=GENERAL_SHEET):
    """Read QX/QY/SX/SY + per-event threshold/level arrays from the General sheet."""
    print(f"Loading General params from {path} / {sheet_name}...")
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb[sheet_name]
    params = {
        'SX': float(_v(ws, 5, 'S')),
        'QX': int(_v(ws, 5, 'T')),
        'SY': float(_v(ws, 6, 'S')),
        'QY': int(_v(ws, 6, 'T')),
        'THERM_THRESHOLDS': _read_levels(ws, 17),   # "Thermal Levels"
        'FF_LEVELS':        _read_levels(ws, 18),   # "Flammable concentration levels"
        'EXP_THRESHOLDS':   _read_levels(ws, 19),   # "Overpressure levels"
        'TOX_LEVELS':       _read_levels(ws, 20),   # "Toxic Levels"
        'JF_N_DIRS_DEFAULT':      _v(ws, 21, 'D'),  # "Discharge directions (Default)"
        'THERM_EXP_TIME_DEFAULT': _v(ws, 22, 'D'),  # "Exposure duration (Thermal effects)"
    }
    wb.close()
    print(f"  QX={params['QX']} QY={params['QY']} SX={params['SX']} SY={params['SY']}")
    print(f"  THERM_THRESHOLDS={params['THERM_THRESHOLDS']}")
    print(f"  FF_LEVELS={params['FF_LEVELS']}  EXP_THRESHOLDS={params['EXP_THRESHOLDS']}"
          f"  TOX_LEVELS={params['TOX_LEVELS']}")
    print(f"  JF_N_DIRS_DEFAULT={params['JF_N_DIRS_DEFAULT']}"
          f"  THERM_EXP_TIME_DEFAULT={params['THERM_EXP_TIME_DEFAULT']}")
    return params


def load_general_params_live(ws):
    """Live xlwings equivalent of load_general_params() — reads directly from the
    ALREADY-OPEN workbook's General sheet instead of re-opening the file via openpyxl
    (which collides with xlwings' file lock on the same file). Same output shape."""
    print(f"Loading General params (live) from '{ws.name}'...")

    st = _bulk_read(ws, 5, 6, _col('S'), _col('T'))   # [[SX, QX], [SY, QY]]
    sx, qx = st[0]
    sy, qy = st[1]

    def _levels_live(row):
        col_j = _col('J')
        row_vals = _bulk_read(ws, row, row, col_j, col_j + 9)[0]
        vals = []
        for v in row_vals:
            if v is None:
                break
            vals.append(float(v))
        return vals

    defaults = _bulk_read(ws, 21, 22, _col('D'), _col('D'))
    params = {
        'SX': float(sx), 'QX': int(qx),
        'SY': float(sy), 'QY': int(qy),
        'THERM_THRESHOLDS': _levels_live(17),
        'FF_LEVELS':        _levels_live(18),
        'EXP_THRESHOLDS':   _levels_live(19),
        'TOX_LEVELS':       _levels_live(20),
        'JF_N_DIRS_DEFAULT':      defaults[0][0],
        'THERM_EXP_TIME_DEFAULT': defaults[1][0],
    }
    print(f"  QX={params['QX']} QY={params['QY']} SX={params['SX']} SY={params['SY']}")
    print(f"  THERM_THRESHOLDS={params['THERM_THRESHOLDS']}")
    print(f"  FF_LEVELS={params['FF_LEVELS']}  EXP_THRESHOLDS={params['EXP_THRESHOLDS']}"
          f"  TOX_LEVELS={params['TOX_LEVELS']}")
    print(f"  JF_N_DIRS_DEFAULT={params['JF_N_DIRS_DEFAULT']}"
          f"  THERM_EXP_TIME_DEFAULT={params['THERM_EXP_TIME_DEFAULT']}")
    return params


def _scenario_val(core_entry, field, default):
    """Per-scenario override from Core; falls back to a General-sheet default when empty."""
    if core_entry is not None:
        v = core_entry.get(field)
        if v not in (None, ''):
            try:
                return float(v)
            except (TypeError, ValueError):
                pass
    return float(default)


# ── Core loader ───────────────────────────────────────────────────────────────
# Column indices (0-based) in the production MacroQRAV6 `Core` sheet:
#   A(0)=key, D(3)=X, E(4)=Y, G(6)=size,
#   H(7)=Discharge directions, I(8)=Offset Angle, J(9)=Step Angle, K(10)=Exposure Time
#     (all four PER-SCENARIO — see _scenario_val() for the General-default fallback)
#   N(13)=P_TOXIC, O(14)=P_JF, P(15)=P_LPF, Q(16)=P_EPF,
#   R(17)=P_FB, S(18)=P_CVE, T(19)=P_BLV, U(20)=P_FF, V(21)=P_LEXP
_CORE_COL_DISCHARGE_DIRS = 7
_CORE_COL_OFFSET_ANGLE   = 8
_CORE_COL_STEP_ANGLE     = 9
_CORE_COL_EXP_TIME       = 10

_CORE_FREQ_IDX = {
    'P_TOXIC': 13, 'P_JF': 14, 'P_LPF': 15, 'P_EPF': 16,
    'P_FB': 17, 'P_CVE': 18, 'P_BLV': 19, 'P_FF': 20, 'P_LEXP': 21,
}

def load_core(core_path=CORE_FILE, sheet_name=CORE_SHEET):
    """Return dict {key_scenario: {X, Y, P_TOXIC, P_JF, P_LPF, ..., discharge_dirs,
    offset_angle, step_angle, exp_time}} from the Core sheet."""
    print(f"Loading Core from {core_path} / {sheet_name}...")
    wb = openpyxl.load_workbook(core_path, data_only=True, read_only=True)
    ws = wb[sheet_name]
    core = {}
    for row in ws.iter_rows(min_row=4, values_only=True):   # row 3 = header
        key = row[0]
        if not key:
            continue
        entry = {'X': row[3], 'Y': row[4]}
        entry['discharge_dirs'] = row[_CORE_COL_DISCHARGE_DIRS] if len(row) > _CORE_COL_DISCHARGE_DIRS else None
        entry['offset_angle']   = row[_CORE_COL_OFFSET_ANGLE]   if len(row) > _CORE_COL_OFFSET_ANGLE   else None
        entry['step_angle']     = row[_CORE_COL_STEP_ANGLE]     if len(row) > _CORE_COL_STEP_ANGLE     else None
        entry['exp_time']       = row[_CORE_COL_EXP_TIME]       if len(row) > _CORE_COL_EXP_TIME       else None
        for fname, idx in _CORE_FREQ_IDX.items():
            entry[fname] = row[idx] if idx < len(row) else None
        core[str(key)] = entry
    wb.close()
    print(f"  {len(core)} scenarios loaded from Core.")
    return core


def _lookup(core, key, freq_col):
    """Return (src_x, src_y, frequency) for a key, or None if missing."""
    entry = core.get(key)
    if entry is None:
        return None
    src_x = entry.get('X')
    src_y = entry.get('Y')
    freq  = entry.get(freq_col) or 0.0
    if src_x is None or src_y is None:
        return None
    return float(src_x), float(src_y), float(freq)


def _fmt_list(lst):
    """Format a list of floats/None for CSV header — compact, no huge decimals."""
    parts = []
    for v in lst:
        if v is None or (isinstance(v, float) and math.isnan(v)):
            parts.append('None')
        else:
            parts.append(f'{float(v):.4g}')
    return '[' + ', '.join(parts) + ']'


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    global QX, QY, SX, SY, KW_THRESHOLDS, FF_LEVELS, EXP_THRESHOLDS, TOX_MIN_PROB
    global JF_N_DIRS_DEFAULT, THERM_EXP_TIME_DEFAULT

    gen = load_general_params()   # reads GENERAL_FILE / GENERAL_SHEET constants at top
    QX, QY, SX, SY = gen['QX'], gen['QY'], gen['SX'], gen['SY']
    KW_THRESHOLDS = np.array(gen['THERM_THRESHOLDS'])
    FF_LEVELS = gen['FF_LEVELS']
    EXP_THRESHOLDS = gen['EXP_THRESHOLDS']
    TOX_MIN_PROB = gen['TOX_LEVELS'][0] if gen['TOX_LEVELS'] else 0.01
    JF_N_DIRS_DEFAULT = gen['JF_N_DIRS_DEFAULT']
    THERM_EXP_TIME_DEFAULT = gen['THERM_EXP_TIME_DEFAULT']

    print("Loading Macro workbook...")
    wb = openpyxl.load_workbook(INPUT_FILE, data_only=True)

    core = load_core()   # reads CORE_FILE / CORE_SHEET constants at top

    OUTPUT_DIR.mkdir(exist_ok=True)
    thermal_labels = [e['label'] for e in EVENT_CONFIG if e['formula'] == 'thermal']
    dirs = ['ff', 'ff_risk', 'toxic', 'toxic_risk', 'jf', 'jf_risk', 'exp', 'exp_risk']
    for lbl in thermal_labels:
        dirs += [f'thermal_{lbl.lower()}', f'thermal_{lbl.lower()}_risk']
    for d in dirs:
        (OUTPUT_DIR / d).mkdir(exist_ok=True)

    n_missing = {'thermal': 0, 'ff': 0, 'toxic': 0, 'jf': 0, 'exp': 0}

    # ── THERMAL ──────────────────────────────────────────────────────────────
    # 3 separate production events (EPF/LPF/FB), each its own sheet + Core frequency
    # column — the kernel's single ImpactThermMatrix collapsed all three into one.
    for evt in [e for e in EVENT_CONFIG if e['formula'] == 'thermal']:
        label, sheet_name, freq_col = evt['label'], evt['sheet'], evt['freq_col']
        print(f"\n=== THERMAL / {label} ({freq_col} from Core, sheet '{sheet_name}') ===")
        scenarios = read_therm(wb[sheet_name])
        summary = []

        for sc in scenarios:
            key = sc['path_code']
            coords = _lookup(core, key, freq_col)
            if coords is None:
                print(f"  SKIP (no Core match): {key}")
                n_missing['thermal'] += 1
                continue
            src_x, src_y, freq = coords
            exp_time = _scenario_val(core.get(key), 'exp_time', THERM_EXP_TIME_DEFAULT)

            dist_m, _, _ = _build_grids(src_x, src_y)
            mat = compute_thermal(dist_m, sc['distances'], exp_time=exp_time)

            name = _safe_name(key)
            _base_hdr = [
                f" KEY_SCENARIO  : {key}",
                f" EVENT         : thermal / {label} ({sheet_name} / {freq_col})",
                f" QX            : {QX}",
                f" QY            : {QY}",
                f" SX_m_per_cell : {SX}",
                f" SY_m_per_cell : {SY}",
                f" SRC_X_m       : {src_x}",
                f" SRC_Y_m       : {src_y}",
                f" {freq_col:<13} : {freq}",
                f" EXP_TIME_s    : {exp_time}",
                f" KW_THRESHOLDS : {[float(x) for x in KW_THRESHOLDS]}",
                f" DISTANCES_m   : {_fmt_list(sc['distances'])}",
                f" GRID_NOTE     : row 1 = north edge (y = QY*SY - SY/2); col 1 = west edge",
            ]
            dname = f'thermal_{label.lower()}'
            _write_matrix_csv(OUTPUT_DIR / dname / f"{name}.csv", mat,
                              _base_hdr + [" MATRIX_TYPE   : IMPACT  (kW/m2 interpolated)"])
            risk_mat = _to_risk(mat, freq, 'thermal', exp_time)
            _write_matrix_csv(OUTPUT_DIR / f'{dname}_risk' / f"{name}.csv", risk_mat,
                              _base_hdr + [f" MATRIX_TYPE   : RISK  ({freq_col} x lethality_prob)"])

            row_out = {'key_scenario': key, 'src_x_m': round(src_x, 4),
                       'src_y_m': round(src_y, 4), f'freq_{freq_col}': freq}
            for kw in KW_THRESHOLDS:
                row_out[f'r_{kw}kW_m'] = round(_radial_range(mat, src_x, src_y, SX, SY, kw - 0.01), 1)
            summary.append(row_out)
            print(f"  {key}")

        _write_summary_csv(
            OUTPUT_DIR / f'thermal_{label.lower()}_summary.csv', summary,
            ['key_scenario', 'src_x_m', 'src_y_m', f'freq_{freq_col}'] + [f'r_{kw}kW_m' for kw in KW_THRESHOLDS],
        )
        print(f"  -> output_v2.1/thermal_{label.lower()}  ({len(summary)} written, skipped so far: {n_missing['thermal']})")

    # ── FLASH FIRE ────────────────────────────────────────────────────────────
    print("\n=== FLASH FIRE (P_FF from Core, sheet 'Flash Fire Results') ===")
    scenarios = read_ff(wb['Flash Fire Results'])
    summary = []

    for sc in scenarios:
        key = sc['path_code']
        coords = _lookup(core, key, 'P_FF')
        if coords is None:
            print(f"  SKIP (no Core match): {key}")
            n_missing['ff'] += 1
            continue
        src_x, src_y, freq = coords

        dist_m, _, _ = _build_grids(src_x, src_y)
        mat = compute_ff(dist_m, sc['lfl_r'], sc['lflf_r'])

        name = _safe_name(key)
        _base_hdr = [
            f" KEY_SCENARIO  : {key}",
            f" EVENT         : flash fire (ImpactFFMatrix / P_FF)",
            f" QX            : {QX}",
            f" QY            : {QY}",
            f" SX_m_per_cell : {SX}",
            f" SY_m_per_cell : {SY}",
            f" SRC_X_m       : {src_x}",
            f" SRC_Y_m       : {src_y}",
            f" P_FF          : {freq}",
            f" LFL_R_m       : {sc['lfl_r']}",
            f" LFLF_R_m      : {sc['lflf_r']}",
            f" GRID_NOTE     : row 1 = north edge (y = QY*SY - SY/2); col 1 = west edge",
        ]
        _write_matrix_csv(OUTPUT_DIR / 'ff' / f"{name}.csv", mat,
                          _base_hdr + [" MATRIX_TYPE   : IMPACT  (0=safe 1=LFLF-zone 2=LFL-zone)"],
                          integer=True)
        risk_mat = _to_risk(mat, freq, 'ff')
        _write_matrix_csv(OUTPUT_DIR / 'ff_risk' / f"{name}.csv", risk_mat,
                          _base_hdr + [" MATRIX_TYPE   : RISK  (P_FF x 1_inside_LFL)"])

        lfl_comp  = round(_radial_range(mat, src_x, src_y, SX, SY, 1.9), 1)
        lflf_comp = round(_radial_range(mat, src_x, src_y, SX, SY, 0.9), 1)
        summary.append({
            'key_scenario': key, 'src_x_m': round(src_x, 4), 'src_y_m': round(src_y, 4),
            'freq_P_FF': freq,
            'lfl_r_input_m': sc['lfl_r'], 'lflf_r_input_m': sc['lflf_r'],
            'computed_lfl_r_m': lfl_comp, 'computed_lflf_r_m': lflf_comp,
        })
        print(f"  {key}")

    _write_summary_csv(
        OUTPUT_DIR / 'ff_summary.csv', summary,
        ['key_scenario','src_x_m','src_y_m','freq_P_FF',
         'lfl_r_input_m','lflf_r_input_m','computed_lfl_r_m','computed_lflf_r_m'],
    )
    print(f"  -> output_v2.1/ff  ({len(summary)} written, {n_missing['ff']} skipped)")

    # ── TOXIC ─────────────────────────────────────────────────────────────────
    print(f"\n=== TOXIC (P_TOXIC from Core, sheet 'Outdoor Toxic Results', blob unit={TOXIC_UNIT_MODE!r}) ===")
    scenarios = read_tox(wb['Outdoor Toxic Results'])
    summary = []

    for sc in scenarios:
        key = sc['path_code']
        coords = _lookup(core, key, 'P_TOXIC')
        if coords is None:
            print(f"  SKIP (no Core match): {key}")
            n_missing['toxic'] += 1
            continue
        src_x, src_y, freq = coords

        dist_m, _, _ = _build_grids(src_x, src_y)
        mat = compute_toxic(dist_m, sc['blob'])

        name = _safe_name(key)
        _base_hdr = [
            f" KEY_SCENARIO  : {key}",
            f" EVENT         : toxic (ImpactToxMatrix / P_TOXIC)",
            f" QX            : {QX}",
            f" QY            : {QY}",
            f" SX_m_per_cell : {SX}",
            f" SY_m_per_cell : {SY}",
            f" SRC_X_m       : {src_x}",
            f" SRC_Y_m       : {src_y}",
            f" P_TOXIC       : {freq}",
            f" BLOB_UNIT_MODE: {TOXIC_UNIT_MODE}",
            f" GRID_NOTE     : row 1 = north edge (y = QY*SY - SY/2); col 1 = west edge",
        ]
        _write_matrix_csv(OUTPUT_DIR / 'toxic' / f"{name}.csv", mat,
                          _base_hdr + [" MATRIX_TYPE   : IMPACT  (lethality probability 0-1)"])
        risk_mat = _to_risk(mat, freq, 'toxic')
        _write_matrix_csv(OUTPUT_DIR / 'toxic_risk' / f"{name}.csv", risk_mat,
                          _base_hdr + [" MATRIX_TYPE   : RISK  (P_TOXIC x lethality_prob)"])

        row_out = {'key_scenario': key, 'src_x_m': round(src_x, 4),
                   'src_y_m': round(src_y, 4), 'freq_P_TOXIC': freq}
        for p_thr in [0.01, 0.10, 0.50, 0.90]:
            row_out[f'r_{int(p_thr*100)}pct_m'] = round(
                _radial_range(mat, src_x, src_y, SX, SY, p_thr), 1)
        summary.append(row_out)
        print(f"  {key}")

    _write_summary_csv(
        OUTPUT_DIR / 'toxic_summary.csv', summary,
        ['key_scenario','src_x_m','src_y_m','freq_P_TOXIC',
         'r_1pct_m','r_10pct_m','r_50pct_m','r_90pct_m'],
    )
    print(f"  -> output_v2.1/toxic  ({len(summary)} written, {n_missing['toxic']} skipped)")

    # ── EXPLOSION (CVE) ───────────────────────────────────────────────────────
    # Source coords are Xign/Yign from col P/Q of each row (NOT from Core).
    # Frequency (P_CVE) still comes from Core.
    # The kernel formula uses MATCH(AA3, B:B, 0) = first row per unique key.
    # We generate one CSV per unique key (first matching row) to match kernel behaviour.
    print("\n=== EXPLOSION (P_CVE from Core, sheet 'CVE Results', coords from Xign/Yign) ===")
    scenarios = read_exp(wb['CVE Results'])
    summary = []

    first_seen = {}   # key -> first scenario row (mirrors kernel MATCH first-match)
    for sc in scenarios:
        key = sc['path_code']
        if key not in first_seen:
            first_seen[key] = sc

    for key, sc in first_seen.items():
        xign = sc['xign']
        yign = sc['yign']
        if math.isnan(xign) or math.isnan(yign):
            print(f"  SKIP (no Xign/Yign): {key}")
            n_missing['exp'] += 1
            continue

        # Frequency from Core (P_CVE); fall back to 0 if key not in Core
        core_entry = core.get(key)
        freq = float(core_entry['P_CVE'] or 0.0) if core_entry and core_entry.get('P_CVE') else 0.0

        dist_m, _, _ = _build_grids(xign, yign)
        mat = compute_exp(dist_m, sc['distances'], sc['impact_v'])

        name = _safe_name(key)
        _base_hdr = [
            f" KEY_SCENARIO  : {key}",
            f" EVENT         : explosion (ImpactExpMatrix / P_CVE)",
            f" QX            : {QX}",
            f" QY            : {QY}",
            f" SX_m_per_cell : {SX}",
            f" SY_m_per_cell : {SY}",
            f" XIGN_m        : {xign}",
            f" YIGN_m        : {yign}",
            f" P_CVE         : {freq}",
            f" LOCATION      : {sc['location']}",
            f" IMPACT_V      : {_fmt_list(sc['impact_v'])}",
            f" DISTANCES_m   : {_fmt_list(sc['distances'])}",
            f" GRID_NOTE     : row 1 = north edge (y = QY*SY - SY/2); col 1 = west edge",
        ]
        _write_matrix_csv(OUTPUT_DIR / 'exp' / f"{name}.csv", mat,
                          _base_hdr + [" MATRIX_TYPE   : IMPACT  (step fn: 1 inside min(distV), 0 outside)"])
        risk_mat = compute_exp(dist_m, sc['distances'], sc['impact_v'], return_risk=True, frequency=freq)
        _write_matrix_csv(OUTPUT_DIR / 'exp_risk' / f"{name}.csv", risk_mat,
                          _base_hdr + [" MATRIX_TYPE   : RISK  (P_CVE x EXP_THRESHOLDS step fn — see compute_exp() CVE note)"])

        summary.append({
            'key_scenario': key,
            'xign_m': round(xign, 4),
            'yign_m': round(yign, 4),
            'location': sc['location'],
            'freq_P_CVE': freq,
            'distances_m': _fmt_list(sc['distances']),
        })
        print(f"  {key}  [{sc['location']}]")

    _write_summary_csv(
        OUTPUT_DIR / 'exp_summary.csv', summary,
        ['key_scenario', 'xign_m', 'yign_m', 'location', 'freq_P_CVE', 'distances_m'],
    )
    print(f"  -> output_v2.1/exp  ({len(summary)} written, {n_missing['exp']} skipped)")

    # ── JET FIRE ──────────────────────────────────────────────────────────────
    print("\n=== JET FIRE (P_JF from Core, sheet 'Jet Fire Results') ===")
    scenarios = read_jf(wb['Jet Fire Results'])
    summary = []

    for sc in scenarios:
        key = sc['path_code']
        coords = _lookup(core, key, 'P_JF')
        if coords is None:
            print(f"  SKIP (no Core match): {key}")
            n_missing['jf'] += 1
            continue
        src_x, src_y, freq = coords
        core_entry   = core.get(key)
        n_dirs       = int(_scenario_val(core_entry, 'discharge_dirs', JF_N_DIRS_DEFAULT))
        angle_offset = _scenario_val(core_entry, 'offset_angle', JF_ANGLE_OFFSET_DEFAULT)
        angle_step   = _scenario_val(core_entry, 'step_angle', JF_ANGLE_STEP_DEFAULT)
        exp_time     = _scenario_val(core_entry, 'exp_time', THERM_EXP_TIME_DEFAULT)

        _, x_c, y_c = _build_grids(src_x, src_y)
        mat = compute_jf(x_c, y_c, sc['dist_v'], sc['half_v'], sc['cent_v'],
                         n_dirs=n_dirs, angle_offset=angle_offset, angle_step=angle_step,
                         exp_time=exp_time)

        dv0 = sc['dist_v'][0]
        try:
            reach_m = float(dv0) * SX if dv0 not in ('', None) else 0.0
        except (TypeError, ValueError):
            reach_m = 0.0

        name = _safe_name(key)
        _base_hdr = [
            f" KEY_SCENARIO  : {key}",
            f" EVENT         : jet fire (ImpactJFMatrix / P_JF)",
            f" QX            : {QX}",
            f" QY            : {QY}",
            f" SX_m_per_cell : {SX}",
            f" SY_m_per_cell : {SY}",
            f" SRC_X_m       : {src_x}",
            f" SRC_Y_m       : {src_y}",
            f" P_JF          : {freq}",
            f" EXP_TIME_s    : {exp_time}",
            f" FLAME_LEN_m   : {sc['flame_len']}",
            f" N_DIRS        : {n_dirs}  (offset={angle_offset} deg, step={angle_step} deg"
            f"{' -> uniform 360/N' if angle_step == 0 else ''})",
            f" KW_THRESHOLDS : {[float(x) for x in KW_THRESHOLDS]}",
            f" distV_cells   : {_fmt_list(sc['dist_v'])}",
            f" halfWV_cells  : {_fmt_list(sc['half_v'])}",
            f" centV_cells   : {_fmt_list(sc['cent_v'])}",
            f" distV_m       : {_fmt_list([v * SX if not (math.isnan(v) if isinstance(v, float) else False) else float('nan') for v in sc['dist_v']])}",
            f" GRID_NOTE     : row 1 = north edge (y = QY*SY - SY/2); col 1 = west edge",
        ]
        _write_matrix_csv(OUTPUT_DIR / 'jf' / f"{name}.csv", mat,
                          _base_hdr + [f" MATRIX_TYPE   : IMPACT  (kW/m2 MAX ellipse over {n_dirs} dirs x 10 thresholds)"])
        risk_mat = _to_risk(mat, freq, 'jf', exp_time)
        _write_matrix_csv(OUTPUT_DIR / 'jf_risk' / f"{name}.csv", risk_mat,
                          _base_hdr + [" MATRIX_TYPE   : RISK  (P_JF x lethality_prob)"])

        row_out = {'key_scenario': key, 'src_x_m': round(src_x, 4),
                   'src_y_m': round(src_y, 4), 'freq_P_JF': freq,
                   'flame_len_m': sc['flame_len'],
                   'reach_1.6kW_m': round(reach_m, 1)}
        for kw in KW_THRESHOLDS:
            row_out[f'r_{kw}kW_m'] = round(
                _radial_range(mat, src_x, src_y, SX, SY, kw - 0.01), 1)
        summary.append(row_out)
        print(f"  {key}")

    _write_summary_csv(
        OUTPUT_DIR / 'jf_summary.csv', summary,
        ['key_scenario','src_x_m','src_y_m','freq_P_JF','flame_len_m','reach_1.6kW_m']
        + [f'r_{kw}kW_m' for kw in KW_THRESHOLDS],
    )
    print(f"  -> output_v2.1/jf  ({len(summary)} written, {n_missing['jf']} skipped)")

    print(f"\nAll done.  Results in: {OUTPUT_DIR.resolve()}")
    total_skip = sum(n_missing.values())
    if total_skip:
        print(f"  {total_skip} scenarios had no Core match (see SKIP lines above)")

    print_warning_summary(n_missing)


def print_warning_summary(n_missing=None):
    """Print every structural warning collected via _warn() in one place, plus the
    per-event skip counts. Meant to be the one thing you have to check at the end of
    a run — important when this engine is eventually driven from Excel instead of a
    console, where scrolled-past per-line prints are easy to miss."""
    sep = '=' * 65
    print(f"\n{sep}")
    print(f"RUN SUMMARY")
    print(sep)
    if n_missing:
        for evt, n in n_missing.items():
            if n:
                print(f"  {evt}: {n} scenario(s) skipped (no Core match)")
    if _warnings:
        print(f"\n  {len(_warnings)} structural warning(s):")
        for i, msg in enumerate(_warnings, 1):
            print(f"    {i:>3}. {msg}")
    else:
        print("  No structural warnings.")
    print(sep)


# ══════════════════════════════════════════════════════════════════════════════
# Live Core/PageControl iteration (xlwings) — per-size accumulated matrices
# ══════════════════════════════════════════════════════════════════════════════
# main() above is the static, per-scenario CSV mode (openpyxl only, no Excel needed).
# Everything below is the live mode: opens the Macro in Excel (VBA disabled), iterates
# Core!C2 per active Impact ID from PageControl (clearing the Core!H2 size filter first
# so every size shows up in one pass), accumulates impact/risk matrices per size
# (Total/S/M/L/XL/INST), writes CSV checkpoints, and writes the matrices into
# 'Impact Matrix Result'/'Risk Matrix Result' in a SAVED COPY (never the original file).

def _index_by_key(rows):
    """First-row-wins index by path_code — mirrors Excel's MATCH(key, B:B, 0)."""
    idx = {}
    for r in rows:
        k = r['path_code']
        if k not in idx:
            idx[k] = r
    return idx


def _normalize_size(v):
    """Map a raw Core!G size label to one of SIZES (case-insensitive), or None."""
    if v is None:
        return None
    return _SIZE_NORM.get(str(v).strip().upper())


# ── PageControl (mirrors qra_v6_engine.py) ────────────────────────────────────
_PC_READ_RANGE = 'J2:R500'
_PC_EVENT_COL  = 0   # offset from J -> event description
_PC_ID_COL     = 2   # offset from J -> Impact ID (column L)
_PC_SIZE_COLS  = {'Total': 3, 'S': 4, 'M': 5, 'L': 6, 'XL': 7, 'INST': 8}   # offsets from J


def _as_text(v):
    return str(v).strip() if v is not None else ''


def _looks_like_range(v):
    return bool(re.match(r'^[\$A-Za-z]+[\$\d]+:[\$A-Za-z]+[\$\d]+$', _as_text(v).replace('$', '')))


def _row_ranges(row):
    return {sz: _as_text(row[col]) for sz, col in _PC_SIZE_COLS.items() if _looks_like_range(row[col])}


def read_page_control(ws_pc):
    """Return the set of active Impact IDs from PageControl!J2:R500 — a row counts as
    active if col L parses as an int AND at least one size column (M:R) holds a range."""
    active = set()
    data = ws_pc.range(_PC_READ_RANGE).value
    if not data:
        _warn('PageControl appears empty — treating all EVENT_CONFIG entries as active.')
        return {e['impact_id'] for e in EVENT_CONFIG}
    for row in data:
        if not row or all(c is None for c in row):
            if active:
                break
            continue
        try:
            iid = int(row[_PC_ID_COL])
        except (TypeError, ValueError):
            continue
        if not _row_ranges(row):
            continue
        active.add(iid)
        name = _as_text(row[_PC_EVENT_COL])
        if name:
            print(f"    Impact ID {iid}: {name}")
    if not active:
        _warn('No active impacts found in PageControl — treating all EVENT_CONFIG entries as active.')
        return {e['impact_id'] for e in EVENT_CONFIG}
    return active


def read_directions(ws_pc):
    """Return {impact_id: {size: range_str}} — destination ranges for the final write-back."""
    directions = {}
    data = ws_pc.range(_PC_READ_RANGE).value
    if not data:
        return directions
    for row in data:
        if not row or all(c is None for c in row):
            if directions:
                break
            continue
        try:
            iid = int(row[_PC_ID_COL])
        except (TypeError, ValueError):
            continue
        rng = _row_ranges(row)
        if rng:
            directions[iid] = rng
    return directions


def _parse_range_start(range_str):
    """'$LD$318:$XF$634' -> (start_row, start_col) 1-based."""
    clean = range_str.replace('$', '').split(':')[0]
    m = re.match(r'([A-Za-z]+)(\d+)', clean)
    if not m:
        raise ValueError(f'Cannot parse range: {range_str!r}')
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - ord('A') + 1)
    return int(m.group(2)), col


# ── Live Core reader (xlwings) ────────────────────────────────────────────────
_CORE_LIVE_RANGE = 'A4:W500'   # wide enough for the current N:V frequency block + W "Filter" helper


def read_core_live(ws_core):
    """Read Core!A4:W500 from the CURRENT live/recalculated state (after C2/H2 are set).
    One dict per row — same field names as load_core()'s per-key entries, so
    _scenario_val() works unchanged against either."""
    data = ws_core.range(_CORE_LIVE_RANGE).value
    scenarios = []
    if not data:
        return scenarios
    for row in data:
        if not row:
            break
        key = row[0]
        if key is None or key == 0 or key == -1:
            break
        entry = {
            'key': str(key),
            'X': row[3], 'Y': row[4],
            'size': row[6],
            'discharge_dirs': row[_CORE_COL_DISCHARGE_DIRS],
            'offset_angle':   row[_CORE_COL_OFFSET_ANGLE],
            'step_angle':     row[_CORE_COL_STEP_ANGLE],
            'exp_time':       row[_CORE_COL_EXP_TIME],
        }
        for fname, idx in _CORE_FREQ_IDX.items():
            entry[fname] = row[idx] if idx < len(row) else None
        scenarios.append(entry)
    return scenarios


# ── Per-size accumulation ─────────────────────────────────────────────────────
def compute_event_live(evt, core_scenarios, results_idx):
    """Accumulate impact/risk matrices per size for one event (one Impact ID).

    core_scenarios : list of dicts from read_core_live() — the LIVE Core row set for
                      this event (already reflecting the current C2/H2 state).
    results_idx     : {path_code: result_row} from _index_by_key() on the matching
                      Results-sheet reader output (read_therm/read_ff/read_tox/read_exp/read_jf).

    Returns (impact_mats, risk_mats, matched, unmatched) — each *_mats is
    {size: (QY,QX) ndarray} for size in SIZES.
    """
    formula  = evt['formula']
    freq_col = evt['freq_col']
    impact_mats = {sz: np.zeros((QY, QX)) for sz in SIZES}
    risk_mats   = {sz: np.zeros((QY, QX)) for sz in SIZES}
    matched = unmatched = 0

    for core_row in core_scenarios:
        key = core_row['key']
        res = results_idx.get(key)
        if res is None:
            unmatched += 1
            continue

        freq     = float(core_row.get(freq_col) or 0.0)
        size     = _normalize_size(core_row.get('size'))
        exp_time = _scenario_val(core_row, 'exp_time', THERM_EXP_TIME_DEFAULT)
        src_x, src_y = core_row.get('X'), core_row.get('Y')

        if formula in ('thermal', 'ff', 'toxic', 'jf') and (src_x is None or src_y is None):
            unmatched += 1
            continue

        if formula == 'thermal':
            dist_m, _, _ = _build_grids(float(src_x), float(src_y))
            imp  = compute_thermal(dist_m, res['distances'], exp_time=exp_time)
            risk = _to_risk(imp, freq, 'thermal', exp_time)

        elif formula == 'ff':
            dist_m, _, _ = _build_grids(float(src_x), float(src_y))
            imp  = compute_ff(dist_m, res['lfl_r'], res['lflf_r'])
            risk = _to_risk(imp, freq, 'ff')

        elif formula == 'toxic':
            dist_m, _, _ = _build_grids(float(src_x), float(src_y))
            imp  = compute_toxic(dist_m, res['blob'])
            risk = _to_risk(imp, freq, 'toxic')

        elif formula == 'exp':
            xign, yign = res.get('xign'), res.get('yign')
            if xign is None or yign is None or math.isnan(xign) or math.isnan(yign):
                unmatched += 1
                continue
            dist_m, _, _ = _build_grids(xign, yign)
            imp  = compute_exp(dist_m, res['distances'], res['impact_v'])
            risk = compute_exp(dist_m, res['distances'], res['impact_v'], return_risk=True, frequency=freq)

        elif formula == 'jf':
            n_dirs       = int(_scenario_val(core_row, 'discharge_dirs', JF_N_DIRS_DEFAULT))
            angle_offset = _scenario_val(core_row, 'offset_angle', JF_ANGLE_OFFSET_DEFAULT)
            angle_step   = _scenario_val(core_row, 'step_angle', JF_ANGLE_STEP_DEFAULT)
            _, x_c, y_c = _build_grids(float(src_x), float(src_y))
            imp = compute_jf(x_c, y_c, res['dist_v'], res['half_v'], res['cent_v'],
                             n_dirs=n_dirs, angle_offset=angle_offset, angle_step=angle_step,
                             exp_time=exp_time)
            risk = _to_risk(imp, freq, 'jf', exp_time)

        else:
            unmatched += 1
            continue

        matched += 1
        impact_mats['Total'] += imp
        risk_mats['Total']   += risk
        if size is not None:
            impact_mats[size] += imp
            risk_mats[size]   += risk

    return impact_mats, risk_mats, matched, unmatched


_LIVE_READER_FN = {
    'thermal': read_therm_live,
    'ff':      read_ff_live,
    'toxic':   read_tox_live,
    'exp':     read_exp_live,
    'jf':      read_jf_live,
}


def _run_pipeline(wb, app):
    """Shared Core/PageControl iteration — per-size accumulated impact/risk
    matrices, written into SHEET_IMPACT/SHEET_RISK. Used by both run_live()
    (CLI, opens its own xw.App) and run_from_button() (reuses the caller's
    already-open app/workbook). Returns the all_results dict. Populates the
    module-level _timings dict (reading/calculation/iteration/writing)."""
    global QX, QY, SX, SY, KW_THRESHOLDS, FF_LEVELS, EXP_THRESHOLDS, TOX_MIN_PROB
    global JF_N_DIRS_DEFAULT, THERM_EXP_TIME_DEFAULT

    _timings.clear()

    with _timed('reading'):
        gen = load_general_params_live(wb.sheets[GENERAL_SHEET])
        QX, QY, SX, SY = gen['QX'], gen['QY'], gen['SX'], gen['SY']
        KW_THRESHOLDS = np.array(gen['THERM_THRESHOLDS'])
        FF_LEVELS = gen['FF_LEVELS']
        EXP_THRESHOLDS = gen['EXP_THRESHOLDS']
        TOX_MIN_PROB = gen['TOX_LEVELS'][0] if gen['TOX_LEVELS'] else 0.01
        JF_N_DIRS_DEFAULT = gen['JF_N_DIRS_DEFAULT']
        THERM_EXP_TIME_DEFAULT = gen['THERM_EXP_TIME_DEFAULT']

        ws_core = wb.sheets['Core']
        ws_pc   = wb.sheets['PageControl']

        print("Reading PageControl...")
        active_impacts = read_page_control(ws_pc)
        directions      = read_directions(ws_pc)

    runnable = [e for e in EVENT_CONFIG if e['impact_id'] in active_impacts and e['formula']]
    n = len(runnable)

    all_results = {}
    for i, evt in enumerate(runnable, 1):
        iid = evt['impact_id']
        print(f"\n{'-'*65}")
        print(f"  Impact {iid} | {evt['label']} | {evt['sheet']}")
        print(f"{'-'*65}")
        _progress(app, f"Procesando {evt['label']} ({i}/{n})...")

        with _timed('calculation'):
            # Clear the size filter FIRST, then set the Impact ID — order matters:
            # H2 must be empty so every size shows up in this one Core read.
            ws_core.range('H2').value = None
            ws_core.range('C2').value = iid
            time.sleep(1.5)
            app.api.Calculate()
            time.sleep(1.0)

            h2 = ws_core.range('H2').value
            if h2:
                _warn(f"Impact {iid} ({evt['label']}): Core!H2 did not stay empty "
                      f"(still {h2!r}) — some sizes may be missing from this read.")

        with _timed('reading'):
            core_scenarios = read_core_live(ws_core)
            print(f"  Core scenarios (live, all sizes): {len(core_scenarios)}")

            live_reader = _LIVE_READER_FN[evt['formula']]
            results_rows = live_reader(wb.sheets[evt['sheet']])
            results_idx = _index_by_key(results_rows)
            print(f"  {evt['sheet']}: {len(results_rows)} rows read (live)")

        with _timed('iteration'):
            impact_mats, risk_mats, matched, unmatched = compute_event_live(evt, core_scenarios, results_idx)

        print(f"  matched={matched}  unmatched={unmatched}")
        if unmatched:
            _warn(f"Impact {iid} ({evt['label']}): {unmatched} Core scenario(s) had no "
                  f"matching row in '{evt['sheet']}' (or missing X/Y/Xign-Yign).")

        for sz in SIZES:
            print(f"    [{sz:5s}] impact_max={impact_mats[sz].max():.4g}  "
                  f"risk_max={risk_mats[sz].max():.4e}  nonzero={np.count_nonzero(impact_mats[sz])}")

        all_results[iid] = {'impact': impact_mats, 'risk': risk_mats, 'label': evt['label']}

    for evt in EVENT_CONFIG:
        if evt['impact_id'] in active_impacts and not evt['formula']:
            _warn(f"Impact {evt['impact_id']} ({evt['label']}): no reader/formula "
                  f"implemented yet — skipped.")

    print(f"\n{'='*65}\nWriting matrices back into the workbook...\n{'='*65}")
    with _timed('writing'):
        _write_results_to_workbook(wb, all_results, directions)

    print(f"\n{_format_timing_report(_timings)}")

    return all_results


# ── Combined matrices (final union across events / sizes) ────────────────────
# Confirmed directly from the Macro's VBA source (qra_v6_engine.py-adjacent analysis,
# see parameters_joint.md): a separate "MatrixOp" workbook (General!D9 path) combines
# per-event/per-size results with two DIFFERENT cell-wise operations:
#   RISK   -> SUM across events, for one leak size ("Risk-Sum" calc sheet in MatrixOp;
#             same pattern confirmed in the IRPA/LSIR point-based calc: each event
#             contributes its own risk column, TOTAL = SUM(...) of those columns).
#             Physically valid: risk is frequency x fatality-probability (1/yr), additive
#             across independent hazard scenarios.
#   IMPACT -> MAX across leak sizes, WITHIN one event only ("Impact-Max" calc sheet;
#             the VBA loop Exit-Fors after the first selected event — impact is never
#             combined ACROSS events, since kW/m², bar, and lethality-fraction aren't
#             comparable/summable units).
# Both are cell-wise (position (i,j) of the result depends only on position (i,j) of each
# input matrix) — valid because every matrix here has the same (QY, QX) shape, built from
# the same _build_grids() call, so cell (i,j) means the same physical location everywhere.

def combine_risk(all_results, size):
    """RISK: cell-wise SUM across every computed event, for one leak size."""
    combined = np.zeros((QY, QX))
    for res in all_results.values():
        combined += res['risk'][size]
    return combined


def combine_impact_envelope(evt_result):
    """IMPACT: cell-wise MAX across leak sizes S/M/L/XL/INST, for ONE event only."""
    stacks = [evt_result['impact'][sz] for sz in ('S', 'M', 'L', 'XL', 'INST')]
    return np.maximum.reduce(stacks)


# ── Write-back into the workbook (saved copy, not the original) ──────────────
def _get_or_create_sheet(wb, name):
    names = [s.name for s in wb.sheets]
    if name in names:
        ws = wb.sheets[name]
        ws.clear()
    else:
        ws = wb.sheets.add(name)
    return ws


def _write_results_to_workbook(wb, all_results, directions):
    ws_imp  = _get_or_create_sheet(wb, SHEET_IMPACT)
    ws_risk = _get_or_create_sheet(wb, SHEET_RISK)
    written = 0
    for iid, res in all_results.items():
        dir_map = directions.get(iid, {})
        if not dir_map:
            _warn(f"Impact {iid} ({res['label']}): no destination range in PageControl "
                  f"— matrices not written to the workbook.")
            continue
        for sz, range_str in dir_map.items():
            if sz not in res['impact']:
                continue
            try:
                r0, c0 = _parse_range_start(range_str)
            except ValueError as e:
                _warn(str(e))
                continue
            ws_imp.range((r0, c0)).value  = res['impact'][sz].tolist()
            ws_risk.range((r0, c0)).value = res['risk'][sz].tolist()
            written += 1
    print(f"  Wrote {written} impact/risk block(s) to '{SHEET_IMPACT}'/'{SHEET_RISK}'.")


# ── Orchestration ──────────────────────────────────────────────────────────────
def run_live():
    print("\nOpening Macro workbook in Excel (VBA disabled)...")
    app = xw.App(visible=False)
    app.display_alerts = False
    app.screen_updating = False
    app.api.AutomationSecurity = 3   # msoAutomationSecurityForceDisable — no VBA runs

    OUTPUT_DIR.mkdir(exist_ok=True)
    (OUTPUT_DIR / 'matrices').mkdir(exist_ok=True)

    try:
        wb = app.books.open(INPUT_FILE)
        all_results = _run_pipeline(wb, app)

        # CSV checkpoint per event/size — CLI-only convenience, not part of the
        # shared pipeline (the button path doesn't write these).
        for iid, res in all_results.items():
            for sz in SIZES:
                hdr = [f" IMPACT_ID : {iid}", f" EVENT     : {res['label']}",
                       f" SIZE      : {sz}", f" QX/QY     : {QX}/{QY}", f" SX/SY     : {SX}/{SY}"]
                _write_matrix_csv(OUTPUT_DIR / 'matrices' / f"{res['label']}_{sz}_impact.csv",
                                  res['impact'][sz], hdr + [" MATRIX_TYPE : IMPACT"])
                _write_matrix_csv(OUTPUT_DIR / 'matrices' / f"{res['label']}_{sz}_risk.csv",
                                  res['risk'][sz], hdr + [" MATRIX_TYPE : RISK"])

        # Phase 2b: combined matrices (final union across events / sizes) — CSV only,
        # CLI-specific, not wired to CalcBUImpact/CalcBURisk. See docs spec, "Explicitly
        # out of scope".
        print(f"\n{'='*65}\nCombined matrices (risk: SUM across events; impact: MAX across sizes)\n{'='*65}")

        print("  RISK — sum across events, per size:")
        for sz in SIZES:
            mat = combine_risk(all_results, sz)
            print(f"    [{sz:5s}] risk_max={mat.max():.4e}  nonzero={np.count_nonzero(mat)}")
            hdr = [f" SIZE     : {sz}", f" QX/QY    : {QX}/{QY}", f" SX/SY    : {SX}/{SY}",
                   " COMBINE  : SUM across events (cell-wise, position (i,j) summed independently)"]
            _write_matrix_csv(OUTPUT_DIR / 'matrices' / f"COMBINED_RISK_{sz}.csv", mat,
                              hdr + [" MATRIX_TYPE : RISK (combined)"])

        print("  IMPACT — max across sizes (S/M/L/XL/INST), per event:")
        for iid, res in all_results.items():
            mat = combine_impact_envelope(res)
            print(f"    [{res['label']:5s}] impact_max={mat.max():.4g}  nonzero={np.count_nonzero(mat)}")
            hdr = [f" IMPACT_ID : {iid}", f" EVENT     : {res['label']}", f" QX/QY     : {QX}/{QY}",
                   f" SX/SY     : {SX}/{SY}",
                   " COMBINE   : MAX across sizes S/M/L/XL/INST (cell-wise; NOT combined across events)"]
            _write_matrix_csv(OUTPUT_DIR / 'matrices' / f"{res['label']}_IMPACT_ENVELOPE.csv", mat,
                              hdr + [" MATRIX_TYPE : IMPACT (envelope)"])

        export_path = str(Path(INPUT_FILE).resolve().parent / EXPORT_FILE)
        wb.save(export_path)
        print(f"  Saved: {export_path}")
        wb.close()
    finally:
        app.quit()

    print_warning_summary()


def run_from_button(wb=None):
    """Entry point called from the IRiskMap button via
    RunPython "import qra_engine_v2_1 as eng; eng.run_from_button()".
    wb=None (real button path) resolves the ALREADY open caller workbook via
    xw.Book.caller(); an explicit wb (headless test path) is used as-is."""
    _warnings.clear()   # defensive — RunPython normally spawns a fresh process
                         # per click, but the headless test path can call this
                         # more than once in the same interpreter.

    if wb is None:
        wb = xw.Book.caller()
    app = wb.app

    with _frozen(app):
        try:
            _run_pipeline(wb, app)
        except Exception as exc:
            _warn(f"ERROR FATAL: {exc}")
        finally:
            app.api.StatusBar = False

    summary = _format_warning_summary(_warnings) + "\n\n" + _format_timing_report(_timings)
    _show_msgbox(summary)


if __name__ == '__main__':
    import sys
    if '--live' in sys.argv:
        run_live()
    else:
        main()
