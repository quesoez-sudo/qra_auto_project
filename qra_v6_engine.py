"""
QRA Engine V6 — Dynamic Parameter Reading
==========================================
Reads all parameters (grid, thresholds, exposure time) from the General sheet of
MacroQRAV6. Uses the actual consequence result sheets for calculations, matching
threshold columns dynamically instead of assuming fixed column positions.

Pipeline
--------
1. Open MacroQRAV6 (version 1).xlsm via xlwings
2. Load and print parameters from General sheet (verify before accepting results)
3. Read PageControl  → active Impact IDs + destination ranges
4. For each active Impact ID:
     a. Set Core!C2 = impact_id  →  force recalculation
     b. Warn if Core!G2 hasIs  a size filter active (can hide scenarios)
     c. Read Core scenarios (A4:R89); skip sentinel rows (0 or -1)
     d. Read the corresponding effect results sheet with dynamic column matching
     e. Compute impact & risk matrices on the QY×QX grid
5. Create/overwrite "Impact Matrix Result" and "Risk Matrix Result" sheets
6. Write each matrix at the destination range from PageControl (bulk COM write)
7. Save workbook as macro-disabled .xlsx
8. Print warning/error summary

Notes on dynamic column matching
---------------------------------
Thermal result sheets (Jet Fire, Pool Fires, Fireball) have row-1 headers that
are the kW/m² threshold values pulled from General row 17.  Only the thresholds
that PHAST actually computed will have numeric headers; unused columns are None.
The engine matches General-sheet thresholds to those headers and only reads the
columns that match.  Columns with None or non-matching headers are silently
skipped (their threshold slot is set to None → contributes nothing to impact).

CVE/BLEVE result sheets have row-1 headers that should be the bar overpressure
thresholds from General row 19.  If those headers are wrong (e.g. 1/-1 due to a
formula error in the workbook) the engine falls back to POSITIONAL column
assignment (J=first threshold, K=second, …) and logs a warning.
"""

import os
import re
import sys
import time

import numpy as np
import xlwings as xw
from dotenv import load_dotenv

sys.stdout.reconfigure(encoding='utf-8')

load_dotenv()

# ── Paths ─────────────────────────────────────────────────────────────────────
# _WORKSPACE is loaded from .env file; falls back to current directory if not found
_WORKSPACE   = os.getenv('WORKSPACE', os.getcwd())
V6_PATH      = os.path.join(_WORKSPACE, 'MacroQRAV6 (version 1).xlsm')
EXPORT_PATH  = os.path.join(_WORKSPACE, 'MacroQRAV6_export_result.xlsx')

# ── Normal-CDF approximation constants (Abramowitz & Stegun, max error < 1.5e-7)
# Module-level so they are assigned once, not re-captured on every call.
_NC_A1 =  0.254829592
_NC_A2 = -0.284496736
_NC_A3 =  1.421413741
_NC_A4 = -1.453152027
_NC_A5 =  1.061405429
_NC_P  =  0.3275911

# ── Thermal probit model constants ────────────────────────────────────────────
# Probit = _THERM_PA + _THERM_PB * ln( (1000*kW)^_THERM_DOE * t_exp )
# P_fatality = Φ( (Probit - _THERM_PROBIT_MEAN) / sqrt(2) )
_THERM_PA          = -36.38
_THERM_PB          =   2.56
_THERM_DOE         =   4.0 / 3.0   # thermal dose exponent
_THERM_PROBIT_MEAN =   5.0          # probit centre (standard 50 % point)
_THERM_KW_SCALE    = 1000.0         # kW → W conversion inside probit argument

# ── Explosion step-function fatality thresholds ───────────────────────────────
# These are NOT in the General sheet (they come from the ImpactExpMatrix AA column
# of the original KernelV0).  They are kept here as named constants.
# _EXP_FATAL_BAR : overpressure (bar) at and above which fatality factor = 1
# _EXP_LOW_BAR   : overpressure (bar) below which fatality factor = 0
# NOTE: verify these values with the project team before accepting results.
_EXP_LOW_BAR    = 0.1   # limitOV1
_EXP_FATAL_BAR  = 0.3   # limitOV2
_EXP_F_MID      = 0.0   # fatality factor between _EXP_LOW_BAR and _EXP_FATAL_BAR
_EXP_F_FATAL    = 1.0   # fatality factor at or above _EXP_FATAL_BAR

# ── Sheet and size labels ─────────────────────────────────────────────────────
SHEET_IMPACT = 'Impact Matrix Result'
SHEET_RISK   = 'Risk Matrix Result'
SIZES        = ['Total', 'S', 'M', 'L', 'XL', 'INST']
_SIZE_NORM   = {sz.upper(): sz for sz in SIZES}   # case-insensitive lookup

# ── Impact event configuration ────────────────────────────────────────────────
# prob_idx : 0-based index into the Core row (col A = 0) for this event's frequency
IMPACT_CONFIG = {
    16: {'sheet': 'Outdoor Toxic Results',  'event': 'TOXIC', 'prob_idx': 9,  'formula': 'toxic'},
    17: {'sheet': 'Jet Fire Results',        'event': 'JF',    'prob_idx': 10, 'formula': 'jf_ellipse'},
    18: {'sheet': 'Late Pool Fire Results',  'event': 'LPF',   'prob_idx': 11, 'formula': 'thermal'},
    19: {'sheet': 'Early Pool Fire Results', 'event': 'EPF',   'prob_idx': 12, 'formula': 'thermal'},
    20: {'sheet': 'Fireball Results',        'event': 'FB',    'prob_idx': 13, 'formula': 'thermal'},
    21: {'sheet': 'CVE Results',             'event': 'CVE',   'prob_idx': 14, 'formula': 'explosion', 'cve': True},
    22: {'sheet': 'BLEVE Results',           'event': 'BLV',   'prob_idx': 15, 'formula': 'explosion'},
    23: {'sheet': 'Flash Fire Results',      'event': 'FF',    'prob_idx': 16, 'formula': 'ff'},
}

# ── PageControl layout constants ──────────────────────────────────────────────
_PC_READ_RANGE  = 'J2:R500'
_PC_EVENT_COL   = 0   # offset in the read range: column J → event description
_PC_ID_COL      = 2   # offset: column L → Impact ID integer
_PC_SIZE_COLS   = {'Total': 3, 'S': 4, 'M': 5, 'L': 6, 'XL': 7, 'INST': 8}

# ── General sheet cell addresses (1-based row, col) ───────────────────────────
# Row 5  → X-axis grid parameters
# Row 6  → Y-axis grid parameters
# Row 17 → Thermal radiation thresholds [kW/m²]
# Row 18 → Flammable concentration zone values (outside / transition / inside)
# Row 19 → Overpressure thresholds [bar]
# Row 20 → Toxic probability levels
# Row 21 → Thermal exposure duration [s]
_GEN_SX_RC       = (5, 19)
_GEN_QX_RC       = (5, 20)
_GEN_SY_RC       = (6, 19)
_GEN_QY_RC       = (6, 20)
_GEN_THERM_RC    = (17, 10)   # J17 start; up to 10 consecutive non-None values
_GEN_FF_RC       = (18, 10)   # J18 start; 3 values: outside / transition / inside LFL
_GEN_EXP_RC      = (19, 10)   # J19 start; up to 5 values
_GEN_TOX_RC      = (20, 10)   # J20 start; first value = min probability filter
_GEN_TEXP_RC     = (21,  4)   # D21: thermal exposure time [s]
_GEN_JF_INDEX_RC  = (22,  4)   # D22: JF direction index count
_GEN_JF_OFFSET_RC = (23,  4)   # D23: JF angle offset [deg]
_GEN_JF_DIRS_RC   = (24,  4)   # D24: JF number of directions
_GEN_JF_STEP_RC   = (25,  4)   # D25: JF angle step [deg] (0 = equal spacing)

# ── Module-level grid (populated by load_general_params) ─────────────────────
XX = None   # (QY, QX) array of cell-centre X coordinates [m]
YY = None   # (QY, QX) array of cell-centre Y coordinates [m]

# ── Module-level params dict (populated by load_general_params) ──────────────
_P = {}

# ── Warning / error collector ─────────────────────────────────────────────────
_warnings = []

def _warn(msg):
    _warnings.append(msg)
    print(f'  [WARN] {msg}')


# ══════════════════════════════════════════════════════════════════════════════
# Normal CDF
# ══════════════════════════════════════════════════════════════════════════════

def _norm_cdf(x):
    """Standard normal CDF — vectorized.
    Uses module-level _NC_* constants (Abramowitz & Stegun approximation)."""
    x   = np.asarray(x, dtype=float)
    sgn = np.sign(x)
    xa  = np.abs(x)
    t   = 1.0 / (1.0 + _NC_P * xa)
    y   = 1.0 - (
        ((((_NC_A5 * t + _NC_A4) * t + _NC_A3) * t + _NC_A2) * t + _NC_A1)
        * t * np.exp(-xa * xa)
    )
    return 0.5 * (1.0 + sgn * y)


# ══════════════════════════════════════════════════════════════════════════════
# General-sheet loader
# ══════════════════════════════════════════════════════════════════════════════

def load_general_params(wb):
    """Read all runtime parameters from General sheet.
    Builds global XX / YY grid arrays and populates module-level _P dict."""
    global XX, YY, _P

    ws = wb.sheets['General']

    def _fv(rc):
        v = ws.range(rc).value
        return float(v) if v is not None else None

    def _read_levels(rc, max_n):
        r, c0 = rc
        vals = []
        for i in range(max_n):
            v = ws.range((r, c0 + i)).value
            if v is None:
                break
            try:
                vals.append(float(v))
            except (TypeError, ValueError):
                break
        return np.array(vals, dtype=float)

    QX = int(_fv(_GEN_QX_RC))
    QY = int(_fv(_GEN_QY_RC))
    SX = _fv(_GEN_SX_RC)
    SY = _fv(_GEN_SY_RC)

    therm = _read_levels(_GEN_THERM_RC, 10)
    ff    = _read_levels(_GEN_FF_RC,     3)
    exp   = _read_levels(_GEN_EXP_RC,    5)
    tox   = _read_levels(_GEN_TOX_RC,    5)
    t_exp = _fv(_GEN_TEXP_RC)

    xc = SX * (np.arange(QX) + 0.5)
    yc = SY * (np.arange(QY) + 0.5)
    XX, YY = np.meshgrid(xc, yc)

    _P = {
        'QX':               QX,
        'QY':               QY,
        'SX':               SX,
        'SY':               SY,
        'THERM_THRESHOLDS': therm,
        'FF_OUTSIDE':       float(ff[0]) if len(ff) > 0 else 0.0,
        'FF_TRANSITION':    float(ff[1]) if len(ff) > 1 else 1.0,
        'FF_INSIDE_LFL':    float(ff[2]) if len(ff) > 2 else 2.0,
        'EXP_THRESHOLDS':   exp,
        'TOX_MIN_PROB':     float(tox[0]) if len(tox) > 0 else 0.01,
        'THERM_T_EXP':      float(t_exp)  if t_exp is not None else 20.0,
        'JF_THRESHOLDS':    np.array([1.6, 5.0, 7.3, 9.5, 12.5, 16.0, 20.9, 25.0, 30.0, 35.0]),
        'JF_DIRECTIONS':    int(_fv(_GEN_JF_DIRS_RC)  or 8),
        'JF_ANGLE_OFFSET':  float(_fv(_GEN_JF_OFFSET_RC) or 0.0),
        'JF_ANGLE_STEP':    float(_fv(_GEN_JF_STEP_RC)   or 0.0),
    }
    return _P


def print_params(p):
    """Print all loaded parameters for manual verification."""
    sep = '─' * 65
    print(f'\n{sep}')
    print('PARAMETERS  (read from General sheet — verify before accepting results)')
    print(sep)
    print(f'  Grid            : QX={p["QX"]}  QY={p["QY"]}')
    print(f'  Cell size       : SX={p["SX"]:.10f}  SY={p["SY"]:.10f}')
    print(f'  Thermal kW/m²   : {p["THERM_THRESHOLDS"].tolist()}')
    print(f'  FF zone values  : outside={p["FF_OUTSIDE"]}  '
          f'transition={p["FF_TRANSITION"]}  inside={p["FF_INSIDE_LFL"]}')
    print(f'  Overpressure bar: {p["EXP_THRESHOLDS"].tolist()}')
    print(f'  Exp fatal limit : >= {_EXP_FATAL_BAR} bar  '
          f'[NOT from General sheet — verify with project team]')
    print(f'  Tox min prob    : {p["TOX_MIN_PROB"]}')
    print(f'  Thermal t_exp   : {p["THERM_T_EXP"]} s')
    print(sep + '\n')


# ══════════════════════════════════════════════════════════════════════════════
# Formula implementations
# ══════════════════════════════════════════════════════════════════════════════

def dist_grid(sx, sy):
    """Euclidean distance [m] from point (sx, sy) to every grid cell → (QY, QX)."""
    return np.sqrt((XX - sx) ** 2 + (YY - sy) ** 2)


def formula_jf(sx, sy, dist_vals, halfW_vals, center_vals):
    """Jet Fire directional ellipse impact formula.

    For each of JF_DIRECTIONS equally-spaced jet angles, builds an ellipse for
    each kW/m² threshold.  Returns the highest threshold kW/m² for which the
    cell falls inside at least one directional ellipse.

    All spatial quantities are in cell units (distances from ImpactJFMatrix are
    stored in cell units, consistent with the Excel formula dividing by SX/SY).

    Returns (QY, QX) array of kW/m² values (0.0 or one of JF_THRESHOLDS).
    """
    SX = _P['SX']
    SY = _P['SY']
    thresholds  = _P['JF_THRESHOLDS']
    n_dirs      = _P['JF_DIRECTIONS']
    offset      = _P['JF_ANGLE_OFFSET']
    step        = _P['JF_ANGLE_STEP']

    x_rel = (XX - sx) / SX   # (QY, QX) cell-unit X offset from source
    y_rel = (YY - sy) / SY   # (QY, QX) cell-unit Y offset from source

    if step == 0:
        angles_deg = offset + np.arange(n_dirs) * 360.0 / n_dirs
    else:
        angles_deg = offset + np.arange(n_dirs) * step

    ct = np.cos(np.radians(angles_deg))   # (n_dirs,)
    st = np.sin(np.radians(angles_deg))   # (n_dirs,)

    x_exp = x_rel[..., np.newaxis]   # (QY, QX, 1)
    y_exp = y_rel[..., np.newaxis]

    result = np.zeros((int(_P['QY']), int(_P['QX'])))

    for i in range(len(thresholds)):
        dist   = dist_vals[i]
        halfW  = halfW_vals[i]
        center = center_vals[i]

        if dist is None or halfW is None or center is None:
            continue
        if dist <= 0 or halfW <= 0:
            continue

        a = dist - center
        b = halfW
        c = center

        if a <= 0:
            continue

        imp = thresholds[i]

        proj_along = x_exp * ct + y_exp * st - c   # (QY, QX, n_dirs)
        proj_perp  = x_exp * st - y_exp * ct       # (QY, QX, n_dirs)

        eVals = (proj_along / a) ** 2 + (proj_perp / b) ** 2

        inside_any = np.any(eVals <= 1.0, axis=2)   # (QY, QX) bool
        result = np.where(inside_any, np.maximum(result, imp), result)

    return result


def formula_thermal(dist, therm_dists):
    """Thermal radiation: interpolate kW/m², apply probit, return P_fatality.

    therm_dists : list aligned with _P['THERM_THRESHOLDS']; None where not computed.
    """
    thresholds = _P['THERM_THRESHOLDS']
    t_exp      = _P['THERM_T_EXP']
    valid      = [(i, d) for i, d in enumerate(therm_dists) if d is not None and d > 0]
    if not valid:
        return np.zeros(dist.shape)

    kw_v  = thresholds[[i for i, _ in valid]]
    d_v   = np.array([d for _, d in valid])

    kw_at = np.interp(dist, d_v[::-1], kw_v[::-1], left=kw_v[::-1][0], right=0.0)
    result = np.zeros(dist.shape)
    mask   = kw_at > 0.0
    if np.any(mask):
        kw      = kw_at[mask]
        certain = kw >= thresholds[-1]        # at or above max threshold → certain death
        probit  = (_THERM_PA
                   + _THERM_PB * np.log((_THERM_KW_SCALE * kw) ** _THERM_DOE * t_exp))
        p = np.clip(_norm_cdf(probit - _THERM_PROBIT_MEAN), 0.0, 1.0)
        p[certain] = 1.0
        result[mask] = p
    return result


def formula_explosion(dist, exp_dists):
    """Overpressure: interpolate bar, apply step-function fatality factor.

    exp_dists : list aligned with _P['EXP_THRESHOLDS']; None where not computed.
    Returns 0.0 or _EXP_F_FATAL per cell.
    """
    thresholds = _P['EXP_THRESHOLDS']
    valid      = [(i, d) for i, d in enumerate(exp_dists) if d is not None and d > 0]
    if not valid:
        return np.zeros(dist.shape)

    bar_v  = thresholds[[i for i, _ in valid]]
    d_v    = np.array([d for _, d in valid])
    bar_at = np.interp(dist, d_v[::-1], bar_v[::-1], left=bar_v[::-1][0], right=0.0)

    return np.where(
        bar_at >= _EXP_FATAL_BAR, _EXP_F_FATAL,
        np.where(bar_at >= _EXP_LOW_BAR, _EXP_F_MID, 0.0)
    )


def formula_ff(dist, lfl_dist, lflf_dist):
    """Flash fire: step function on LFL and LFL-fraction radii.

    Returns matrix with values FF_INSIDE_LFL / FF_TRANSITION / FF_OUTSIDE.
    """
    outside    = _P['FF_OUTSIDE']
    transition = _P['FF_TRANSITION']
    inside     = _P['FF_INSIDE_LFL']
    result = np.full(dist.shape, outside)
    result[dist <= lflf_dist] = transition
    result[dist <= lfl_dist]  = inside
    return result


def formula_toxic(dist, tox_dists, tox_probs):
    """Toxic: interpolate fatality probability from distance-probability profile."""
    if len(tox_dists) == 0:
        return np.zeros(dist.shape)
    return np.clip(
        np.interp(dist, tox_dists, tox_probs, left=tox_probs[0], right=0.0),
        0.0, 1.0
    )


# ══════════════════════════════════════════════════════════════════════════════
# PageControl readers
# ══════════════════════════════════════════════════════════════════════════════

def _as_text(v):
    return str(v).strip() if v is not None else ''


def _looks_like_range(v):
    return bool(re.match(r'^[\$A-Za-z]+[\$\d]+:[\$A-Za-z]+[\$\d]+$',
                         _as_text(v).replace('$', '')))


def _row_ranges(row):
    return {sz: _as_text(row[col])
            for sz, col in _PC_SIZE_COLS.items()
            if _looks_like_range(row[col])}


def read_page_control(ws_pc):
    """Return {impact_id: True} for all active rows in PageControl."""
    active = {}
    data   = ws_pc.range(_PC_READ_RANGE).value
    if not data:
        _warn('PageControl appears empty — using all IMPACT_CONFIG entries.')
        return {iid: True for iid in IMPACT_CONFIG}
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
        active[iid] = True
        name = _as_text(row[_PC_EVENT_COL])
        if name:
            print(f'    Impact ID {iid}: {name}')
    if not active:
        _warn('No active impacts in PageControl — using all IMPACT_CONFIG entries.')
        return {iid: True for iid in IMPACT_CONFIG}
    return active


def read_directions(ws_pc):
    """Return {impact_id: {size: range_str}} from PageControl."""
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
    """'$LD$318:$XF$634' → (start_row, start_col) 1-based."""
    clean = range_str.replace('$', '').split(':')[0]
    m = re.match(r'([A-Za-z]+)(\d+)', clean)
    if not m:
        raise ValueError(f'Cannot parse range: {range_str!r}')
    col = 0
    for ch in m.group(1).upper():
        col = col * 26 + (ord(ch) - ord('A') + 1)
    return int(m.group(2)), col


# ══════════════════════════════════════════════════════════════════════════════
# Core scenarios reader
# ══════════════════════════════════════════════════════════════════════════════

_CORE_PROB_COLS = {
    9:  'prob',  10: 'p_jf',  11: 'p_lpf', 12: 'p_epf',
    13: 'p_fb',  14: 'p_cve', 15: 'p_blv', 16: 'p_ff',
}


def read_core_scenarios(ws_core):
    """Read active scenarios from Core sheet (A4:R89).

    Row sentinel values:
      - None in col A  → end of data (stop reading)
      - 0 in col A     → end of data
      - -1 in col A    → no results for current Impact ID (stop reading, not an error)
    """
    data = ws_core.range('A4:R89').value
    scenarios = []
    if not data:
        return scenarios

    def _f(v): return float(v) if isinstance(v, (int, float)) else 0.0
    def _s(v): return str(v)   if v is not None else ''

    for row in data:
        if not row:
            break
        key = row[0]
        if key is None or key == 0 or key == -1:
            break
        scenarios.append({
            'key':         _s(key),
            'size':        _s(row[6]),   # col G (index 6); col C (index 2) = Dia/Noche
            'x':           _f(row[3]),
            'y':           _f(row[4]),
            'prob':        _f(row[9]),
            'p_jf':        _f(row[10]),
            'p_lpf':       _f(row[11]),
            'p_epf':       _f(row[12]),
            'p_fb':        _f(row[13]),
            'p_cve':       _f(row[14]),
            'p_blv':       _f(row[15]),
            'p_ff':        _f(row[16]),
        })
    return scenarios


# ══════════════════════════════════════════════════════════════════════════════
# Dynamic column-matching helper
# ══════════════════════════════════════════════════════════════════════════════

def _match_threshold_cols(headers, thresholds, sheet_name=''):
    """Match row-1 numeric headers to threshold values (tolerance 1e-4).

    Returns dict: {threshold_value → 1-based column index in the sheet}.
    If a threshold has no matching header, it is absent from the dict.
    """
    matches   = {}
    tol       = 1e-4
    unmatched = []
    for threshold in thresholds:
        found = False
        for col_0based, h in enumerate(headers):
            if h is None:
                continue
            try:
                h_val = float(h)
            except (TypeError, ValueError):
                continue
            if abs(h_val - threshold) < tol:
                matches[threshold] = col_0based + 1   # convert to 1-based
                found = True
                break
        if not found:
            unmatched.append(threshold)

    if unmatched:
        _warn(f'{sheet_name}: threshold columns not found in row 1: {unmatched}')
    return matches


# ══════════════════════════════════════════════════════════════════════════════
# Result-sheet readers
# ══════════════════════════════════════════════════════════════════════════════

def read_jf_results(ws, max_rows=1200):
    """Read JF ellipse data from ImpactJFMatrix-style sheet.

    Expected column layout (mirrors ImpactJFMatrix in KernelV0):
      Col B: path key
      Cols F:O  (6-15):  far-tip distances for 10 kW/m² thresholds (cell units)
      Cols P:Y  (16-25): semi-minor axis (half-width) for 10 thresholds (cell units)
      Cols Z:AI (26-35): center distances for 10 thresholds (cell units)

    Returns: {path_key: {'dist': [...], 'halfW': [...], 'center': [...]}}
    """
    def _flt(v):
        return float(v) if isinstance(v, (int, float)) else None

    results = {}
    for r in range(2, max_rows):
        path = ws.range((r, 2)).value
        if not path:
            break
        row_d = ws.range((r, 6),  (r, 15)).value or []
        row_w = ws.range((r, 16), (r, 25)).value or []
        row_c = ws.range((r, 26), (r, 35)).value or []
        results[str(path)] = {
            'dist':   [_flt(v) for v in row_d],
            'halfW':  [_flt(v) for v in row_w],
            'center': [_flt(v) for v in row_c],
        }
    return results


def read_thermal_results(ws, max_rows=1200):
    """Read thermal radiation distance columns, matched dynamically to thresholds.

    Returns: {path_key: [dist_or_None, ...]} — list aligned with THERM_THRESHOLDS.
    """
    thresholds = _P['THERM_THRESHOLDS']
    headers    = ws.range('A1:AI1').value          # wide enough for all thermal cols
    col_map    = _match_threshold_cols(headers, thresholds, ws.name)

    if not col_map:
        _warn(f'{ws.name}: no threshold columns matched — sheet will produce zero matrices.')
        return {}

    found = sorted(col_map.keys())
    print(f'  {ws.name}: matched thresholds {found}')

    results = {}
    for r in range(2, max_rows):
        path = ws.range((r, 2)).value
        if not path:
            break
        dists = []
        for t in thresholds:
            col = col_map.get(t)
            if col is None:
                dists.append(None)
            else:
                v = ws.range((r, col)).value
                dists.append(float(v) if isinstance(v, (int, float)) else None)
        results[str(path)] = dists
    return results


def read_ff_results(ws, max_rows=1200):
    """Read LFL and LFL-fraction distances from Flash Fire results.

    Col E = Distance downwind to LFL [m]
    Col F = Distance downwind to LFL Fraction [m]
    Returns: {path_key: (lfl_dist, lflf_dist)}.
    """
    results = {}
    for r in range(2, max_rows):
        path = ws.range((r, 2)).value
        if not path:
            break
        lfl  = ws.range((r, 5)).value
        lflf = ws.range((r, 6)).value
        if isinstance(lfl, (int, float)):
            lflf_val = float(lflf) if isinstance(lflf, (int, float)) else float(lfl)
            results[str(path)] = (float(lfl), lflf_val)
    return results


def read_explosion_results(ws, is_cve=False, max_rows=200):
    """Read overpressure distance columns from CVE or BLEVE result sheet.

    For CVE (is_cve=True) each row also carries its own X/Y ignition coordinates
    from cols V (col 22) and W (col 23), header 'X' / 'Y'.

    Returns: {path_key: list_of_entries}
      Each entry: {'dists': [...], 'x': float|None, 'y': float|None}
      A key can map to MULTIPLE entries (one per ignition location for CVE).

    Column matching strategy:
    - Read row-1 headers; match to EXP_THRESHOLDS.
    - If matching fails (e.g. formula error returning 1/-1), fall back to
      POSITIONAL assignment (J=col10, K=col11, …) and log a warning.
    """
    thresholds = _P['EXP_THRESHOLDS']
    n          = len(thresholds)
    headers    = ws.range('A1:Z1').value

    # Check whether col2 starts with 'No Data'
    if _as_text(headers[1]).lower().startswith('no data'):
        _warn(f'{ws.name}: sheet reports "No Data" — skipping.')
        return {}

    col_map  = _match_threshold_cols(headers, thresholds, ws.name)
    fallback = False

    if len(col_map) < n:
        _warn(
            f'{ws.name}: only {len(col_map)}/{n} overpressure threshold columns matched '
            f'(headers={[h for h in headers if h is not None]}). '
            f'Falling back to positional assignment: J–{"JKLMN"[n-1]} = thresholds.'
        )
        # Positional fallback: J=col10, K=col11, …
        col_map  = {t: 10 + i for i, t in enumerate(thresholds)}
        fallback = True

    if not fallback:
        print(f'  {ws.name}: matched overpressure thresholds {sorted(col_map.keys())}')

    results = {}
    for r in range(2, max_rows):
        path = ws.range((r, 2)).value
        if not path:
            break
        dists = []
        for t in thresholds:
            col = col_map.get(t)
            v   = ws.range((r, col)).value if col is not None else None
            dists.append(float(v) if isinstance(v, (int, float)) else None)

        x_coord = y_coord = None
        if is_cve:
            xv = ws.range((r, 22)).value   # col V header 'X'
            yv = ws.range((r, 23)).value   # col W header 'Y'
            x_coord = float(xv) if isinstance(xv, (int, float)) else None
            y_coord = float(yv) if isinstance(yv, (int, float)) else None

        entry = {'dists': dists, 'x': x_coord, 'y': y_coord}
        key   = str(path)
        if key not in results:
            results[key] = []
        results[key].append(entry)

    return results


def read_toxic_results(ws, max_rows=100):
    """Read the compressed CSV blob in col G and expand to (distances, probs) arrays.

    Col G contains one multi-line CSV text cell per scenario.
    The engine parses it, filters rows where distance>=0 and probability >= TOX_MIN_PROB,
    sorts by ascending distance, and returns arrays ready for np.interp.

    Returns: {path_key: (tox_dists_array, tox_probs_array)}.
    """
    min_prob = _P['TOX_MIN_PROB']
    results  = {}

    for r in range(2, max_rows):
        path = ws.range((r, 2)).value
        if not path:
            break
        blob = ws.range((r, 7)).value
        if not blob:
            continue
        pairs = []
        for line in str(blob).strip().split('\n'):
            parts = line.strip().split(',')
            if len(parts) < 4:
                continue
            try:
                d = float(parts[0])
                p = float(parts[3])
                if d >= 0 and p >= min_prob:
                    pairs.append((d, p))
            except ValueError:
                continue
        if not pairs:
            continue
        pairs.sort(key=lambda t: t[0])
        results[str(path)] = (
            np.array([t[0] for t in pairs]),
            np.array([t[1] for t in pairs]),
        )
    return results


# ══════════════════════════════════════════════════════════════════════════════
# Core grid computation
# ══════════════════════════════════════════════════════════════════════════════

def compute_event(formula_type, prob_idx, scenarios, results_data, is_cve=False):
    """Accumulate Impact and Risk matrices for one event across all size filters.

    For CVE (is_cve=True) each result-data entry is a list (one item per
    ignition location); the engine iterates over each entry separately.
    For all other formulas each entry is a list of length 1.

    Returns:
        {size_label: (impact_matrix (QY,QX), risk_matrix (QY,QX))}
    """
    QY      = _P['QY']
    QX      = _P['QX']
    prob_col = _CORE_PROB_COLS.get(prob_idx, 'prob')

    impact_mats = {sz: np.zeros((QY, QX)) for sz in SIZES}
    risk_mats   = {sz: np.zeros((QY, QX)) for sz in SIZES}
    matched = unmatched = 0

    for sc in scenarios:
        key  = sc['key']
        prob = sc.get(prob_col, 0.0)
        size = sc['size']

        if key not in results_data:
            unmatched += 1
            continue
        matched += 1

        entries = results_data[key]   # list of dicts for CVE; list-of-1 for others

        for entry in entries:
            if is_cve:
                sx = entry.get('x')
                sy = entry.get('y')
                if sx is None or sy is None:
                    continue
            else:
                sx = sc['x']
                sy = sc['y']

            d = dist_grid(sx, sy)

            if formula_type == 'thermal':
                cell_imp = formula_thermal(d, entry['dists'])
            elif formula_type == 'explosion':
                cell_imp = formula_explosion(d, entry['dists'])
            elif formula_type == 'ff':
                lfl_dist, lflf_dist = entry['dists']
                cell_imp = formula_ff(d, lfl_dist, lflf_dist)
            elif formula_type == 'toxic':
                tox_d, tox_p = entry['dists']
                cell_imp = formula_toxic(d, tox_d, tox_p)
            elif formula_type == 'jf_ellipse':
                cell_imp = formula_jf(sx, sy, entry['dist'], entry['halfW'], entry['center'])
            else:
                continue

            risk_c = cell_imp * prob

            impact_mats['Total'] += cell_imp
            risk_mats['Total']   += risk_c
            sz_key = _SIZE_NORM.get(str(size).strip().upper())
            if sz_key and sz_key != 'Total':
                impact_mats[sz_key] += cell_imp
                risk_mats[sz_key]   += risk_c

    return {sz: (impact_mats[sz], risk_mats[sz]) for sz in SIZES}, matched, unmatched


# ══════════════════════════════════════════════════════════════════════════════
# CSV export
# ══════════════════════════════════════════════════════════════════════════════

def _sum_final_matrices(all_results):
    """Sum Total impact and risk matrices across all events → (QY×QX, QY×QX)."""
    QY = _P['QY']
    QX = _P['QX']
    imp  = np.zeros((QY, QX))
    risk = np.zeros((QY, QX))
    for res in all_results.values():
        imp  += res['impact']['Total']
        risk += res['risk']['Total']
    return imp, risk


def export_matrices_csv(all_results, output_dir, directions=None):
    """Build two composite CSVs that mirror the full PageControl sheet layout.

    Each event×size matrix is placed at its PageControl destination (row, col).
    Pasting impact_sheet.csv at A1 of 'Impact Matrix Result' and
    risk_sheet.csv at A1 of 'Risk Matrix Result' produces the correct layout.
    """
    QY = _P['QY']
    QX = _P['QX']
    os.makedirs(output_dir, exist_ok=True)

    if not directions:
        _warn('export_matrices_csv: directions empty — writing only summed Total CSVs.')
        imp_final, risk_final = _sum_final_matrices(all_results)
        np.savetxt(os.path.join(output_dir, 'impact_sheet.csv'),  imp_final,  delimiter=',', fmt='%.6e')
        np.savetxt(os.path.join(output_dir, 'risk_sheet.csv'),    risk_final, delimiter=',', fmt='%.6e')
        print(f'  impact_sheet.csv | risk_sheet.csv  (summed Total, {QY}x{QX})')
        return

    # Determine full sheet extent from PageControl destinations
    max_row = 0
    max_col = 0
    blocks  = []   # (impact_id, sz, sr, sc)

    for impact_id, res in all_results.items():
        dir_map = directions.get(impact_id, {})
        for sz in SIZES:
            rng_str = dir_map.get(sz)
            if not rng_str:
                continue
            try:
                sr, sc = _parse_range_start(rng_str)
            except ValueError:
                continue
            max_row = max(max_row, sr + QY - 1)
            max_col = max(max_col, sc + QX - 1)
            blocks.append((impact_id, sz, sr, sc))

    if not blocks:
        _warn('export_matrices_csv: no valid destinations parsed — skipping composite export.')
        return

    print(f'  Building composite arrays ({max_row}x{max_col}) for {len(blocks)} blocks …')
    imp_sheet  = np.zeros((max_row, max_col))
    risk_sheet = np.zeros((max_row, max_col))

    for impact_id, sz, sr, sc in blocks:
        r0 = sr - 1   # Excel 1-based → numpy 0-based
        c0 = sc - 1
        imp_sheet[r0:r0+QY, c0:c0+QX]  = all_results[impact_id]['impact'][sz]
        risk_sheet[r0:r0+QY, c0:c0+QX] = all_results[impact_id]['risk'][sz]

    imp_path  = os.path.join(output_dir, 'impact_sheet.csv')
    risk_path = os.path.join(output_dir, 'risk_sheet.csv')
    np.savetxt(imp_path,  imp_sheet,  delimiter=',', fmt='%.6e')
    np.savetxt(risk_path, risk_sheet, delimiter=',', fmt='%.6e')
    print(f'  impact_sheet.csv | risk_sheet.csv  ({max_row}x{max_col}, {len(blocks)} blocks placed)')
    print(f'  → paste each at A1 of the respective result sheet')


# ══════════════════════════════════════════════════════════════════════════════
# Excel result writers
# ══════════════════════════════════════════════════════════════════════════════

def _bulk_write(ws, start_row, start_col, matrix):
    ws.range((start_row, start_col)).value = np.nan_to_num(matrix).tolist()


def _ensure_sheet(wb, name):
    sheet_names = [s.name for s in wb.sheets]
    if name in sheet_names:
        ws = wb.sheets[name]
        ws.clear()
        return ws
    wb.sheets.add(name, after=wb.sheets[-1])
    return wb.sheets[name]


def write_result_sheets(wb, all_results, directions):
    ws_imp  = _ensure_sheet(wb, SHEET_IMPACT)
    ws_risk = _ensure_sheet(wb, SHEET_RISK)

    app = wb.app
    app.api.ScreenUpdating = False
    app.api.Calculation    = -4135   # xlCalculationManual
    app.api.EnableEvents   = False

    written = 0
    try:
        for impact_id, res in all_results.items():
            dir_map   = directions.get(impact_id, {})
            imp_mats  = res['impact']
            risk_mats = res['risk']

            for sz in SIZES:
                rng_str = dir_map.get(sz)
                if not rng_str:
                    continue
                try:
                    sr, sc = _parse_range_start(rng_str)
                except ValueError as e:
                    _warn(str(e))
                    continue

                _bulk_write(ws_imp,  sr, sc, imp_mats[sz])
                _bulk_write(ws_risk, sr, sc, risk_mats[sz])
                written += 1
                print(f'    [{res["event"]}/{sz}] → row {sr}, col {sc}')

    finally:
        app.api.ScreenUpdating = True
        app.api.Calculation    = -4105   # xlCalculationAutomatic
        app.api.EnableEvents   = True

    if not written:
        _warn('No blocks written — directions map may be empty or all_results is empty.')
    print(f'  {written} blocks written.')


# ══════════════════════════════════════════════════════════════════════════════
# Warning / error summary
# ══════════════════════════════════════════════════════════════════════════════

def print_error_summary():
    sep = '═' * 65
    print(f'\n{sep}')
    if _warnings:
        print(f'WARNING / ERROR SUMMARY  ({len(_warnings)} items)')
        print(sep)
        for i, msg in enumerate(_warnings, 1):
            print(f'  {i:>3}. {msg}')
    else:
        print('No warnings or errors.')
    print(sep)


# ══════════════════════════════════════════════════════════════════════════════
# Main
# ══════════════════════════════════════════════════════════════════════════════

def main():
    t0 = time.time()
    print('=' * 65)
    print('QRA Engine V6 — Dynamic Parameter Reading')
    print('=' * 65)
    print(f'Source : {V6_PATH}')
    print(f'Export : {EXPORT_PATH}')

    if not os.path.exists(V6_PATH):
        print(f'\nERROR: Workbook not found: {V6_PATH}')
        sys.exit(1)

    print('\nOpening Excel workbook …')
    app = xw.App(visible=False)
    app.display_alerts  = False
    app.screen_updating = False

    try:
        wb          = app.books.open(V6_PATH)
        ws_core     = wb.sheets['Core']
        sheet_names = [s.name for s in wb.sheets]
        print(f'  Workbook opened.  Sheets: {len(sheet_names)}')

        # ── Load and verify parameters ────────────────────────────────────────
        print('\nReading General sheet parameters …')
        params = load_general_params(wb)
        print_params(params)

        # ── PageControl ───────────────────────────────────────────────────────
        if 'PageControl' not in sheet_names:
            _warn('PageControl sheet missing — using all IMPACT_CONFIG entries.')
            active_impacts = {iid: True for iid in IMPACT_CONFIG}
            directions     = {}
        else:
            print('  Reading PageControl …')
            ws_pc          = wb.sheets['PageControl']
            active_impacts = read_page_control(ws_pc)
            directions     = read_directions(ws_pc)

        # ── Iterate impact IDs ────────────────────────────────────────────────
        all_results = {}

        for impact_id in sorted(active_impacts.keys()):
            if impact_id not in IMPACT_CONFIG:
                _warn(f'Impact ID {impact_id} not in IMPACT_CONFIG — skipped.')
                continue

            cfg          = IMPACT_CONFIG[impact_id]
            event_name   = cfg['event']
            sheet_name   = cfg['sheet']
            prob_idx     = cfg['prob_idx']
            formula_type = cfg['formula']
            is_cve       = cfg.get('cve', False)

            print(f'\n{"─"*65}')
            print(f'  Impact {impact_id} | {event_name} | {sheet_name} | {formula_type}')
            print(f'{"─"*65}')

            # Set Impact ID in Core and recalculate
            ws_core.range('C2').value = impact_id
            time.sleep(1.5)
            app.api.Calculate()
            time.sleep(1.0)
            print(f'  Core!C2 = {impact_id} — recalculated.')

            # Warn if Core!G2 has a size filter (can hide scenarios)
            size_filter = ws_core.range('G2').value
            if size_filter and str(size_filter).strip():
                _warn(f'Impact {impact_id}: Core!G2 = "{size_filter}" '
                      f'— size filter is active; some scenarios may be hidden.')

            # Read Core scenarios
            scenarios = read_core_scenarios(ws_core)
            print(f'  Scenarios from Core: {len(scenarios)}')
            if scenarios:
                distinct_sizes = sorted({sc['size'] for sc in scenarios})
                print(f'  Size labels in Core col C: {distinct_sizes}')
            if not scenarios:
                _warn(f'Impact {impact_id}: no scenarios returned by Core (sentinel -1 or empty).')
                continue

            # Read effect results sheet
            if sheet_name not in sheet_names:
                _warn(f'Impact {impact_id}: sheet "{sheet_name}" not found — skipping.')
                continue

            ws_res  = wb.sheets[sheet_name]
            t_read  = time.time()

            if formula_type == 'thermal':
                results_data = read_thermal_results(ws_res)
                # Wrap entries to uniform list-of-dicts format expected by compute_event
                results_data = {
                    k: [{'dists': v, 'x': None, 'y': None}]
                    for k, v in results_data.items()
                }

            elif formula_type == 'ff':
                raw = read_ff_results(ws_res)
                results_data = {
                    k: [{'dists': (lfl, lflf), 'x': None, 'y': None}]
                    for k, (lfl, lflf) in raw.items()
                }

            elif formula_type == 'explosion':
                results_data = read_explosion_results(ws_res, is_cve=is_cve)

            elif formula_type == 'toxic':
                raw = read_toxic_results(ws_res)
                results_data = {
                    k: [{'dists': (td, tp), 'x': None, 'y': None}]
                    for k, (td, tp) in raw.items()
                }

            else:
                results_data = {}

            print(f'  Results loaded: {len(results_data)} keys  ({time.time()-t_read:.1f}s)')

            if not results_data:
                _warn(f'Impact {impact_id}: results sheet "{sheet_name}" produced no usable data.')
                continue

            # Compute matrices
            t_calc = time.time()
            size_results, matched, unmatched = compute_event(
                formula_type, prob_idx, scenarios, results_data, is_cve=is_cve
            )
            print(f'  Grid computed: {time.time()-t_calc:.1f}s  '
                  f'matched={matched}  unmatched={unmatched}')

            if unmatched > 0:
                _warn(f'Impact {impact_id}: {unmatched} Core scenarios had no matching '
                      f'row in "{sheet_name}".')

            for sz in ['Total', 'S', 'M']:
                imp, risk = size_results[sz]
                print(f'    [{sz:5s}]  impact_max={imp.max():.4f}  '
                      f'risk_max={risk.max():.4e}  nonzero={np.count_nonzero(imp)}')

            all_results[impact_id] = {
                'impact': {sz: size_results[sz][0] for sz in SIZES},
                'risk':   {sz: size_results[sz][1] for sz in SIZES},
                'event':  event_name,
            }

        # ── Export CSVs ───────────────────────────────────────────────────────
        if all_results:
            csv_dir = os.path.join(_WORKSPACE, 'output', 'matrices')
            export_matrices_csv(all_results, csv_dir, directions=directions)

        # ── Write result sheets ───────────────────────────────────────────────
        if all_results:
            write_result_sheets(wb, all_results, directions)
        else:
            _warn('No results computed — result sheets not written.')

        # ── Save export workbook ──────────────────────────────────────────────
        print(f'\nSaving export workbook …\n  → {EXPORT_PATH}')
        wb.api.SaveAs(EXPORT_PATH, FileFormat=51)   # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.close()
        print('  Saved OK.')

    except Exception as exc:
        import traceback
        print(f'\nFATAL ERROR: {exc}')
        traceback.print_exc()
    finally:
        try:
            app.quit()
        except Exception:
            pass

    print_error_summary()
    elapsed = time.time() - t0
    print(f'\nTotal elapsed: {elapsed:.1f}s  ({elapsed/60:.1f} min)')
    print('=' * 65)


if __name__ == '__main__':
    main()
