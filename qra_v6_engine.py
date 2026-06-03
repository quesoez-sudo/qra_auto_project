"""
QRA Engine V6 — Basic Backline Export
======================================
Standalone script for running on a target machine that has Excel + Python.
 
Pipeline
--------
1. Open MacroQRAV6 (version 1).xlsm via xlwings
2. Read PageControl  → determine which Impact IDs are active
3. For each active Impact ID:
     a. Set Core!C2 = impact_id  →  force recalculation
     b. Read Core scenarios (A4:R89)
     c. Read Directions  →  get destination ranges per size (Total/S/M/L/XL/INST)
     d. Read the corresponding effect results sheet
     e. Compute impact & risk matrices on the QY×QX grid
4. Create (or overwrite) two sheets in the workbook:
        "Impact Matrix Result"
        "Risk Matrix Result"
5. Write each matrix at the destination range found in Directions
6. Save workbook as macro-disabled .xlsx  (no macro = safe for export)
 
PageControl sheet assumptions
------------------------------
  Row 1  = header
  Col A  = Impact ID   (integer, e.g. 16, 17 … 23)
  Col B  = Event name  (e.g. TOXIC, JF …)         [informational]
  Col C  = Active flag (1 = active, 0 / blank = skip)
 
Directions sheet assumptions
-----------------------------
  Row 1  = header
  Col A  = Impact ID   (integer)
  Col B  = Size label  (Total / S / M / L / XL / INST)
  Col C  = Destination range string  (e.g. '$LD$318:$XF$634')
             used to position the matrix in the result sheets
"""
 
import os
import re
import sys
import time
 
import numpy as np
import xlwings as xw
 
sys.stdout.reconfigure(encoding='utf-8')
 
# ── Paths ─────────────────────────────────────────────────────────────────────
# Change these two lines for the target machine
_WORKSPACE   = r'C:\Users\herra\OneDrive - Chevron\python\QRA_tool'
V6_PATH      = os.path.join(_WORKSPACE, 'MacroQRAV6 (version 1).xlsm')
EXPORT_PATH  = os.path.join(_WORKSPACE, 'MacroQRAV6_export_result.xlsx')  # output file
 
# ── Grid ──────────────────────────────────────────────────────────────────────
QX = 315
QY = 317
SX = 1.0698412698412698
SY = 1.069425
 
_xc = SX * (np.arange(QX) + 0.5)
_yc = SY * (np.arange(QY) + 0.5)
XX, YY = np.meshgrid(_xc, _yc)
 
SIZES = ['Total', 'S', 'M', 'L', 'XL', 'INST']
 
# ── Formula constants ─────────────────────────────────────────────────────────
THERM_THRESHOLDS  = np.array([1.6, 5.0, 7.3, 9.5, 12.5, 16.0, 20.9, 25.0, 30.0, 35.0])
THERM_T_EXP       = 20.0
EXP_THRESHOLDS    = np.array([0.04, 0.1, 0.35, 0.5, 1.0])
EXP_LIM_OV1       = 0.1
EXP_LIM_OV2       = 0.3
EXP_LIM_F1        = 0.0
EXP_LIM_F2        = 1.0
FF_OUTSIDE        = 0.0
FF_TRANSITION     = 1.0
FF_INSIDE_LFL     = 2.0
TOX_MIN_PROB      = 0.01
 
# ── Static Impact Config (fallback if PageControl is absent / empty) ──────────
# Maps Impact ID -> effect results sheet name, prob column index, formula type
IMPACT_CONFIG = {
    16: {'sheet': 'Outdoor Toxic Results',  'event': 'TOXIC', 'prob_idx': 9,  'formula': 'toxic'},
    17: {'sheet': 'Jet Fire Results',        'event': 'JF',    'prob_idx': 10, 'formula': 'thermal'},
    18: {'sheet': 'Late Pool Fire Results',  'event': 'LPF',   'prob_idx': 11, 'formula': 'thermal'},
    19: {'sheet': 'Early Pool Fire Results', 'event': 'EPF',   'prob_idx': 12, 'formula': 'thermal'},
    20: {'sheet': 'Fireball Results',        'event': 'FB',    'prob_idx': 13, 'formula': 'thermal'},
    21: {'sheet': 'CVE Results',             'event': 'CVE',   'prob_idx': 14, 'formula': 'explosion'},
    22: {'sheet': 'BLEVE Results',           'event': 'BLV',   'prob_idx': 15, 'formula': 'explosion'},
    23: {'sheet': 'Flash Fire Results',      'event': 'FF',    'prob_idx': 16, 'formula': 'ff'},
}
 
# ── Result sheet names ────────────────────────────────────────────────────────
SHEET_IMPACT = 'Impact Matrix Result'
SHEET_RISK   = 'Risk Matrix Result'
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# Formula helpers  (identical to qra_engine_v6.py)
# ═══════════════════════════════════════════════════════════════════════════════
 
def _norm_cdf(x):
    a1, a2, a3, a4, a5, p = (
        0.254829592, -0.284496736, 1.421413741,
        -1.453152027, 1.061405429, 0.3275911)
    x    = np.asarray(x, dtype=float)
    sign = np.sign(x)
    xa   = np.abs(x)
    t    = 1.0 / (1.0 + p * xa)
    y    = 1.0 - (((((a5*t + a4)*t + a3)*t + a2)*t + a1)*t * np.exp(-xa*xa))
    return 0.5 * (1.0 + sign * y)
 
def dist_grid(sx, sy):
    return np.sqrt((XX - sx)**2 + (YY - sy)**2)
 
def formula_thermal(dist, therm_dists):
    valid = [i for i, d in enumerate(therm_dists) if d is not None and d > 0]
    if not valid:
        return np.zeros(dist.shape)
    kw_v  = THERM_THRESHOLDS[valid]
    d_v   = np.array([therm_dists[i] for i in valid])
    kw_at = np.interp(dist, d_v[::-1], kw_v[::-1], left=kw_v[::-1][0], right=0.0)
    result = np.zeros(dist.shape)
    mask   = kw_at > 0.0
    if np.any(mask):
        kw     = kw_at[mask]
        probit = -36.38 + 2.56 * np.log((1000.0 * kw)**(4.0/3.0) * THERM_T_EXP)
        result[mask] = np.clip(_norm_cdf(probit - 5.0), 0.0, 1.0)
    return result
 
def formula_explosion(dist, exp_dists):
    valid = [i for i, d in enumerate(exp_dists) if d is not None and d > 0]
    if not valid:
        return np.zeros(dist.shape)
    bar_v  = EXP_THRESHOLDS[valid]
    d_v    = np.array([exp_dists[i] for i in valid])
    bar_at = np.interp(dist, d_v[::-1], bar_v[::-1], left=bar_v[::-1][0], right=0.0)
    return np.where(bar_at >= EXP_LIM_OV2, EXP_LIM_F2,
           np.where(bar_at >= EXP_LIM_OV1, EXP_LIM_F1, 0.0))
 
def formula_ff(dist, lfl_dist, lflf_dist):
    result = np.full(dist.shape, FF_OUTSIDE)
    result[dist <= lflf_dist] = FF_TRANSITION
    result[dist <= lfl_dist]  = FF_INSIDE_LFL
    return result
 
def formula_toxic(dist, tox_dists, tox_probs):
    if len(tox_dists) == 0:
        return np.zeros(dist.shape)
    return np.clip(np.interp(dist, tox_dists, tox_probs,
                             left=tox_probs[0], right=0.0), 0.0, 1.0)
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# PageControl reader
# ═══════════════════════════════════════════════════════════════════════════════
 
def read_page_control(ws_pc):
    """
    Read PageControl sheet.
    Returns dict: {impact_id (int): True}  for all active impacts.
 
    Expected layout (row 1 = header):
      Col A = Impact ID
      Col B = Event name   (ignored here)
      Col C = Active flag  (1 = active)
    """
    active = {}
    print("  Reading PageControl...")
    data = ws_pc.range('A2:C100').value
    if not data:
        print("  WARNING: PageControl appears empty — will use all IMPACT_CONFIG entries.")
        return {iid: True for iid in IMPACT_CONFIG}
 
    for row in data:
        if not row or row[0] is None:
            break
        try:
            iid  = int(row[0])
            flag = row[2]
            if flag is None or str(flag).strip() == '':
                # Treat blank as active
                active[iid] = True
            elif float(flag) == 1.0:
                active[iid] = True
            else:
                print(f"    Impact ID {iid} flagged inactive — skipping.")
        except (ValueError, TypeError):
            continue
 
    if not active:
        print("  WARNING: No active impacts found in PageControl — using all.")
        return {iid: True for iid in IMPACT_CONFIG}
 
    print(f"  Active Impact IDs from PageControl: {sorted(active.keys())}")
    return active
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# Directions reader
# ═══════════════════════════════════════════════════════════════════════════════
 
def read_directions(ws_dir):
    """
    Read Directions sheet.
    Returns dict: {impact_id (int): {size (str): range_str (str)}}
 
    Expected layout (row 1 = header):
      Col A = Impact ID
      Col B = Size  (Total / S / M / L / XL / INST)
      Col C = Destination range  e.g. '$LD$318:$XF$634'
    """
    directions = {}
    print("  Reading Directions...")
    data = ws_dir.range('A2:C500').value
    if not data:
        print("  WARNING: Directions sheet appears empty.")
        return directions
 
    for row in data:
        if not row or row[0] is None:
            break
        try:
            iid  = int(row[0])
            sz   = str(row[1]).strip() if row[1] else ''
            rng  = str(row[2]).strip() if row[2] else ''
        except (ValueError, TypeError):
            continue
 
        if sz not in SIZES or not rng:
            continue
 
        directions.setdefault(iid, {})[sz] = rng
 
    total_entries = sum(len(v) for v in directions.values())
    print(f"  Directions loaded: {len(directions)} impacts, {total_entries} size-range entries.")
    return directions
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# Range string → (start_row, start_col) — 1-based
# ═══════════════════════════════════════════════════════════════════════════════
 
def _parse_range_start(range_str):
    """'$LD$318:$XF$634'  →  (318, col_number_of_LD)"""
    clean = range_str.replace('$', '').split(':')[0]
    m = re.match(r'([A-Za-z]+)(\d+)', clean)
    if not m:
        raise ValueError(f"Cannot parse range: {range_str!r}")
    letters, row_str = m.group(1).upper(), m.group(2)
    col = 0
    for ch in letters:
        col = col * 26 + (ord(ch) - ord('A') + 1)
    return int(row_str), col
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# Core scenarios reader  (identical to qra_engine_v6.py)
# ═══════════════════════════════════════════════════════════════════════════════
 
def read_core_scenarios(ws_core):
    data = ws_core.range('A4:R89').value
    scenarios = []
    if not data:
        return scenarios
    for row in data:
        if not row or row[0] is None or row[0] == 0:
            break
        def _f(v):  return float(v) if isinstance(v, (int, float)) else 0.0
        def _s(v):  return str(v)   if v is not None else ''
        scenarios.append({
            'key':           _s(row[0]),
            'scenario_code': _s(row[1]),
            'size':          _s(row[2]),
            'x':             _f(row[3]),
            'y':             _f(row[4]),
            'description':   _s(row[5]),
            'disch_range':   _s(row[6]),
            'base_freq':     _f(row[7]),
            'p_weather':     _f(row[8]),
            'prob':          _f(row[9]),   # P_TOXIC
            'p_jf':          _f(row[10]),
            'p_lpf':         _f(row[11]),
            'p_epf':         _f(row[12]),
            'p_fb':          _f(row[13]),
            'p_cve':         _f(row[14]),
            'p_blv':         _f(row[15]),
            'p_ff':          _f(row[16]),
            'filter':        _s(row[17]) if len(row) > 17 else '',
        })
    return scenarios
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# Effect results readers  (identical to qra_engine_v6.py)
# ═══════════════════════════════════════════════════════════════════════════════
 
def read_thermal_results(ws, max_rows=1200):
    results = {}
    for r in range(2, max_rows):
        path = ws.range((r, 2)).value
        if not path:
            break
        dists = []
        for c in range(6, 16):
            v = ws.range((r, c)).value
            dists.append(float(v) if isinstance(v, (int, float)) else None)
        results[str(path)] = dists
    return results
 
def read_ff_results(ws, max_rows=1200):
    results = {}
    for r in range(2, max_rows):
        path = ws.range((r, 2)).value
        if not path:
            break
        lfl_dist = ws.range((r, 5)).value
        lfl_frac = ws.range((r, 6)).value
        if isinstance(lfl_dist, (int, float)):
            lfl_f = float(lfl_frac) if isinstance(lfl_frac, (int, float)) else float(lfl_dist)
            results[str(path)] = (float(lfl_dist), lfl_f)
    return results
 
def read_explosion_results(ws, max_rows=200):
    results = {}
    for r in range(2, max_rows):
        path = ws.range((r, 2)).value
        if not path:
            break
        dists = []
        for c in range(10, 15):
            v = ws.range((r, c)).value
            dists.append(float(v) if isinstance(v, (int, float)) else None)
        exp_centre = ws.range((r, 9)).value
        ign_src    = ws.range((r, 7)).value
        results[str(path)] = {
            'dists':      dists,
            'exp_centre': float(exp_centre) if isinstance(exp_centre, (int, float)) else None,
            'ign_src':    float(ign_src)    if isinstance(ign_src,    (int, float)) else None,
        }
    return results
 
def read_toxic_results(ws, max_rows=100):
    results = {}
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
            if len(parts) >= 4:
                try:
                    d = float(parts[0]);  p = float(parts[3])
                    if d >= 0 and p >= TOX_MIN_PROB:
                        pairs.append((d, p))
                except ValueError:
                    pass
        if pairs:
            pairs.sort(key=lambda t: t[0])
            results[str(path)] = pairs
    return results
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# Size mapping helper
# ═══════════════════════════════════════════════════════════════════════════════
 
def map_disch_range_to_size(disch_range):
    dr = str(disch_range).strip().upper()
    return {'S':'S','M':'M','L':'L','XL':'XL',
            'INST':'INST','INSTANTANEOUS':'INST'}.get(dr, '')
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# Core grid computation  (identical logic to qra_engine_v6.py)
# ═══════════════════════════════════════════════════════════════════════════════
 
def compute_event(formula_type, prob_idx, scenarios, results_data):
    impact_mats = {sz: np.zeros((QY, QX)) for sz in SIZES}
    risk_mats   = {sz: np.zeros((QY, QX)) for sz in SIZES}
    matched = unmatched = 0
 
    prob_col = {9:'prob', 10:'p_jf', 11:'p_lpf', 12:'p_epf',
                13:'p_fb', 14:'p_cve', 15:'p_blv', 16:'p_ff'}
 
    for sc in scenarios:
        key  = sc['key']
        sx   = sc['x'];  sy = sc['y']
        size = map_disch_range_to_size(sc['disch_range'])
        prob = sc.get(prob_col.get(prob_idx, 'prob'), 0.0)
 
        if key not in results_data:
            unmatched += 1
            continue
        matched += 1
 
        d = dist_grid(sx, sy)
 
        if formula_type == 'thermal':
            cell_imp = formula_thermal(d, results_data[key])
        elif formula_type == 'ff':
            lfl_dist, lfl_frac = results_data[key]
            cell_imp = formula_ff(d, lfl_dist, lfl_frac)
        elif formula_type == 'explosion':
            cell_imp = formula_explosion(d, results_data[key]['dists'])
        elif formula_type == 'toxic':
            pairs = results_data[key]
            td = np.array([p[0] for p in pairs])
            tp = np.array([p[1] for p in pairs])
            cell_imp = formula_toxic(d, td, tp)
        else:
            continue
 
        risk_contrib = cell_imp * prob
        impact_mats['Total'] += cell_imp
        risk_mats['Total']   += risk_contrib
        if size in SIZES:
            impact_mats[size] += cell_imp
            risk_mats[size]   += risk_contrib
 
    return impact_mats, risk_mats, matched, unmatched
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# Write result sheets into workbook via xlwings
# ═══════════════════════════════════════════════════════════════════════════════
 
def _ensure_sheet(wb, name):
    """Return sheet by name, creating it at end if it does not exist."""
    sheet_names = [s.name for s in wb.sheets]
    if name in sheet_names:
        ws = wb.sheets[name]
        ws.clear()
        return ws
    # Add after last sheet
    wb.sheets.add(name, after=wb.sheets[-1])
    return wb.sheets[name]
 
 
def write_result_sheets(wb, all_results, directions):
    """
    all_results : {impact_id: {'impact': {sz: ndarray}, 'risk': {sz: ndarray}}}
    directions  : {impact_id: {sz: range_str}}
 
    Writes matrices into SHEET_IMPACT and SHEET_RISK.
    Each matrix is placed starting at the top-left cell of the direction range.
    """
    print(f"\n  Creating result sheets …")
    ws_imp  = _ensure_sheet(wb, SHEET_IMPACT)
    ws_risk = _ensure_sheet(wb, SHEET_RISK)
 
    # Header labels so the sheets are self-documented
    ws_imp.range('A1').value  = 'Impact Matrix Result — written by qra_engine_v6_basic_backline_export.py'
    ws_risk.range('A1').value = 'Risk Matrix Result   — written by qra_engine_v6_basic_backline_export.py'
 
    written = 0
    for impact_id, res in all_results.items():
        dir_map = directions.get(impact_id, {})
        if not dir_map:
            print(f"    Impact ID {impact_id}: no Directions entries — matrices not placed in sheet.")
            continue
 
        imp_mats  = res['impact']
        risk_mats = res['risk']
 
        for sz in SIZES:
            rng_str = dir_map.get(sz)
            if not rng_str:
                continue   # this size not mapped in Directions — skip silently
 
            try:
                start_row, start_col = _parse_range_start(rng_str)
            except ValueError as e:
                print(f"    WARNING: {e}")
                continue
 
            imp_mat  = imp_mats[sz]
            risk_mat = risk_mats[sz]
 
            # xlwings accepts a 2-D list; convert from numpy
            ws_imp.range((start_row, start_col)).value  = imp_mat.tolist()
            ws_risk.range((start_row, start_col)).value = risk_mat.tolist()
            written += 1
 
    print(f"  Matrices written to result sheets: {written} total blocks.")
 
 
# ═══════════════════════════════════════════════════════════════════════════════
# Main
# ═══════════════════════════════════════════════════════════════════════════════
 
def main():
    t0 = time.time()
    print("=" * 80)
    print("QRA Engine V6 — Basic Backline Export")
    print("=" * 80)
    print(f"Source workbook : {V6_PATH}")
    print(f"Export workbook : {EXPORT_PATH}")
    print(f"Grid            : {QX} x {QY}  (SX={SX:.6f}  SY={SY:.6f})")
 
    if not os.path.exists(V6_PATH):
        print(f"\nERROR: Workbook not found:\n  {V6_PATH}")
        print("Update V6_PATH at the top of this script and re-run.")
        sys.exit(1)
 
    # ── Open Excel ────────────────────────────────────────────────────────────
    print("\nOpening Excel workbook …")
    app = xw.App(visible=False)
    app.display_alerts  = False
    app.screen_updating = False
 
    try:
        wb       = app.books.open(V6_PATH)
        ws_core  = wb.sheets['Core']
        print("  Workbook opened OK.")
 
        # ── Verify required sheets exist ──────────────────────────────────────
        sheet_names = [s.name for s in wb.sheets]
        for required in ('Core', 'PageControl', 'Directions'):
            if required not in sheet_names:
                print(f"\n  WARNING: Sheet '{required}' not found.")
                print(f"  Available sheets: {sheet_names}")
 
        # ── Read PageControl ──────────────────────────────────────────────────
        if 'PageControl' in sheet_names:
            active_impacts = read_page_control(wb.sheets['PageControl'])
        else:
            print("  PageControl sheet missing — using all IMPACT_CONFIG entries.")
            active_impacts = {iid: True for iid in IMPACT_CONFIG}
 
        # ── Read Directions ───────────────────────────────────────────────────
        if 'Directions' in sheet_names:
            directions = read_directions(wb.sheets['Directions'])
        else:
            print("  Directions sheet missing — matrices will be computed but NOT placed in result sheets.")
            directions = {}
 
        # ── Iterate impacts ───────────────────────────────────────────────────
        all_results = {}   # {impact_id: {'impact': …, 'risk': …}}
 
        for impact_id in sorted(active_impacts.keys()):
            if impact_id not in IMPACT_CONFIG:
                print(f"\n  Impact ID {impact_id} not in IMPACT_CONFIG — skipped.")
                continue
 
            cfg          = IMPACT_CONFIG[impact_id]
            event_name   = cfg['event']
            sheet_name   = cfg['sheet']
            prob_idx     = cfg['prob_idx']
            formula_type = cfg['formula']
 
            print(f"\n{'─'*70}")
            print(f"  Impact ID {impact_id} | Event={event_name} | "
                  f"Sheet='{sheet_name}' | Formula={formula_type}")
            print(f"{'─'*70}")
 
            # ── Set Impact ID in Core and recalculate ─────────────────────────
            ws_core.range('C2').value = impact_id
            time.sleep(1.5)
            app.api.Calculate()
            time.sleep(1.0)
            print(f"  Core!C2 set to {impact_id} — recalculated.")
 
            # ── Read scenarios from Core ──────────────────────────────────────
            scenarios = read_core_scenarios(ws_core)
            print(f"  Scenarios read from Core: {len(scenarios)}")
            if not scenarios:
                print("  SKIPPED — no scenarios returned by Core.")
                continue
 
            # ── Read effect results sheet ─────────────────────────────────────
            if sheet_name not in sheet_names:
                print(f"  ERROR: Effect sheet '{sheet_name}' not found — skipping impact.")
                continue
 
            ws_res = wb.sheets[sheet_name]
            t_read = time.time()
 
            if formula_type == 'thermal':
                results_data = read_thermal_results(ws_res)
            elif formula_type == 'ff':
                results_data = read_ff_results(ws_res)
            elif formula_type == 'explosion':
                results_data = read_explosion_results(ws_res)
            elif formula_type == 'toxic':
                results_data = read_toxic_results(ws_res)
            else:
                results_data = {}
 
            print(f"  Effect results loaded: {len(results_data)} entries "
                  f"({time.time()-t_read:.1f}s)")
 
            # ── Compute matrices ──────────────────────────────────────────────
            t_calc = time.time()
            impact_mats, risk_mats, matched, unmatched = compute_event(
                formula_type, prob_idx, scenarios, results_data
            )
            print(f"  Grid computed: {time.time()-t_calc:.1f}s  "
                  f"matched={matched}  unmatched={unmatched}")
 
            # Print quick sanity numbers
            for sz in ['Total', 'S', 'M']:
                imp_max  = impact_mats[sz].max()
                risk_max = risk_mats[sz].max()
                nz       = np.count_nonzero(impact_mats[sz])
                print(f"    [{sz:5s}]  impact_max={imp_max:.4f}  "
                      f"risk_max={risk_max:.4e}  nonzero={nz}")
 
            all_results[impact_id] = {
                'impact': impact_mats,
                'risk':   risk_mats,
                'event':  event_name,
            }
 
        # ── Write result sheets ───────────────────────────────────────────────
        if all_results:
            write_result_sheets(wb, all_results, directions)
        else:
            print("\n  No results computed — result sheets not written.")
 
        # ── Save as macro-disabled .xlsx ──────────────────────────────────────
        print(f"\nSaving export workbook …\n  → {EXPORT_PATH}")
        wb.api.SaveAs(EXPORT_PATH, FileFormat=51)   # 51 = xlOpenXMLWorkbook (.xlsx)
        wb.close()
        print("  Saved OK.")
 
    except Exception as exc:
        import traceback
        print(f"\nFATAL ERROR: {exc}")
        traceback.print_exc()
    finally:
        try:
            app.quit()
        except Exception:
            pass
 
    elapsed = time.time() - t0
    print(f"\nDone — total elapsed: {elapsed:.1f}s  ({elapsed/60:.1f} min)")
    print("=" * 80)
 
 
if __name__ == '__main__':
    main()