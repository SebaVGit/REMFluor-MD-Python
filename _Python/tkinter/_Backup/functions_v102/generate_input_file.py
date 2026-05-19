"""
generate_input_file.py — standalone replacement for xlwings version.

Reads all model parameters from AppState, fills template.inp,
and writes input.inp to work_dir.
Mirrors the logic in Source_Py/input_variables.py and generate_input_file.py.
"""
import math
import os
from tkinter import messagebox

from .state import get_state

FT2M = 0.3048   # feet -> metres (input.inp always in metres)
SEC_PER_YR = 60 * 60 * 24 * 365


def _safe_float(val, default=0.0):
    """Parse `val` as float, tolerating comma thousands separators.
    The §7 source-concentration cells store values like "1,600.00"
    for display; without the comma strip, float() raises ValueError
    and every source concentration in input.inp comes out as 0 —
    meaning the model runs with no PFAS at all."""
    if val is None:
        return default
    try:
        return float(val)
    except (TypeError, ValueError):
        # Strip commas (display formatting) and retry once.
        try:
            return float(str(val).replace(",", "").strip())
        except (TypeError, ValueError):
            return default


def _safe_int(val, default=0):
    if val is None:
        return default
    try:
        return int(float(val))
    except (TypeError, ValueError):
        try:
            return int(float(str(val).replace(",", "").strip()))
        except (TypeError, ValueError):
            return default


def _read_txt(path: str) -> dict:
    """Parse a simple key,value .txt file."""
    result = {}
    if not os.path.exists(path):
        return result
    with open(path) as f:
        for line in f:
            line = line.strip()
            if ',' in line:
                parts = line.split(',', 1)
                if parts[0] != 'Parameter':
                    try:
                        result[parts[0].strip()] = float(parts[1].strip())
                    except ValueError:
                        result[parts[0].strip()] = parts[1].strip()
    return result


def _read_hetero(path: str) -> tuple:
    """Return (mdflag, volfrac, difflen) from heterogeneity_inputs.txt."""
    mdflag = 0; volfrac = 1.0; difflen = 0.1
    if not os.path.exists(path):
        return mdflag, volfrac, difflen
    with open(path) as f:
        for line in f:
            ls = line.strip()
            if ls.startswith("mdflag:"):
                try: mdflag = int(ls.split(':')[1].strip())
                except: pass
            elif ls.startswith("Transmissive Fraction of Model (-):"):
                try: volfrac = float(ls.split(':')[1].strip())
                except: pass
            elif ls.startswith("Diffusion Length (m):"):
                try: difflen = float(ls.split(':')[1].strip())
                except: pass
    return mdflag, volfrac, difflen


def _read_retardation(path: str) -> dict:
    """Parse retardation_inputs.txt -> dict of zone data."""
    result = {}
    if not os.path.exists(path):
        return result
    with open(path) as f:
        lines = [ln.rstrip() for ln in f]

    current_zone = None
    skip = False
    in_pfas = False
    for line in lines:
        ls = line.strip()
        if "PFAS Names from Excel" in ls:
            in_pfas = True; continue
        if in_pfas:
            if not ls: in_pfas = False
            continue
        if not ls: continue
        if ls.endswith("Zone"):
            current_zone = ls.replace(" Zone", "")
            result[current_zone] = {}
            skip = True; continue
        if skip: skip = False; continue
        if current_zone and ',' in ls:
            parts = ls.split(',')
            key = parts[0].strip()
            vals = [v.strip() for v in parts[1:]]
            if "Soil Bulk Density" in key:
                try: result[current_zone]['bulkden'] = float(vals[0])
                except: pass
            elif "foc" in key:
                try: result[current_zone]['foc'] = float(vals[0])
                except: pass
            elif "Koc" in key:
                try: result[current_zone]['koc'] = [float(v) for v in vals if v]
                except: pass
    return result


def _read_transformation(path: str) -> dict:
    """Parse transformation_inputs.txt."""
    result = {}
    if not os.path.exists(path):
        return result
    with open(path) as f:
        lines = [ln.rstrip() for ln in f]
    current_zone = None
    skip = False
    for line in lines:
        ls = line.strip()
        if not ls: continue
        if ls.endswith("Zone"):
            current_zone = ls.replace(" Zone", "")
            result[current_zone] = {}
            skip = True; continue
        if skip: skip = False; continue
        if current_zone and ',' in ls:
            parts = ls.split(',')
            key = parts[0].strip()
            vals = [v.strip() for v in parts[1:]]
            if "Yield" in key:
                try: result[current_zone]['Yield'] = [float(v) for v in vals if v]
                except: pass
            elif "Decay" in key:
                try: result[current_zone]['Decay'] = [float(v) for v in vals if v]
                except: pass
    return result


def _smart_fmt(v, fallback="0"):
    """Smart precision picker that trims trailing zeros.  Used to make
    input.inp values human-readable after feet->meters conversion (e.g.
    1.999488 -> "2", 3.2000001024 -> "3.2", 0.04 -> "0.04").  Rules:
      abs(v) >= 100  : 0 decimals
      abs(v) >= 1    : up to 2 decimals
      abs(v) >= 0.01 : up to 4 decimals
      else           : 4 significant figures
    Trailing zeros after the decimal are stripped.
    """
    try:
        if v is None:
            return fallback
        x = float(v)
    except (TypeError, ValueError):
        return fallback
    if x == 0:
        return "0"
    a = abs(x)
    if a >= 100:
        s = f"{x:.0f}"
    elif a >= 1:
        s = f"{x:.2f}"
    elif a >= 0.01:
        s = f"{x:.4f}"
    else:
        s = f"{x:.4g}"
    if '.' in s and 'e' not in s and 'E' not in s:
        s = s.rstrip('0').rstrip('.')
        if not s or s == '-':
            s = "0"
    return s


def build_inp_data(state) -> dict:
    """
    Read all values from AppState and build the same dict that
    input_variables.read_excel_data() returns.
    """
    work_dir = state.work_dir or os.getcwd()

    version_flag = state.get("A8", 2)     # 1=Simple, 2=Detailed
    unit_flag    = state.get("AD1", 2)    # 1=feet, 2=metres
    ft = (unit_flag == 1)

    def c(val):
        """Convert feet->metres if needed."""
        v = _safe_float(val) if val is not None else None
        return v * FT2M if (ft and v is not None) else v

    PSB_flag = bool(state.get("R22", False))

    # version / precursor
    ncomp = 1
    g38 = state.get("G38")
    if g38 and str(g38).strip() not in ('', 'None', 'none'):
        ncomp = 2
    # K38 (v_pfaa3) is the §5 Detailed Precursor 1 species dropdown.
    # "PFAA 1-able" = a precursor that transforms into PFAA 1 (i.e.
    # precursor mode IS active).  Bug fix: previous code excluded
    # "PFAA 1-able" from triggering ipre=1, which meant Detailed runs
    # silently dropped to ipre=0 — Precursor 1 source concentrations
    # zeroed out and the Detailed example produced wrong / no plume
    # output even though the Excel reference always sets ipre=1 here.
    k38 = state.get("K38")
    precursor_flag = (version_flag == 2 and k38 and
                      str(k38).strip() not in ('', 'None', 'none'))
    ipre  = 1 if precursor_flag else 0
    # v102: iwall = 2 when PSB activated, 0 when off (matches the
    # Detailed Excel/Fortran convention used in inptest.inp).
    iwall = 2 if PSB_flag else 0

    # time
    startT = _safe_float(state.get("E18", 1977))
    endT   = _safe_float(state.get("E19", 2077))

    # high-K transport
    porf  = _safe_float(state.get("G22", 0.2))
    tortf = 0.000001

    # retardation file
    retard_data = _read_retardation(os.path.join(work_dir, 'retardation_inputs.txt'))
    bulkden = retard_data.get('Transmissive', {}).get('bulkden', 1.6)
    foc     = retard_data.get('Transmissive', {}).get('foc', 0.001)
    koc_list = retard_data.get('Transmissive', {}).get('koc', [120, 1, 1, 1])
    ock2, ock4 = (koc_list[0] if len(koc_list)>0 else 120,
                  koc_list[1] if len(koc_list)>1 else 1)
    ock1, ock3 = ((koc_list[2] if len(koc_list)>2 else 1,
                   koc_list[3] if len(koc_list)>3 else 1) if ipre==1 else (0, 0))

    # dispersivity (convert if feet)
    alphax = c(state.get("V4", 3.2))
    alphay = c(state.get("X4", 0.04))
    alphaz = c(state.get("Z4", 0.004))

    # model size (convert if feet)
    # v102: snap to 4 decimals to wipe floating-point dust from
    # feet↔meters round-trips (e.g. 500 m → 1640.42 ft → ×0.3048 →
    # 500.000016 m).  Without this, downstream banker's-rounding in
    # nx1/ny/nz can flip on the dust (round(50.000008)=50 stable but
    # round(2.500008)=3 vs round(2.5)=2 banker's → different lzsourcemin).
    X = c(state.get("E11", 500))
    Y = c(state.get("E12", 50))
    Z = c(state.get("E13", 10))
    if X is not None: X = round(X, 4)
    if Y is not None: Y = round(Y, 4)
    if Z is not None: Z = round(Z, 4)

    # cell sizes from cellsize_input.txt
    # v102: respect the sidecar's OWN Unit Flag, not the live UI flag.
    # Previously this used `ft` (the UI's current state["AD1"]), which
    # caused a double-conversion bug: user pastes example in meters
    # → sidecar has values + "Unit Flag: 2", toggle UI to feet, run
    # again → sidecar STILL has meters values + flag 2, but `ft=True`
    # made the code multiply by 0.3048 anyway → cell sizes shrank to
    # ~1/3 their true value (dy 5 m → 1.524 m, dz 2 m → 0.61 m).
    cell_data = _read_txt(os.path.join(work_dir, 'cellsize_input.txt'))
    dx = cell_data.get('Cell Size X:', 5.0)
    dy_file = cell_data.get('Cell Size Y:', None)
    dz_file = cell_data.get('Cell Size Z:', None)
    # Read the sidecar's own unit flag (1=feet, 2=meters).  Default 2
    # if the sidecar is missing the line (back-compat with older files).
    try:
        cell_unit = int(float(cell_data.get('Unit Flag:', 2)))
    except (TypeError, ValueError):
        cell_unit = 2
    if cell_unit == 1:    # sidecar values are in FEET → convert
        dx      = dx * FT2M      if dx      else dx
        dy_file = dy_file * FT2M if dy_file else dy_file
        dz_file = dz_file * FT2M if dz_file else dz_file

    # grid — PSB distance + width.  The §9 UI cells now live at the
    # MODERN addresses X74 (psb_dist) and Y82 (psb_width); AC82 is the
    # # of cells in PSB.  Earlier versions wrote/read these at the
    # LEGACY addresses AE25 / AH27 / AH29 — but the UI was migrated in
    # v38 to the modern set, so the legacy reads here returned None
    # and the PSB block silently disappeared from input.inp.
    # Resolution order: modern UI addr → legacy addr → 0.
    def _addr_first(*addrs):
        for a in addrs:
            v = state.get(a)
            if v is not None and str(v).strip() not in ("", "None"):
                return v
        return None

    x1 = c(_addr_first("X74", "AE25")); x1 = x1 or 0
    x2_offset = c(_addr_first("Y82", "AH27")); x2_offset = x2_offset or 0
    x2 = x1 + x2_offset

    if iwall == 0:
        nx1 = round(X / dx) if dx else 100
        nxpsb = 0; nx2 = 0
    else:
        nxpsb_raw = _addr_first("AC82", "AH29")
        nxpsb = _safe_int(nxpsb_raw, 0)
        nx1   = round(x1 / dx) if dx else 0
        nx2   = round((X - x2) / dx) if dx else 0

    if dy_file:
        dy = round(dy_file, 2)
        ny = round(Y / dy) if dy > 0 else 5
    else:
        ny = 5
        dy = round(Y / ny, 2)

    if dz_file:
        dz = round(dz_file, 2)
        nz = round(Z / dz) if dz > 0 else 10
    else:
        nz = 10
        dz = round(Z / nz, 2)

    # source concentrations (11 rows)
    # v102: build raw lists first, then trim ".0" off integer values
    # so input.inp shows "0, 1600, 1600, 0, 0" instead of "0, 1600.0,..."
    def _trim_num(x):
        try:
            xf = float(x)
        except (TypeError, ValueError):
            return x
        if xf == int(xf):
            return int(xf)
        return xf
    times  = [_trim_num(_safe_float(state.get(f"U{8+i}"), startT) - startT)
              for i in range(11)]
    concs2 = [_trim_num(_safe_float(state.get(f"V{8+i}"), 0)) for i in range(11)]
    concs4 = [_trim_num(_safe_float(state.get(f"X{8+i}"), 0)) for i in range(11)]
    concs1 = ([_safe_float(state.get(f"Z{8+i}"), 0) for i in range(11)]
              if ipre == 1 else [0]*11)                                    # Precursor 1
    concs3 = ([_safe_float(state.get(f"AB{8+i}"), 0) for i in range(11)]
              if (ipre == 1 and ncomp == 2) else [0]*11)                  # Precursor 2

    # NOTE: calibration czero(*,n) multipliers are now baked into the
    # §7 cells PHYSICALLY at calibration finish + Load Optimal Data
    # (see main.py _apply_calib_multipliers_to_s7).  Reading the §7
    # cells already reflects the multiplied values, so we don't need
    # a virtual multiply here — that would double-apply.

    # source geometry
    # v102: snap source_width to wipe FP dust (same reason as X/Y/Z).
    source_width = c(state.get("E15", 60))
    if source_width is not None:
        source_width = round(source_width, 4)
    lysource = math.ceil(source_width / dy / 2) if source_width and dy else 0

    top_z = Z
    source_thick = c(state.get("E16", 5))
    # v102: snap top_z / source_thick to 4 decimals to wipe floating-
    # point dust from feet↔meters conversions (e.g. 10 m → 32.8084 ft
    # → c() → 10.0000016 m).  Without this, lzsourcemin's banker's-
    # rounding behaviour flipped between unit toggles (round(2.5)=2,
    # round(2.5000008)=3), giving different cell indices in input.inp.
    if source_thick is not None:
        source_thick = round(source_thick, 4)
    top_z = round(top_z, 4) if top_z is not None else top_z
    bot_z = (top_z - source_thick) if source_thick is not None else None
    if bot_z is not None:
        bot_z = round(bot_z, 4)

    # compute layer indices
    fzGridBlock = round(Z / dz) if dz else nz
    lzsourcemin = round(bot_z / round(dz, 2)) + 1 if bot_z is not None and dz else 1
    lzsourcemin = min(lzsourcemin, fzGridBlock)
    lzsourcemax_val = round(top_z / round(dz, 2)) if dz else nz
    lzsourcemax_val = max(lzsourcemax_val, lzsourcemin)
    lzsourcemax_val = min(lzsourcemax_val, fzGridBlock)

    # velocity
    # v102: read v_darcy via state["C22"] FIRST (that's the live UI
    # value that gets unit-converted by _apply_units on §1 toggle).
    # state["E22"] is a one-shot fallback written by inp_to_state during
    # Paste Example / Load Data; it never updates after a units toggle,
    # so reading it first caused vd to come out in the wrong unit (e.g.
    # 10 m/yr in state["E22"] interpreted as ft/yr → 3.048 m/yr in
    # input.inp after toggling to feet).
    vd = c(state.get("C22", state.get("E22", 10)))
    # v102: keep state["E22"] in sync with the live v_darcy reading so
    # downstream callers (cali_1, generate_dashboard) that may still
    # consult E22 don't see a stale Paste-Example value.  Cheap defence.
    try:
        if vd is not None:
            state.set("E22", state.get("C22"))
    except Exception:
        pass

    # low-K zone
    porm  = _safe_float(state.get("K27", 0.48))
    tortm = _safe_float(state.get("K28", 0.56))
    focm  = 0.001
    # v102: state.get returns the STORED value (which may be None or "")
    # rather than the second-arg default when the key exists.  After
    # Clear All Data, state["E44"] is None — passing that to _safe_float
    # gave 0.0, and diff = 0 × seconds_per_year = 0 made the Fortran
    # solver produce NaN concentrations (mdflag=2 + diff=0 = division
    # by zero in the matrix-diffusion solver).  Defensive: if the
    # state value is None, empty, zero, or unparseable, fall back to
    # the standard PFOS-class diffusion 3.5e-10 m²/s.
    _e44 = state.get("E44")
    if _e44 is None or (isinstance(_e44, str) and not str(_e44).strip()):
        diff_m2s = 3.5e-10
    else:
        diff_m2s = _safe_float(_e44, 3.5e-10)
        if not diff_m2s or diff_m2s <= 0:
            diff_m2s = 3.5e-10
    diff = diff_m2s * SEC_PER_YR   # m²/yr for .inp

    # heterogeneity
    mdflag, volfrac, difflen = _read_hetero(
        os.path.join(work_dir, 'heterogeneity_inputs.txt'))

    # high-K yield / decay
    yieldf2 = _safe_float(state.get("K42", 0)) if ipre == 1 else 0
    yieldf4 = (_safe_float(state.get("M42", 0)) if (ipre == 1 and ncomp == 2) else 0)
    yieldf3 = 0

    def _decay(addr):
        hl = _safe_float(state.get(addr))
        if hl and hl > 0:
            v = math.log(2) / hl
            return v if v > 0.001 else 0
        return 0

    decayf1 = _decay("K41") if ipre == 1 else 0
    decayf2 = 0
    decayf3 = _decay("M41") if (ipre == 1 and ncomp == 2) else 0
    decayf4 = 0

    # low-K transformation
    yieldm2 = yieldm4 = 0; yieldm3 = 0
    decaym1 = decaym2 = decaym3 = decaym4 = 0
    if ipre == 1:
        transf = _read_transformation(os.path.join(work_dir, 'transformation_inputs.txt'))
        lowk = transf.get('Low-K', {})
        if 'Yield' in lowk and len(lowk['Yield']) >= 2:
            yieldm2, yieldm4 = lowk['Yield'][0], lowk['Yield'][1]
        if 'Decay' in lowk and len(lowk['Decay']) >= 1:
            v1 = math.log(2) / lowk['Decay'][0] if lowk['Decay'][0] else 0
            decaym1 = v1 if v1 > 0.001 else 0
            if ncomp == 2 and len(lowk['Decay']) >= 2 and lowk['Decay'][1]:
                v3 = math.log(2) / lowk['Decay'][1]
                decaym3 = v3 if v3 > 0.001 else 0
        # High-K values as fallback for low-K if not in transform file
        if decaym1 == 0: decaym1 = decayf1
        if decaym3 == 0: decaym3 = decayf3

    # PSB parameters
    if not PSB_flag:
        fra1 = fra2 = fra3 = fra4 = 0
        fcackf1 = fcackf2 = fcackf3 = fcackf4 = 0
        tplume1 = tplume2 = 0
    else:
        # Prefer AA82 (the UI-bound v_psb_load StringVar — stored as
        # PERCENT, e.g. 0.24 means 0.24% = 0.0024 fraction).  Strip a
        # trailing "%" if the user typed it, then divide by 100.
        # Fall back to legacy AH28 (raw fraction) if AA82 is empty.
        _aa82 = state.get("AA82")
        if _aa82 is not None and str(_aa82).strip() not in ("", "None"):
            try:
                _s = str(_aa82).rstrip('%').strip()
                PSBloading = float(_s) / 100.0
            except (ValueError, TypeError):
                PSBloading = _safe_float(state.get("AH28", 0))
        else:
            PSBloading = _safe_float(state.get("AH28", 0))
        # Year PSB installed lives at AB28 (v_psb_yr).  Fall back to the
        # legacy Y75 only — AB26 in the legacy code was the Converted-Kf
        # cell, totally unrelated; using it for tplume1 produced garbage.
        tplume1_raw = state.get("AB28") or state.get("Y75")
        tplume1 = (_safe_float(tplume1_raw) - startT) if tplume1_raw else 0
        tplume2 = 700

        fra2  = _safe_float(state.get("V23", 0))
        fra4  = _safe_float(state.get("X23", 0))
        fra1  = _safe_float(state.get("Z23", 0)) if ipre == 1 else 0
        fra3  = _safe_float(state.get("AB23", 0)) if (ipre == 1 and ncomp == 2) else 0

        # v102: use the CONVERTED Kf cells (V26 / X26 / Z26 / AB26) as
        # the authoritative source, falling back to the raw Kf cells
        # (V24 / X24 / Z24 / AB24) only if the converted value is blank.
        # The Fortran solver expects Kf in canonical (ug/kg)(ug/L)^-a
        # units — the §9 unit dropdown auto-fills V26..AB26 with the
        # raw V24..AB24 values rescaled to ug/kg by the trace in
        # _recompute_psb_conv_kf.  Reading V24 directly would feed the
        # raw user-typed value in whatever unit the dropdown shows,
        # which would skew the model.
        def _kf_pick(conv_addr, raw_addr):
            cv = state.get(conv_addr)
            if cv is not None and str(cv).strip() not in ("", "None"):
                return _safe_float(cv, 0)
            return _safe_float(state.get(raw_addr, 0), 0)
        kf_conv = _kf_pick("V26", "V24") or 1
        fcackf2 = _kf_pick("V26",  "V24")  * PSBloading
        fcackf4 = _kf_pick("X26",  "X24")  * PSBloading
        fcackf1 = (_kf_pick("Z26",  "Z24")  * PSBloading
                   if ipre == 1 else 0)
        fcackf3 = (_kf_pick("AB26", "AB24") * PSBloading
                   if (ipre == 1 and ncomp == 2) else 0)

    # numerical
    num_data = _read_txt(os.path.join(work_dir, 'numerical_inputs.txt'))
    iTVD = int(num_data.get("iTVD", 1))
    dt   = num_data.get("Timestep Size (yr) ", 0.1)
    nt   = int((endT - startT) / dt)
    see_every = _safe_float(state.get("V47", 10))
    npt  = int(see_every / dt) if dt else 100
    tol_raw = num_data.get("Convergence Tolerance (ug/L)", 1e-6)
    tol  = "{:.1e}".format(float(tol_raw))
    maxit = 5000

    # observation wells
    # v102 default: zwelltop = Z, zwellbot = Z/2 (constant for every well).
    # Matches Example/*/input.inp format:  1, 60.0, 0, 10.0, 5.0
    #
    # v102 override: if the user loaded a calibration .xlsx and
    # popups_calibration._import_xlsx_into_app saved per-well screen
    # depths to mw_observations.json (well_depths), use those.  This
    # restores the legacy Source_Py Detailed-mode behaviour of
    # honouring real screen-interval depths from the .xlsx Model
    # Location sheet — but only when those columns exist.
    wells = []
    z_top_default = float(Z) if (Z is not None and Z != 0) else 10.0
    z_bot_default = z_top_default / 2.0

    # Try to load per-well depths from mw_observations.json sidecar.
    # The sidecar stores depths in the USER'S unit (ft or m); c() will
    # convert to meters for input.inp.
    well_depths_map = {}
    try:
        import json as _json
        obs_path = os.path.join(work_dir, 'mw_observations.json')
        if os.path.exists(obs_path):
            with open(obs_path, 'r', encoding='utf-8') as fh:
                obs_payload = _json.load(fh) or {}
            wd = obs_payload.get('well_depths', {}) or {}
            for wname, d in wd.items():
                if isinstance(d, dict):
                    well_depths_map[str(wname).strip()] = (
                        d.get('top'), d.get('bot'))
    except Exception:
        well_depths_map = {}

    for i in range(7):
        row = 34 + i
        xwell = c(state.get(f"AF{row}"))
        if xwell is None:
            continue
        try:
            xwell_f = float(xwell)
        except (TypeError, ValueError):
            continue
        nwell = i + 1
        # Resolve per-well depths: check mw_observations.json under
        # this well's name (U34..U40 / v_mw_names) — fall back to
        # default Z, Z/2 if missing or invalid.
        z_top = z_top_default
        z_bot = z_bot_default
        well_name = state.get(f"U{row}")
        if well_name:
            d = well_depths_map.get(str(well_name).strip())
            if d:
                top_v, bot_v = d
                try:
                    if top_v is not None:
                        tv = c(top_v)
                        if tv is not None:
                            z_top = float(tv)
                except Exception:
                    pass
                try:
                    if bot_v is not None:
                        bv = c(bot_v)
                        if bv is not None:
                            z_bot = float(bv)
                except Exception:
                    pass
        wells.append(
            f"{nwell}, {_smart_fmt(xwell_f)}, 0, "
            f"{_smart_fmt(z_top)}, {_smart_fmt(z_bot)}"
        )

    # v102: use module-level _smart_fmt so we can call it from inside
    # the obs-well loop above as well.  Same trim-trailing-zero behaviour.
    _smart = _smart_fmt

    # v102: _smart only on values that went through c() (feet->m) and
    # therefore carry float dust like 0.0040000001280000004.
    alphax = _smart(alphax)
    alphay = _smart(alphay)
    alphaz = _smart(alphaz)
    x1     = _smart(x1)
    x2     = _smart(x2)
    X      = _smart(X)       # xmax
    dy     = _smart(dy)
    dz     = _smart(dz)
    difflen = _smart(difflen)
    # _trim: strip trailing ".0" from integer-valued floats so e.g.
    # times "10.0" -> "10" and concentrations "1600.0" -> "1600", but
    # keep "10.5" as "10.5" and "0.011992" as "0.011992".  Used for
    # fields that DON'T need rounding (raw user data + computed
    # rates) but should still look tidy in input.inp.
    def _trim(v):
        try:
            x = float(v)
        except (TypeError, ValueError):
            return v if v is not None else 0
        if x == int(x):
            return int(x)
        return x
    # v102: vd preserved raw (e.g. 0.011992); diff & tortm to 3 decimals
    # (user request — readable but enough precision for the solver).
    def _fmt3(v, fallback="0"):
        try:
            x = float(v)
        except (TypeError, ValueError):
            return fallback
        if x == 0:
            return "0"
        s = f"{x:.3f}"
        if '.' in s:
            s = s.rstrip('0').rstrip('.')
            if not s or s == '-':
                s = "0"
        return s
    vd       = _trim(vd)        # preserve precision (e.g. 0.011992)
    diff     = _fmt3(diff)      # 3 dec (e.g. 0.011 from 0.0110376)
    tortm    = _fmt3(tortm)     # 3 dec (e.g. 0.562 from 0.5617222...)
    tplume1  = _trim(tplume1)
    tplume2  = _trim(tplume2)
    fra1, fra2, fra3, fra4 = _trim(fra1), _trim(fra2), _trim(fra3), _trim(fra4)
    fcackf1, fcackf2, fcackf3, fcackf4 = (
        _trim(fcackf1), _trim(fcackf2), _trim(fcackf3), _trim(fcackf4))
    # tortm has many digits in some examples (0.5617222837346214) —
    # leave it raw, only _trim if integer-valued.

    return dict(
        ncomp=ncomp, ipre=ipre, iwall=iwall, iTVD=iTVD,
        times=times, concs1=concs1, concs2=concs2, concs3=concs3, concs4=concs4,
        lysource=lysource, lzsourcemax=lzsourcemax_val, lzsourcemin=lzsourcemin,
        vd=vd, porf=porf, tortf=tortf, bulkden=bulkden, foc=foc,
        alphax=alphax, alphay=alphay, alphaz=alphaz,
        ock1=ock1, ock2=ock2, ock3=ock3, ock4=ock4,
        x1=x1, x2=x2, tplume1=tplume1, tplume2=tplume2,
        fra1=fra1, fra2=fra2, fra3=fra3, fra4=fra4,
        fcackf1=fcackf1, fcackf2=fcackf2, fcackf3=fcackf3, fcackf4=fcackf4,
        yieldf2=yieldf2, yieldf3=yieldf3, yieldf4=yieldf4,
        decayf1=decayf1, decayf2=decayf2, decayf3=decayf3, decayf4=decayf4,
        mdflag=mdflag, porm=porm, tortm=tortm, focm=focm, diff=diff,
        volfrac=volfrac, difflen=difflen,
        yieldm2=yieldm2, yieldm3=yieldm3, yieldm4=yieldm4,
        decaym1=decaym1, decaym2=decaym2, decaym3=decaym3, decaym4=decaym4,
        xmax=X, nx1=nx1, nxpsb=nxpsb, nx2=nx2,
        dy=dy, ny=ny, dz=dz, nz=nz, dt=dt, nt=nt, npt=npt, tol=tol, maxit=maxit,
        wells=wells,
    )


def run(app) -> bool:
    """
    Snapshot state from UI, build input.inp, write to work_dir.
    Returns True on success.
    """
    state = get_state()
    state.snapshot(app)

    work_dir = state.work_dir or os.getcwd()
    # v100: template.inp is a bundled read-only asset.  In a frozen
    # .exe build it lives at sys._MEIPASS, not next to the .exe.  Try
    # work_dir first (user override), then state.bundle_dir (frozen
    # bundle).  output_path stays in work_dir (writable).
    template_path = os.path.join(work_dir, 'template.inp')
    if not os.path.exists(template_path):
        bundle_dir = getattr(state, "bundle_dir", "") or work_dir
        alt = os.path.join(bundle_dir, 'template.inp')
        if os.path.exists(alt):
            template_path = alt
    output_path   = os.path.join(work_dir, 'input.inp')

    if not os.path.exists(template_path):
        messagebox.showerror("Error",
            f"template.inp not found in:\n{work_dir}\n"
            f"or in bundle_dir:\n"
            f"{getattr(state, 'bundle_dir', '(unset)')}")
        return False

    try:
        data = build_inp_data(state)
    except Exception as e:
        messagebox.showerror("Error", f"Could not build model parameters:\n{e}")
        return False

    # Format source zone data
    src_lines = []
    for t, c1, c2, c3, c4 in zip(
            data['times'], data['concs1'], data['concs2'],
            data['concs3'], data['concs4']):
        src_lines.append(f"{t or 0}, {c1 or 0}, {c2 or 0}, {c3 or 0}, {c4 or 0}")
    source_zone_str = "\n".join(src_lines)
    wells_str = "\n".join(data['wells'])

    try:
        with open(template_path) as f:
            template = f.read()
        filled = template.format(
            source_zone_data=source_zone_str,
            observation_wells=wells_str,
            **{k: v for k, v in data.items() if k not in ('times','concs1','concs2',
                                                            'concs3','concs4','wells')},
        )
        with open(output_path, 'w') as f:
            f.write(filled)
        print(f"input.inp written to: {output_path}")
        return True
    except Exception as e:
        messagebox.showerror("Error", f"Could not write input.inp:\n{e}")
        return False
