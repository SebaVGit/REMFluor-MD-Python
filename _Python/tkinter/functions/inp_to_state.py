"""
inp_to_state.py — write parsed input.inp data into AppState.

Adapted from write_values_to_excel() in Source_Py/restore_from_saved_folder.py,
but writes to state dict instead of xlwings cells.
"""
import math


M2FT = 1.0 / 0.3048   # meters -> feet


def _blank(v):
    if v is None:
        return None
    if isinstance(v, str) and v.strip().lower() in ('none', ''):
        return None
    return v


def _fmt_kf(v):
    """Format a Freundlich Kf value for the §9 cells.  Whole numbers
    drop the .0 ("1,227,951"); fractional values keep up to 4 decimals
    with trailing zeros stripped ("1,234.56").  Empty / 0 → "0"."""
    try:
        x = float(v)
    except (TypeError, ValueError):
        return "" if v in (None, "") else str(v)
    if x == 0:
        return "0"
    if abs(x - round(x)) < 1e-9:
        return f"{int(round(x)):,}"
    return f"{x:,.4f}".rstrip("0").rstrip(".")


def write_inp_to_state(state, data: dict, additional: dict, unit_flag: int):
    """
    Push all values from a parsed input.inp (data) and additional info
    (additional) into AppState.

    unit_flag: 1 = feet, 2 = meters
    Data in input.inp is always in metres; convert back to feet if unit_flag==1.
    """
    ft = (unit_flag == 1)

    def cv(val):
        """Conditionally convert m -> ft."""
        return val * M2FT if (ft and val is not None) else val

    start_t = additional.get('start_year', 0) or 0

    # ── Source zone concentrations ────────────────────────────────────────
    # Format with comma grouping + 2 decimals so the §7 cells render
    # consistently (e.g. "1,600.00" not "1600.0").  _safe_float in
    # generate_input_file strips commas before re-parsing.
    def _fmt_conc(v):
        if v is None:
            return None
        try:
            return f"{float(v):,.2f}"
        except (TypeError, ValueError):
            return v

    for i, src in enumerate(data.get('source_zone_data', [])[:11]):
        row = 8 + i
        state.set(f"U{row}", src['time'] + start_t)
        state.set(f"V{row}", _fmt_conc(src['c2']))           # PFAA 1
        if data.get('ncomp', 1) == 2:
            state.set(f"X{row}", _fmt_conc(src['c4']))       # PFAA 2
        if data.get('ipre', 0) == 1:
            state.set(f"Z{row}", _fmt_conc(src['c1']))       # Precursor 1
            if data.get('ncomp', 1) == 2:
                state.set(f"AB{row}", _fmt_conc(src['c3']))  # Precursor 2

    # ── High-K transport ──────────────────────────────────────────────────
    state.set("G22", data.get('porf'))
    state.set("V4",  cv(data.get('alphax')))
    state.set("X4",  cv(data.get('alphay')))
    state.set("Z4",  cv(data.get('alphaz')))

    # ── Model dimensions ──────────────────────────────────────────────────
    xmax = data.get('xmax')
    state.set("E11", cv(xmax))

    dy = data.get('dy', 1); ny = data.get('ny', 1)
    state.set("E12", cv(dy * ny))

    dz = data.get('dz', 1); nz = data.get('nz', 1)
    state.set("E13", cv(dz * nz))

    # Source width: lysource * dy * 2
    lysource = data.get('lysource', 0)
    if lysource and dy:
        state.set("E15", cv(lysource * dy * 2))

    # ── Velocity ──────────────────────────────────────────────────────────
    vd_val = cv(data.get('vd'))
    state.set("C22", vd_val)   # UI widget v_darcy
    state.set("E22", vd_val)   # also keep E22 for generate_input_file fallback

    # ── Low-K zone ───────────────────────────────────────────────────────
    _porm = data.get('porm'); state.set("K27", round(_porm, 2) if _porm is not None else None)
    _tortm = data.get('tortm'); state.set("K28", round(_tortm, 2) if _tortm is not None else None)

    # Molecular diffusion: stored as m²/yr in .inp; display as m²/s
    diff = data.get('diff')
    if diff is not None:
        state.set("E44", diff / (60 * 60 * 24 * 365))

    # ── PSB / plume remediation ──────────────────────────────────────────
    iwall = data.get('iwall', 0)
    state.set("R22", bool(iwall))

    if iwall:
        x1 = data.get('x1'); x2 = data.get('x2')
        # PSB Distance from Source — UI cell X74 (legacy AE25 kept too).
        state.set("X74",  cv(x1))
        state.set("AE25", cv(x1))
        if x1 is not None and x2 is not None:
            width = cv(x2 - x1)
            # Total Width of PSB in X-Direction — UI cell Y82 (legacy AH27).
            state.set("Y82",  width)
            state.set("AH27", width)

        version_flag = state.get("A8", 2)
        tplume1 = data.get('tplume1', 0)
        if version_flag == 1:
            state.set("AB26", tplume1 + start_t if tplume1 else None)
        else:
            state.set("AB28", tplume1 + start_t if tplume1 else None)
            n_cells = data.get('nxpsb')
            # # of cells in PSB in x-direction — UI cell AC82 (legacy AH29).
            state.set("AC82", n_cells)
            state.set("AH29", n_cells)

        # v106 FIX: psb_loading from store_info text can be a STRING
        # (e.g. "None" or a comma-grouped number).  Coerce to float BEFORE
        # any arithmetic — previously "0 / psb_load" with a str divisor
        # raised "unsupported operand type(s) for /: 'int' and 'str'" and
        # aborted the whole Load.
        _psb_raw = additional.get('psb_loading')
        try:
            psb_load = (float(str(_psb_raw).replace(",", "").strip())
                        if _psb_raw not in (None, "") else None)
        except (TypeError, ValueError):
            psb_load = None
        if psb_load is not None:
            # PSB Loading 'fcac' — UI cell AA82 displays the value as
            # a *percentage* (e.g. 0.0024 fraction → 0.24 in the §9
            # cell next to the "(%)" unit label).  AH28 keeps the raw
            # fraction for any code path that still reads it.
            state.set("AA82", psb_load * 100)
            state.set("AH28", psb_load)

        state.set("V23", data.get('fra2'))
        state.set("X23", data.get('fra4'))
        if data.get('ipre', 0) == 1:
            state.set("Z23", data.get('fra1'))
            state.set("AB23", data.get('fra3'))
        if psb_load:
            # input.inp stores fcackf in ug; V24 = fcackf/loading is
            # therefore in (ug/kg)(ug/L)^(-a).  Force the unit dropdown
            # so the §9 Converted-Kf auto-formula stays self-consistent.
            state.set("U24", "(ug/kg)(ug/L)^(-a)")
            # Format with thousands separators so the §9 cells read
            # "1,227,951" instead of "1227951.0" (the §9 trace strips
            # commas before re-converting to float).
            state.set("V24", _fmt_kf((data.get('fcackf2') or 0) / psb_load))
            state.set("X24", _fmt_kf((data.get('fcackf4') or 0) / psb_load))
            if data.get('ipre', 0) == 1:
                state.set("Z24",  _fmt_kf((data.get('fcackf1') or 0) / psb_load))
                state.set("AB24", _fmt_kf((data.get('fcackf3') or 0) / psb_load))
    else:
        for addr in ("X74","Y82","AA82","AC82",
                     "V23","X23","Z23","AB23",
                     "V24","X24","Z24","AB24",
                     "AE25","AH27","AH28","AH29"):
            state.set(addr, None)

    # ── Yield / decay (high-K) ────────────────────────────────────────────
    if data.get('ipre', 0) == 1:
        state.set("K42", data.get('yieldf2'))
        if data.get('ncomp', 1) == 2:
            state.set("M42", data.get('yieldf4'))
        for decay_val, addr in [(data.get('decayf1'), 'K41'),
                                (data.get('decayf3'), 'M41')]:
            if decay_val and decay_val > 0.001:
                state.set(addr, math.log(2) / decay_val)

    # ── Numerical ─────────────────────────────────────────────────────────
    # V47 is the §11 "See Results Every" cell — it stores the YEAR
    # interval, not the raw timestep count.  generate_input_file writes
    # npt = int(see_every / dt) to input.inp; to restore the UI we
    # invert: see_every = npt * dt.  Earlier code wrote raw npt here,
    # which made the §11 cell display 1000 instead of 100 on Paste
    # Example / Load Data with the default dt=0.1.
    npt = data.get('npt')
    dt  = data.get('dt')
    if npt is not None and dt:
        see_every = npt * dt
        # Render as int when integral, else trim trailing zeros
        if abs(see_every - round(see_every)) < 1e-9:
            state.set("V47", int(round(see_every)))
        else:
            state.set("V47", round(see_every, 4))
    elif npt is not None:
        state.set("V47", npt)   # last-resort fallback (dt missing)

    # ── Observation wells (simple version) ───────────────────────────────
    for i, well in enumerate(data.get('wells', [])[:7]):
        row = 34 + i
        state.set(f"AF{row}", cv(well.get('xwell')))

    # ── Additional info ───────────────────────────────────────────────────
    if additional:
        state.set("B4",  _blank(additional.get('site_location_id')))
        state.set("E4",  _blank(additional.get('date')))
        state.set("E16", _blank(additional.get('thickness')))
        # v102: explicit overrides — Save Data writes these in user-unit
        # so we DON'T cv()-convert them (the recovered input.inp value
        # was already cv-converted in the §model-dimensions block above;
        # this just overwrites with the exact user-typed value if the
        # store_info has the explicit field).
        if additional.get('source_width') is not None:
            state.set("E15", additional['source_width'])
        if additional.get('model_x_size') is not None:
            state.set("E11", additional['model_x_size'])
        if additional.get('model_y_size') is not None:
            state.set("E12", additional['model_y_size'])
        if additional.get('model_z_size') is not None:
            state.set("E13", additional['model_z_size'])
        if additional.get('vd_user_unit') is not None:
            state.set("C22", additional['vd_user_unit'])
            state.set("E22", additional['vd_user_unit'])
        if additional.get('porf') is not None:
            state.set("G22", additional['porf'])
        state.set("E18", _blank(additional.get('start_year')))
        state.set("E19", _blank(additional.get('end_year')))
        state.set("AH7", _blank(additional.get('source_treatment_start_year')))
        state.set("AH8", _blank(additional.get('source_treatment_end_year')))
        # v102: source_concentration_reduction in legacy store_info files
        # is stored as a FRACTION (e.g. 0.5 means 50%).  The UI's §8 cell
        # expects PERCENT, so multiply by 100 when the loaded value is
        # in the (0, 1] range — keeps any already-percentage values
        # (e.g. 50, 75) intact.
        _scr = additional.get('source_concentration_reduction')
        try:
            _scr_f = float(_scr) if _scr is not None else None
            if _scr_f is not None and 0.0 < _scr_f <= 1.0:
                _scr = _scr_f * 100.0
        except (TypeError, ValueError):
            pass
        state.set("AH9", _blank(_scr))
        # Mirror to UI-bound cells (Section 8)
        state.set("D27", _blank(additional.get('source_treatment_start_year')))
        state.set("D28", _blank(_scr))

        state.set("Y74", _blank(additional.get('sample_year')))  # Section 10 sample yr
        if state.get("A8", 2) == 1:
            state.set("R36", _blank(additional.get('sample_year')))

        state.set("AD1", _blank(additional.get('unit_flag')))
        state.set("AC1", _blank(additional.get('dispersivity_flag')))
        state.set("AH28", _blank(additional.get('psb_loading')) if iwall else None)

        # v103: dropdown selections — push to state cells so state.push()
        # restores the UI widgets on Load Data.  Cell addresses follow
        # CELL_MAP in functions/state.py.
        if additional.get('model_version') is not None:
            mv = str(additional['model_version']).strip().lower()
            # A8 = 1 → Simple, 2 → Detailed (matches xlsm convention)
            state.set("A8", 2 if mv.startswith("detail") else 1)
        if additional.get('heterogeneity') is not None:
            state.set("A1", _blank(additional['heterogeneity']))
        if additional.get('lowk_media') is not None:
            state.set("K26", _blank(additional['lowk_media']))
        if additional.get('pfaa1') is not None:
            state.set("E38", _blank(additional['pfaa1']))
        if additional.get('pfaa2') is not None:
            # Empty / "None" string is meaningful for PFAA-2 — keep it
            state.set("G38", additional['pfaa2'] or "None")
        if additional.get('precursor1') is not None:
            state.set("K38", additional['precursor1'] or "None")
        if additional.get('precursor2') is not None:
            state.set("M38", additional['precursor2'] or "None")
        if additional.get('psb_kf_unit'):
            state.set("U24", additional['psb_kf_unit'])

        for i, name in enumerate(additional.get('monitoring_well_names', [])[:7]):
            state.set(f"U{34+i}", _blank(name))
        for i, conc in enumerate(additional.get('monitoring_well_concentrations', [])[:7]):
            if conc is None or (isinstance(conc, str) and conc.strip().lower() == 'none'):
                state.set(f"V{34+i}", None)
                state.set(f"X{34+i}", None)
            elif isinstance(conc, str) and ',' in conc:
                parts = conc.split(',', 1)
                state.set(f"V{34+i}", _blank(parts[0].strip()))
                state.set(f"X{34+i}", _blank(parts[1].strip()) if len(parts) > 1 else None)
            else:
                state.set(f"V{34+i}", conc)
