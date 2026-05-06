"""
restore_calibration.py — push the example's calibration values into
the §calibration quadrant.

Called from restore_from_example.run() after the main input.inp /
state push.  The values below are baked-in (NOT read from the
CalibrationTemplate_*.xlsx files at the project root, since those
files are *outputs* of the Save-Calibration-Data flow rather than
reliable inputs — relying on them would break the moment the user
clicks Save Calibration Data and overwrites the template, and would
also break for fresh installs that don't have those files yet).

The values are taken from the original Excel reference workbook
(reviewed once, transcribed here verbatim).  When new examples land,
add a new EXAMPLES entry rather than depending on a .xlsx.

Pushed into:
  • Step 3 weighting factors (v_calib_w[0..6])
  • Step 4 parameter ranges (v_calib_low / mid / high keyed by label)
  • Step 4 "Use this Parameter?" checkboxes (v_calib_chk)
  • Step 2 PFOS / Precursor checkboxes
"""
from __future__ import annotations


# ─────────────────────────────────────────────────────────────────────
# Hardcoded example calibration values per active sheet.
# Adding a new example = add a new key here.  Keys must match the
# value of app.active_sheet ("Simple" / "Detailed_2").
# ─────────────────────────────────────────────────────────────────────
EXAMPLES = {
    "Simple": {
        # Step 3: per-well weighting factors (7 wells in §10 order)
        "weights": [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0],
        # Step 2: PFOS / PFOS-None / Precursor / Precursor-None
        "step2": (True, False, False, False),
        # Step 4: dict of "param label" -> (low, mid, high)
        # Labels MUST exactly match main.py:_CALIB_PARAMS strings.
        "ranges": {
            "Hydraulic Conductivity (k)":                                  (4,   20, 100),
            "Hydraulic Gradient (i)":                                      (0.05, 0.1, 0.2),
            "Multiplier to PFAA-1 Source Concentration in #7 (czero(2,n))":(0.5,  1,   2),
        },
    },
    "Detailed_2": {
        # Detailed example — last well stays at 1.0 too; bump only via
        # Excel template if the user really wants downgradient
        # emphasis.  Was 2.0 in the legacy CalibrationTemplate_Detailed.
        "weights": [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0],
        "step2":   (True, False, True, False),
        "ranges": {
            "Hydraulic Conductivity (k)":                                       (4,   20, 100),
            "Hydraulic Gradient (i)":                                           (0.05, 0.1, 0.2),
            "Multiplier to PFAA-1 Source Concentration in #7 (czero(2,n))":     (0.5,  1,   2),
            "Multiplier to Precursor-1 Source Concentration in #7 (czero(1,n))":(0.5,  1,   2),
        },
    },
}


def _fmt(val):
    """Render a numeric value the way the §calibration cells expect:
    integers without trailing ".0", floats with :g."""
    if val is None:
        return ""
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, (int, float)):
        if float(val).is_integer():
            return f"{int(val)}"
        return f"{val:g}"
    return str(val)


def run(app, parent=None) -> bool:
    """Push the active-sheet's example calibration values into the
    app.  Returns True if any cell was touched, False otherwise.
    Best-effort — never raises."""
    sheet = getattr(app, "active_sheet", "Simple")
    ex = EXAMPLES.get(sheet)
    if ex is None:
        # Fall back to Simple if active_sheet is something unexpected
        ex = EXAMPLES["Simple"]

    # ── Step 3: per-well weights ─────────────────────────────────────
    weights = ex.get("weights", [])
    for i, w in enumerate(weights):
        if i < len(getattr(app, "v_calib_w", [])):
            try: app.v_calib_w[i].set(_fmt(w))
            except Exception: pass

    # ── Step 4: parameter ranges + checkbox state ────────────────────
    # Reset every Step 4 row first (clean slate), then set the rows
    # listed in the example.  Mirrors Excel parity: paste-example
    # wipes whatever the user had typed in.
    chk = getattr(app, "v_calib_chk",  [])
    lo_l = getattr(app, "v_calib_low",  [])
    mi_l = getattr(app, "v_calib_mid",  [])
    hi_l = getattr(app, "v_calib_high", [])

    for cv in chk:
        try: cv.set(False)
        except Exception: pass

    try:
        from main import _CALIB_PARAMS  # type: ignore
    except Exception:
        _CALIB_PARAMS = []

    ranges = ex.get("ranges", {})
    for i, label in enumerate(_CALIB_PARAMS):
        if i >= min(len(lo_l), len(mi_l), len(hi_l), len(chk)):
            break
        if label not in ranges:
            continue
        lo, mid, hi = ranges[label]
        # Mid is intentionally NOT set here — it's now driven LIVE
        # from the source cells in §3/§5 via the trace registered in
        # _build_calibration_panel.  Setting mid here would just be
        # overwritten on the next source-cell change anyway, and
        # would briefly show the wrong number in the black cell.
        for var, val in ((lo_l[i], lo), (hi_l[i], hi)):
            try: var.set(_fmt(val))
            except Exception: pass
        try: chk[i].set(True)
        except Exception: pass

    # ── Step 2: Calibrate-using checkboxes ───────────────────────────
    s2 = ex.get("step2", (False, 