"""
REMFluor-MD Model Input Dashboard — TKINTER variant.

Sister build to the PySide6 version at  _Python/main.py .  Same Excel/PDF
storyboard ground truth, same Figures/ folder, same VBA macro dispatch
table.  Lives at  REMFluorMD_v2.6/_Python/tkinter/main.py  so BASE_DIR
walks up two levels to reach the project root.

Visual / behavioral spec:
    - Top header bar:  dark teal  #074F69 with white italic title
    - Bottom bar:      pure black #000000
    - Body:            light gray #F2F2F2
    - ESTCP logo:      loaded from BASE_DIR/Figures/ (PNG/GIF/JPG via PIL)
    - Body font:       Calibri (Excel default)
    - "?" help boxes:  squared, sit flush-adjacent to their entry cells
    - Layout:          fixed proportions inside scrollable Canvas, both
                       horizontal and vertical scrollbars when window is
                       smaller than the canvas

Buttons call Python scripts / open Quarto files using the same paths as
the VBA macros.  Run with:
    cd _Python/tkinter
    python main.py
"""

# v99: PyInstaller multi-mode dispatcher
# ------------------------------------------------------------------
# In a frozen .exe build, sys.executable is the .exe itself.
# When run_model.py wants to launch the Plotly-Dash dashboard, it
# spawns subprocess.Popen([sys.executable, "--mode=dashboard", ...])
# instead of [python, "-u", script, ...].  That re-launches the .exe
# with the special flag — and this dispatcher intercepts it BEFORE
# the tkinter UI starts, routes into generate_dashboard.main(), and
# exits when the dashboard server stops.
#
# In dev mode (`python main.py`) sys.argv has no extra flag, so the
# `if` is False and the rest of the file runs normally.  This block
# is intentionally placed before any heavy imports (tkinter, pandas,
# etc.) so dashboard mode doesn't pay the tkinter init cost.
import sys as _sys
if len(_sys.argv) >= 2 and _sys.argv[1] == "--mode=dashboard":
    # Strip the flag so generate_dashboard.main() sees the
    # original argv (workbook_path, sheet_name) at indices 1, 2.
    _sys.argv = [_sys.argv[0]] + _sys.argv[2:]
    try:
        from functions import generate_dashboard
        generate_dashboard.main()
    except SystemExit:
        raise
    except Exception as _exc:
        # Best-effort error log — frozen subprocess has no console.
        import os as _os
        try:
            _here = _os.path.dirname(_os.path.abspath(_sys.executable))
            with open(_os.path.join(_here, "dashboard_dispatch_error.log"),
                      "w", encoding="utf-8") as _fp:
                import traceback as _tb
                _fp.write(f"Dashboard dispatch failed: {_exc}\n")
                _tb.print_exc(file=_fp)
        except Exception:
            pass
        raise
    _sys.exit(0)

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import subprocess
import os
import re
import sys
import glob
from datetime import datetime

# --- Pure-Python action modules (no .exe subprocess calls) ---
try:
    from functions.state import get_state
    from functions import (
        clear_for_restore,
        restore_from_example,
        restore_from_saved,
        visualize_saved,
        generate_input_file,
        popups_retardation,
        popups_source_remediation,
        popups_transformation,
        popups_numerical,
        popups_cellsize,
        popups_GWvelocity,
        popups_longevity,
        popups_mass_discharge_import,
        popups_calibration,
        popups_heterogeneity,
        run_model,
        cali_1,
    )
    _FUNCS_LOADED = True
except Exception as _e:
    print(f"Warning: could not load functions package: {_e}")
    _FUNCS_LOADED = False

# ─────────────────────────────────────────────────────────────────────────────
# HIGH-DPI / RETINA SHARPNESS FIX  (must run BEFORE creating any Tk widgets)
# ─────────────────────────────────────────────────────────────────────────────
# Without this, Windows treats the app as a 96-DPI legacy app and bitmap-
# stretches the entire window on scaled displays (125% / 150% / 175%), making
# text and images look blurry. With it, Tk renders at the display's native
# pixel resolution.
def _enable_high_dpi_awareness():
    if sys.platform != "win32":
        return
    try:
        import ctypes
        # Windows 8.1+: per-monitor DPI awareness (best)
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
            return
        except (AttributeError, OSError):
            pass
        # Windows Vista/7/8: system DPI awareness (good enough)
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except (AttributeError, OSError):
            pass
    except Exception:
        pass

_enable_high_dpi_awareness()

# ─────────────────────────────────────────────────────────────────────────────
# COLOR CONSTANTS  (sampled directly from the PDF storyboard pixels)
# ─────────────────────────────────────────────────────────────────────────────
BG_HEADER_BAR  = "#074F69"   # top bar – DARK TEAL (sampled from PDF)
BG_BOTTOM_BAR  = "#000000"   # bottom bar – PURE BLACK (sampled)
BG_MAIN        = "#F2F2F2"   # body background – LIGHT GRAY (sampled)
BG_WHITE       = "#FFFFFF"
BG_LEGEND_BAR  = "#FFFFFF"
BG_INPUT_BLUE  = "#FFFFFF"   # legend "Enter value directly"           – WHITE
BG_FORMULA     = "#DAE9F8"   # legend "Cell with formula or default..." – light blue
BG_PULLDOWN    = "#FBE2D5"   # legend "Pull Down Menu"                  – peach
BG_LOCKED      = "#000000"   # legend "Calculated value..."             – BLACK
FG_LOCKED      = "#FFFFFF"   # text color used inside black locked cells
BG_SECTION_HDR = "#DAEEF3"
BTN_FILL       = "#D9D9D9"   # standard gray button
BTN_FILL_BLUE  = "#BDD7EE"   # Run Model buttons (light blue)
BTN_FILL_GREEN = "#E6FFDC"   # Authors / Save Data button (light green)
BTN_FILL_RED   = "#FFD7D7"   # red-tinted button background

FG_TITLE       = "#FFFFFF"   # white title on dark teal
FG_SECTION     = "#0070C0"   # blue section headers (sampled)
FG_BTN_NAVY    = "#002060"
FG_BTN_RED     = "#FF0000"
FG_BTN_BLUE    = "#0070C0"
FG_BTN_GREEN   = "#385723"
FG_INPUT       = "#000000"
FG_GREY        = "#7F7F7F"
FG_YELLOW      = "#FFFF00"
FG_HELP        = "#FF0000"   # red "?" help link text

# ─────────────────────────────────────────────────────────────────────────────
# FONTS  (Calibri = Excel default; Tk falls back gracefully if unavailable)
# ─────────────────────────────────────────────────────────────────────────────
FG_TITLE_GREEN = "#DAF2D0"   # "REMFluor" portion of the title

# ─── FONTS as Tk NAMED FONTS ────────────────────────────────────────────────
# Storing fonts as named fonts (instead of tuples) lets us resize every widget
# at runtime by just calling .config(size=...) on the named font.  The string
# constants below are the names; the actual sizes are configured by
# REMFluorApp._init_named_fonts() after the Tk root is created.
FONT_TITLE     = "AppTitle"
FONT_VERSION   = "AppVersion"
FONT_SECTION   = "AppSection"
FONT_LABEL     = "AppLabel"
FONT_LABEL_B   = "AppLabelB"     # bold body label  (PFAA 1, etc.)
FONT_LABEL_I   = "AppLabelI"     # italic body label
FONT_LABEL_BI  = "AppLabelBI"    # bold italic
FONT_LABEL_SM  = "AppLabelSm"
FONT_LABEL_SMI = "AppLabelSmI"   # small italic
FONT_LABEL_XS  = "AppLabelXs"
FONT_INPUT     = "AppInput"
FONT_BTN       = "AppBtn"
FONT_BTN_SM    = "AppBtnSm"
FONT_BTN_LG    = "AppBtnLg"
FONT_BTN_CALIB = "AppBtnCalib"
FONT_HELP      = "AppHelp"
FONT_RADIO     = "AppRadio"

# Baseline (zoom = 1.0) sizes.  Title/version stay big as the user requested;
# every other size is dialled down so the dashboard fits on small displays.
# The user can press Ctrl++ / Ctrl+- / Ctrl+0 / Ctrl+wheel to zoom.
_FONT_DEFS = {
    # name              family            size  weight    slant
    FONT_TITLE:        ("Arial Narrow",   28,   "bold",   "roman"),
    FONT_VERSION:      ("Arial Narrow",   26,   "bold",   "roman"),
    FONT_SECTION:      ("Calibri",        11,   "bold",   "roman"),
    FONT_LABEL:        ("Calibri",         9,   "normal", "roman"),
    FONT_LABEL_B:      ("Calibri",         9,   "bold",   "roman"),
    FONT_LABEL_I:      ("Calibri",         9,   "normal", "italic"),
    FONT_LABEL_BI:     ("Calibri",         9,   "bold",   "italic"),
    FONT_LABEL_SM:     ("Calibri",         8,   "normal", "roman"),
    FONT_LABEL_SMI:    ("Calibri",         8,   "normal", "italic"),
    FONT_LABEL_XS:     ("Calibri",         8,   "normal", "roman"),
    FONT_INPUT:        ("Calibri",         9,   "normal", "roman"),
    FONT_BTN:          ("Calibri",        10,   "normal", "italic"),
    FONT_BTN_SM:       ("Calibri",         9,   "normal", "italic"),
    FONT_BTN_LG:       ("Calibri",        13,   "bold",   "italic"),
    FONT_BTN_CALIB:    ("Calibri",        10,   "bold",   "roman"),
    FONT_HELP:         ("Arial",           9,   "bold",   "italic"),
    FONT_RADIO:        ("Calibri",        12,   "normal", "roman"),
}


# ─────────────────────────────────────────────────────────────────────────────
# SCRIPT / FILE PATHS  (mirrored exactly from VBA macro targets)
# ─────────────────────────────────────────────────────────────────────────────
# v100: dual-base setup for PyInstaller --onefile builds.
#
# BASE_DIR  = writable runtime location (where input.inp, output.out,
#             dashboard_state.json, sidecar txt files, etc. live).
# BUNDLE_DIR = read-only assets location (where Figures/, Example/,
#              docs/, template.inp, remfluor_v8a.exe live).
#
# In dev mode both point at the project root (REMFluorMD_v2.6/).
# In a frozen --onefile build:
#   - BASE_DIR  → folder containing the .exe (persistent, user-visible)
#   - BUNDLE_DIR → sys._MEIPASS (temp unpack dir for bundled --add-data)
if getattr(sys, "frozen", False):
    BASE_DIR   = os.path.dirname(os.path.abspath(sys.executable))
    BUNDLE_DIR = getattr(sys, "_MEIPASS", BASE_DIR)
else:
    _HERE      = os.path.dirname(os.path.abspath(__file__))
    BASE_DIR   = os.path.abspath(os.path.join(_HERE, "..", ".."))
    BUNDLE_DIR = BASE_DIR


def _bundle_path(rel):
    """Return absolute path to a bundled read-only resource.

    Searches BASE_DIR first (so a user can drop an updated copy next
    to the .exe to override the bundled one), then falls back to
    BUNDLE_DIR (sys._MEIPASS in a frozen build).  Returns the BASE_DIR
    candidate if neither exists, so callers can still report a useful
    "missing file" error message.
    """
    cand = os.path.join(BASE_DIR, rel)
    if os.path.exists(cand):
        return cand
    cand2 = os.path.join(BUNDLE_DIR, rel)
    if os.path.exists(cand2):
        return cand2
    return cand


# Figures/ — read-only resource, prefer bundled location.
def _resolve_figures_dir():
    return _bundle_path("Figures")

FIGURES_DIR = _resolve_figures_dir()

def _html(*parts):
    return os.path.join(BASE_DIR, "docs", "_site", *parts)

HTML_APPENDIX = {
    "OpenAppendix_2_1_Relative": _html("appendix", "appendix_2_1.html"),
    "OpenAppendix_2_2_Relative": _html("appendix", "appendix_2_2.html"),
    "OpenAppendix_4_2_Relative": _html("appendix", "appendix_4_2.html"),
    "OpenAppendix_6_1_Relative": _html("appendix", "appendix_6_1.html"),
    "OpenAppendix_7_1_Relative": _html("appendix", "appendix_7_1.html"),
    "OpenAppendix_8_1_Relative": _html("appendix", "appendix_8_1.html"),
    "OpenAppendix_9_1_Relative":  _html("appendix", "appendix_9_1.html"),
    "OpenAppendix_11_1_Relative": _html("appendix", "appendix_11_1.html"),
}

HTML_CHICKLETS = {
    "OpenTable1":               ("Step1_SystemUnits.html", ""),
    "OpenTable2_1_XDirection":  ("Step2_ModelConfiguration.html", "model-size-in-direction-of-groundwater-flow-x-direction"),
    "OpenTable2_1_YDirection":  ("Step2_ModelConfiguration.html", "model-size-in-direction-of-groundwater-flow-y-direction"),
    "OpenTable2_1_ZDirection":  ("Step2_ModelConfiguration.html", "model-size-in-direction-of-groundwater-flow-z-direction"),
    "OpenTable2_2":             ("Step2_ModelConfiguration.html", "finite-difference-cell-size"),
    "OpenTable2_3":             ("Step2_ModelConfiguration.html", "source-width"),
    "OpenTable2_4":             ("Step2_ModelConfiguration.html", "thickness-of-source-below-water-table"),
    "OpenTable2_5":             ("Step2_ModelConfiguration.html", "starting-year-of-simulation"),
    "OpenTable2_6":             ("Step2_ModelConfiguration.html", "ending-year-of-simulation"),
    "OpenTable3_1":             ("Step3_GroundwaterDarcyVelocity.html", "groundwater-darcy-velocity-vd"),
    "OpenTable3_2":             ("Step3_GroundwaterDarcyVelocity.html", "transmissive-zone-effective-porosity"),
    "OpenTable4_1":             ("Step4_HydrogeologicSettingAndMatrixDiffusion.html", "unconsolidated-aquifers---low-k-media-details"),
    "OpenTable4_2":             ("Step4_HydrogeologicSettingAndMatrixDiffusion.html", "low-k-zone-total-porosity"),
    "OpenTable4_3":             ("Step4_HydrogeologicSettingAndMatrixDiffusion.html", "low-k-zone-tortuosity"),
    "OpenTable5_1":             ("Step5_PFASTransportProperties.html", "constituent"),
    "OpenTable5_2":             ("Step5_PFASTransportProperties.html", "retardation-factor-calculations"),
    "OpenTable5_3":             ("Step5_PFASTransportProperties.html", "retardation-factor-in-t-zone-transmissive-zone"),
    "OpenTable5_4":             ("Step5_PFASTransportProperties.html", "retardation-factor-in-low-k-media"),
    "OpenTable5_5":             ("Step5_PFASTransportProperties.html", "detailed-model-only-precursor-transformation-to-pfaas"),
    "OpenTable5_6":             ("Step5_PFASTransportProperties.html", "microbial-yield-factor"),
    "OpenTable5_7":             ("Step5_PFASTransportProperties.html", "molecular-diffusion-coefficient-in-free-water"),
    "OpenTable6_0":             ("Step6_PlumeTransport.html", ""),
    "OpenTable6_1":             ("Step6_PlumeTransport.html", "longitudinal-dispersivity"),
    "OpenTable6_2":             ("Step6_PlumeTransport.html", "transverse-dispersivity"),
    "OpenTable6_3":             ("Step6_PlumeTransport.html", "vertical-dispersivity"),
    "OpenTable7_1":             ("Step7_PFASSourceTerm.html", "initial-source-concentration"),
    "OpenTable8_1":             ("Step8_SourceRemediation.html", "percent-source-mass-removed-by-remediation"),
    "OpenTable8_2":             ("Step8_SourceRemediation.html", "remediation-started-in-year"),
    "OpenTable8_3":             ("Step8_SourceRemediation.html", "remediation-ended-in-year"),
    "OpenTable9_0":             ("Step9_PlumeRemediationPSB.html", ""),
    "OpenTable9_1":             ("Step9_PlumeRemediationPSB.html", "psb-freundlich-exponent-a"),
    "OpenTable9_2":             ("Step9_PlumeRemediationPSB.html", "psb-freundlich-kf"),
    "OpenTable9_3":             ("Step9_PlumeRemediationPSB.html", "year-psb-barrier-installed"),
    "OpenTable9_4":             ("Step9_PlumeRemediationPSB.html", "total-width-of-psb-in-x-direction"),
    "OpenTable9_5":             ("Step9_PlumeRemediationPSB.html", "psb-loading-fcac"),
    "OpenTable10_1":            ("Step10_FieldDataToCalibrate.html", "sample-year"),
    "OpenTable10_2":            ("Step10_FieldDataToCalibrate.html", "monitoring-well-name"),
    "OpenTable10_3":            ("Step10_FieldDataToCalibrate.html", "concentration-measured"),
    "OpenTable10_4":            ("Step10_FieldDataToCalibrate.html", "distance-from-source"),
    "OpenTable11_1":            ("Step11_ModelingParameters.html", "see-results-every"),
}

# ─────────────────────────────────────────────────────────────────────────────
# ACTION DISPATCHER
# ─────────────────────────────────────────────────────────────────────────────
def _open_html(path_or_url: str, wait: bool = False):
    if sys.platform == "win32":
        url = "file:///" + path_or_url.replace("\\", "/")
        subprocess.Popen(["cmd", "/c", "start", "", url],
                         shell=False, creationflags=subprocess.CREATE_NO_WINDOW
                         if hasattr(subprocess, "CREATE_NO_WINDOW") else 0)
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path_or_url])
    else:
        subprocess.Popen(["xdg-open", path_or_url])


_app_ref = None


def _ensure_state_work_dir():
    """Set state work_dir to BASE_DIR if not already set."""
    if _FUNCS_LOADED and _app_ref is not None:
        st = get_state()
        if not st.work_dir:
            st.work_dir = BASE_DIR
        # v100: also expose BUNDLE_DIR (read-only assets) so
        # restore_from_example, popup _docs_root helpers, and others
        # can resolve bundled Example/ + docs/ folders correctly in
        # a frozen --onefile build.
        try:
            st.bundle_dir = BUNDLE_DIR
        except Exception:
            pass


# ─────────────────────────────────────────────────────────────────────────────
# Calibration sidecar helpers — pure-Python replacements for the dist/.exe
# scripts (export_calibration_data.exe, restore_from_optimal.exe).  Keeps the
# app fully standalone (no xlwings, no Excel COM, no .exe spawn).
# ─────────────────────────────────────────────────────────────────────────────
_CALIB_PARAMS = [
    "Source Start Year (nt)",
    "Hydraulic Conductivity (k)",
    "Hydraulic Gradient (i)",
    "Effective Porosity (porf)",
    "Transmissive Fraction of Model (volfrac)",
    "Average Diffusion Length (difflen)",
    "Retardation Factor of PFAA-1 (ock(2))",
    "Retardation Factor of PFAA-2 (ock(4))",
    "Longitudinal Dispersivity (alphax (m))",
    "Multiplier to PFAA-1 Source Concentration in #7 (czero(2,n))",
    "Multiplier to PFAA-2 Source Concentration in #7 (czero(4,n))",
    "First order decay rate coefficient for Precursors-1 (decayf(1))",
    "First order decay rate coefficient for Precursors-2 (decayf(3))",
    "Retardation Factor of Precursors-1 (ock(1))",
    "Retardation Factor of Precursors-2 (ock(3))",
    "Multiplier to Precursor-1 Source Concentration in #7 (czero(1,n))",
    "Multiplier to Precursor-2 Source Concentration in #7 (czero(3,n))",
]


def _save_calibration_inputs(app):
    """Write calibration_inputs.txt next to input.inp.  Format mirrors
    what the Excel macro emitted — one ranges block + one weights block
    + the iteration count — so a future Python optimizer (or the
    legacy cali_1.exe) can read it without changes.

    v92: preserve any existing "Excel File Path:" line written by the
    §10 button (popups_calibration) — that path is the link to the
    user's calibration .xlsx and the dashboard / cali_1 both depend
    on it.  Without preservation the §10 button and §calibration "1.
    Save" button would clobber each other."""
    out = os.path.join(BASE_DIR, "calibration_inputs.txt")
    n_iter = getattr(app, "v_n_iter", None)
    n_iter_val = n_iter.get() if n_iter is not None else "50"
    # Preserve Excel File Path line from any prior write
    existing_xlsx_line = ""
    if os.path.exists(out):
        try:
            with open(out, "r", encoding="utf-8") as fp:
                for ln in fp:
                    if ln.lstrip().startswith("Excel File Path:"):
                        existing_xlsx_line = ln.rstrip("\n") + "\n"
                        break
        except Exception:
            pass
    with open(out, "w", encoding="utf-8") as fp:
        if existing_xlsx_line:
            fp.write(existing_xlsx_line)
        fp.write("REMFluor-MD Calibration Inputs\n")
        fp.write("=" * 50 + "\n\n")
        fp.write(f"Iterations: {n_iter_val}\n")
        fp.write(f"PFAS species: {app.v_pfaa1.get()} / "
                 f"{app.v_pfaa2.get()}\n")
        fp.write(f"Calibrate against PFAA-1: "
                 f"{bool(app.v_calib_pfoa.get())}\n")
        fp.write(f"Calibrate against PFAA-2: "
                 f"{bool(app.v_calib_none.get())}\n\n")

        fp.write("Step 3 — Monitoring well weighting factors\n")
        fp.write("-" * 50 + "\n")
        fp.write("name,distance,weight\n")
        names  = getattr(app, "v_mw_names", [])
        dists  = getattr(app, "v_mw_dist",  [])
        weights = getattr(app, "v_calib_w", [])
        for i in range(min(len(names), len(dists), len(weights))):
            fp.write(f"{names[i].get()},{dists[i].get()},"
                     f"{weights[i].get()}\n")

        fp.write("\nStep 4 — Calibration parameter ranges\n")
        fp.write("-" * 50 + "\n")
        fp.write("use,parameter,lo,mid,hi\n")
        chk  = getattr(app, "v_calib_chk",  [])
        lo   = getattr(app, "v_calib_low",  [])
        mid  = getattr(app, "v_calib_mid",  [])
        hi   = getattr(app, "v_calib_high", [])
        for i in range(min(len(chk), len(lo), len(mid), len(hi),
                           len(_CALIB_PARAMS))):
            fp.write(f"{int(bool(chk[i].get()))},"
                     f"{_CALIB_PARAMS[i]},"
                     f"{lo[i].get()},{mid[i].get()},{hi[i].get()}\n")
    return out


def _load_optimal_model(app, path):
    """Read a previously-saved optimal_model.txt and push the saved
    state back into the app.  Supports both formats:

      v1 (legacy):
        use=False lo=... mid=... hi=...     (one row per line, by index)

      v2 (current):
        src.<varname>=<value>                (source-cell snapshot)
        row[<i>]
          label=<row label>
          use=<bool>
          lo=<value>
          mid=<value>
          hi=<value>

    v2 matches by label so reordering rows in _CALIB_PARAMS doesn't
    silently desync past saves.  Returns the number of rows restored.

    To prevent the source-cell traces from clobbering the just-loaded
    Mid values, the helper sets app._calib_pushing = True around the
    writes.
    """
    chk  = getattr(app, "v_calib_chk",  [])
    lo   = getattr(app, "v_calib_low",  [])
    mid  = getattr(app, "v_calib_mid",  [])
    hi   = getattr(app, "v_calib_high", [])

    src_snapshot = {}        # v2: src.<varname>=<value>
    s7_snapshot  = {}        # v3: s7.<col>[<i>]=<value>  (col = pfaa1 / pfaa2 / pre1 / pre2)
    rows_by_label = {}       # v2: label → (use, lo, mid, hi)
    legacy_rows  = []        # v1: ordered list of (use, lo, mid, hi)

    cur = None    # current v2 row buffer
    pat_v1 = re.compile(
        r"use=(\S+)\s+lo=(\S+)\s+mid=(\S+)\s+hi=(\S+)")
    pat_s7 = re.compile(r"s7\.(\w+)\[(\d+)\]")
    try:
        with open(path, "r", encoding="utf-8") as fp:
            for line in fp:
                s = line.strip()
                if not s:
                    if cur is not None:
                        rows_by_label[cur.get("label", "")] = cur
                        cur = None
                    continue
                if s.startswith("src."):
                    k, _, v = s.partition("=")
                    src_snapshot[k[len("src."):]] = v
                    continue
                if s.startswith("s7."):
                    k, _, v = s.partition("=")
                    m = pat_s7.match(k)
                    if m:
                        col, idx_s = m.group(1), m.group(2)
                        s7_snapshot[(col, int(idx_s))] = v
                    continue
                if s.startswith("row["):
                    if cur is not None:
                        rows_by_label[cur.get("label", "")] = cur
                    cur = {}
                    continue
                if cur is not None and "=" in s:
                    k, _, v = s.partition("=")
                    cur[k.strip()] = v.strip()
                    continue
                # Legacy v1 fallback — single line with all four fields
                m = pat_v1.search(line)
                if m:
                    legacy_rows.append(m.groups())
        if cur is not None:
            rows_by_label[cur.get("label", "")] = cur
    except Exception as exc:
        raise

    # Suppress _refresh_all_mids during the push so source-cell
    # traces don't clobber freshly-loaded Mid values.
    setattr(app, "_calib_pushing", True)
    try:
        n = 0
        # v2 — match by label
        if rows_by_label:
            for i, label in enumerate(_CALIB_PARAMS):
                if i >= min(len(chk), len(lo), len(mid), len(hi)):
                    break
                row = rows_by_label.get(label)
                if not row:
                    continue
                try: chk[i].set(str(row.get("use", "")).lower() == "true")
                except Exception: pass
                lo[i].set(row.get("lo", ""))
                mid[i].set(row.get("mid", ""))
                hi[i].set(row.get("hi", ""))
                n += 1
        else:
            # v1 — index-based (legacy compatibility)
            for i, (use_s, lo_s, mid_s, hi_s) in enumerate(legacy_rows):
                if i >= min(len(chk), len(lo), len(mid), len(hi)):
                    break
                try: chk[i].set(use_s.lower() == "true")
                except Exception: pass
                lo[i].set(lo_s); mid[i].set(mid_s); hi[i].set(hi_s)
                n += 1

        # Restore §3 / §5 / §6 source snapshot if present (v2 only).
        for vname, val in src_snapshot.items():
            target = getattr(app, vname, None)
            if target is None:
                continue
            try: target.set(val)
            except Exception: pass

        # Restore §7 source-zone concentration grid if present (v3+).
        # This carries the post-multiplier values directly so we
        # don't have to re-apply the calibration multiplier (which
        # would compound on subsequent loads).
        if s7_snapshot:
            col_to_attr = {
                "pfaa1": "v_src_pfaa1",
                "pfaa2": "v_src_pfaa2",
                "pre1":  "v_src_pre1",
                "pre2":  "v_src_pre2",
            }
            for (col, idx), val in s7_snapshot.items():
                attr = col_to_attr.get(col)
                if attr is None:
                    continue
                lst = getattr(app, attr, None)
                if not lst or idx >= len(lst):
                    continue
                try: lst[idx].set(val)
                except Exception: pass
    finally:
        setattr(app, "_calib_pushing", False)
    return n


def run_script(macro_name, extra_args=None):
    _ensure_state_work_dir()

    # ── Standalone Python replacements ──────────────────────────────────
    if _FUNCS_LOADED and _app_ref is not None:

        if macro_name == "Clear_Data":
            deleted = clear_for_restore.run(_app_ref)
            msg = (f"Cleared {len(deleted)} file(s)." if deleted
                   else "Cells cleared. No .txt files found.")
            messagebox.showinfo("Clear Data", msg)
            return

        if macro_name == "Paste_Example":
            # Step 1: clear, Step 2: restore from example
            clear_for_restore.run(_app_ref)
            restore_from_example.run(_app_ref)
            return

        if macro_name == "Load_Data":
            # Step 1: clear, Step 2: restore from saved folder
            clear_for_restore.run(_app_ref)
            restore_from_saved.run(_app_ref)
            return

        if macro_name == "Visualize_Saved_Results":
            # §11: pick a folder with a previous run's output files,
            # stage them into the project dir, and open the dashboard
            # WITHOUT re-running the solver.
            try:
                visualize_saved.run(_app_ref)
            except Exception as exc:
                messagebox.showerror(
                    "Visualize Saved Results",
                    f"Could not visualize saved results:\n{exc}")
            return

        if macro_name == "CalculrateRetardationFactors":
            # In-app port of Source_Py/popups_retardation.py — no .exe, no
            # xlwings/openpyxl.  Opens a Toplevel popup, persists Koc/foc/
            # bulk density to retardation_inputs.txt, then writes computed
            # R values back into the section-5 StringVars.
            try:
                popups_retardation.run(_app_ref)
            except Exception as exc:
                messagebox.showerror(
                    "Calculate Retardation Factors",
                    f"Could not open retardation popup:\n{exc}")
                return
            # Trace recompute as a safety net (e.g. if user only edited foc)
            try:
                _app_ref._on_pfaa_change()
            except Exception:
                pass
            return

        if macro_name == "Save_Data":
            # v102: prompt user for destination folder + copy input.inp
            # AND every sidecar (.txt / .json) so the saved folder can
            # be reloaded with Load Data later.  Also writes a fresh
            # store_info_additional_input.txt for fields not in input.inp
            # (site name, date, sample year, sample event, source
            # treatment %, monitoring-well names + concentrations).
            from tkinter import filedialog as _fd
            from functions.state import get_state as _gs, INPUT_TXT_FILES as _SIDECARS
            try:
                # Step 1: build input.inp in the live work_dir (always)
                if not generate_input_file.run(_app_ref):
                    return
                state = _gs()
                work_dir = state.work_dir or os.getcwd()
                # Step 2: ask user where to save
                dst = _fd.askdirectory(
                    title="Choose folder to save model setup",
                    initialdir=work_dir,
                    mustexist=False,
                )
                if not dst:
                    return  # user cancelled
                # v106: warn before overwriting an existing saved model so
                # the user doesn't clobber a different model by mistake.
                # (Load stays read-only; only Save writes — so the original
                # loaded folder is never touched unless saved over here.)
                if os.path.exists(os.path.join(dst, "input.inp")):
                    if not messagebox.askyesno(
                        "Overwrite Existing Model?",
                        "This folder already contains a saved model:\n"
                        f"{dst}\n\n"
                        "Saving will overwrite its input.inp, sidecars and "
                        "results.\n\nOverwrite?\n"
                        "(Choose No to cancel, then pick a different folder "
                        "to keep both models.)"):
                        return
                os.makedirs(dst, exist_ok=True)
                # Step 3: write store_info_additional_input.txt next to
                # input.inp (UI-only fields not in the .inp).
                # v102 FIXES:
                #   - "Thickness:" must be SOURCE thickness (v_sw_thick →
                #     E16), NOT model depth (v_z_size → E13).  The parser
                #     in inp_parser.py maps "Thickness" to E16.  Writing
                #     v_z_size here caused Save→Load to overwrite the
                #     user's source thickness with the model depth.
                #   - Added explicit "Source Width:" + "Model X/Y/Z Size:"
                #     lines so round-trip is exact (previously source
                #     width was recovered from lysource*dy*2 in input.inp,
                #     which is lossy when (sw_width / 2*dy) isn't an int).
                try:
                    # Helper: pull a StringVar's value, falling back to a
                    # default if the attribute is missing entirely (some
                    # vars are Detailed-only and absent in Simple mode).
                    def _gv(name, default=""):
                        v = getattr(_app_ref, name, None)
                        try: return v.get() if v is not None else default
                        except Exception: return default
                    # Map model_version: "Simple" / "Detailed_2" → file
                    _mv = "Detailed" if getattr(_app_ref, "active_sheet",
                                                 "Simple") == "Detailed_2" \
                                       else "Simple"
                    add_lines = [
                        "Additional Information Not in input.inp",
                        "=" * 50, "",
                        f"Site Location and ID:,{_app_ref.v_site.get()}",
                        f"Date:,{_app_ref.v_date.get()}",
                        f"Thickness:,{_app_ref.v_sw_thick.get()}",
                        f"Source Width:,{_app_ref.v_sw_width.get()}",
                        f"Model X Size:,{_app_ref.v_x_size.get()}",
                        f"Model Y Size:,{_app_ref.v_y_size.get()}",
                        f"Model Z Size:,{_app_ref.v_z_size.get()}",
                        f"Start Year:,{_app_ref.v_yr_start.get()}",
                        f"End Year:,{_app_ref.v_yr_end.get()}", "",
                        f"Source Treatment Start Year:,{_app_ref.v_src_rem_yr.get() or 'None'}",
                        f"Source Concentration Reduction:,{_app_ref.v_src_conc_red.get() or 'None'}", "",
                        f"Sample Year:,{_app_ref.v_sample_yr.get()}", "",
                        # v106: removed the vestigial "Sample Event" line —
                        # there is no v_event field in the app, so it only
                        # ever wrote a blank and was never read back on load.
                        f"Unit Flag (AD1):,"
                        f"{'1' if _app_ref.v_units.get()=='feet' else '2'}",
                        f"Dispersivity Flag (AC1):,2", "",
                        # Dropdown selections — needed so Load Data restores
                        # the full UI state (radios + dropdowns), not just
                        # the numeric cells.  Each line is read by
                        # inp_parser.parse_additional_info().
                        f"Model Version:,{_mv}",
                        f"Units:,{_app_ref.v_units.get()}",
                        f"Heterogeneity:,{_gv('v_het', 'Medium')}",
                        f"Low-k Zone Media:,{_gv('v_lowk_media', '')}",
                        f"PFAA 1:,{_gv('v_pfaa1', 'PFOS')}",
                        f"PFAA 2:,{_gv('v_pfaa2', 'None')}",
                        f"Precursor 1:,{_gv('v_pfaa3', 'None')}",
                        f"Precursor 2:,{_gv('v_pfaa4', 'None')}",
                        f"PSB Kf Unit:,{_gv('v_psb_kf_unit', '')}", "",
                        f"PSB Loading (AH28):,{_app_ref.v_psb_load.get() or 'None'}", "",
                        f"Bulk Darcy Velocity (vd):,{_app_ref.v_darcy.get()}",
                        f"Effective Porosity (porf):,{_app_ref.v_porf.get()}", "",
                        "Monitoring Well Names (Simple Version):",
                    ]
                    for i, nm in enumerate(_app_ref.v_mw_names, 1):
                        add_lines.append(f"Well {i}:,{nm.get()}")
                    add_lines.append("")
                    add_lines.append("Monitoring Well Concentrations (Simple Version):")
                    # v106: save BOTH PFAA-1 (v_mw_conc) and PFAA-2
                    # (v_mw_conc2) §10 concentrations as "Well i:,c1,c2".
                    # inp_parser.parse_additional_info splits on the comma
                    # into V34.. (PFAA1) and X34.. (PFAA2); previously only
                    # c1 was written so the PFAA-2 column was lost on Load.
                    _conc2 = getattr(_app_ref, "v_mw_conc2", [])
                    for i, cv in enumerate(_app_ref.v_mw_conc, 1):
                        c2 = ""
                        if i - 1 < len(_conc2):
                            try: c2 = _conc2[i - 1].get()
                            except Exception: c2 = ""
                        add_lines.append(f"Well {i}:,{cv.get()},{c2}")
                    add_path = os.path.join(dst, "store_info_additional_input.txt")
                    with open(add_path, "w", encoding="utf-8") as fh:
                        fh.write("\n".join(add_lines) + "\n")
                except Exception as exc:
                    print(f"[Save_Data] additional info write failed: {exc}")
                # Step 3b (v106): write dedicated §6 dispersivity + §9 PSB
                # sidecars so Save → Load restores those values EXACTLY
                # (custom "Enter Your Own Value" dispersivity + raw PSB
                # Freundlich Kf with its original unit dropdown).  Written
                # to work_dir (picked up by the generic copy below) AND
                # directly to dst when dst differs.
                try:
                    from functions import sidecars as _sc
                    _sc.write_dispersivity(_app_ref, work_dir)
                    _sc.write_psb(_app_ref, work_dir)
                    if os.path.normcase(os.path.realpath(dst)) != \
                       os.path.normcase(os.path.realpath(work_dir)):
                        _sc.write_dispersivity(_app_ref, dst)
                        _sc.write_psb(_app_ref, dst)
                except Exception as exc:
                    print(f"[Save_Data] sidecar write failed: {exc}")
                # Step 4: copy input.inp + every sidecar listed in
                # INPUT_TXT_FILES from work_dir to dst.  v102: handle
                # three edge cases that broke the previous version:
                #   (a) user picks the SAME folder as work_dir →
                #       shutil.copy2 raised SameFileError because src
                #       and dst resolve to the same path.  Skip the
                #       copy in that case — files are already there.
                #   (b) destination file is read-only (Windows often
                #       drops the read-only attribute after the .exe
                #       run wrote it).  chmod 0o666 before overwriting.
                #   (c) any per-file failure shouldn't abort the whole
                #       save — keep going, log per-file errors.
                import shutil
                # Resolve real paths once so the same-folder check is
                # robust to slashes and case (Windows).
                try:
                    src_dir_real = os.path.realpath(work_dir)
                    dst_dir_real = os.path.realpath(dst)
                except Exception:
                    src_dir_real = work_dir
                    dst_dir_real = dst
                same_folder = (os.path.normcase(src_dir_real) ==
                               os.path.normcase(dst_dir_real))

                def _copy_one(src, dst_path):
                    """Copy src → dst_path with overwrite + same-file
                    safety.  Returns True on success / no-op, False on
                    failure."""
                    if not os.path.exists(src):
                        return False
                    try:
                        # Same file (same folder, same name) → no-op.
                        if (os.path.exists(dst_path)
                                and os.path.samefile(src, dst_path)):
                            return True
                    except Exception:
                        pass
                    # Make destination writable if it already exists.
                    if os.path.exists(dst_path):
                        try: os.chmod(dst_path, 0o666)
                        except Exception: pass
                    try:
                        shutil.copy2(src, dst_path)
                        return True
                    except shutil.SameFileError:
                        return True   # treat as success
                    except Exception as exc:
                        print(f"[Save_Data] copy failed "
                              f"{os.path.basename(src)}: {exc}")
                        return False

                copied = []
                # input.inp
                src_inp = os.path.join(work_dir, "input.inp")
                dst_inp = os.path.join(dst, "input.inp")
                if same_folder and os.path.exists(src_inp):
                    copied.append("input.inp")        # already there
                elif _copy_one(src_inp, dst_inp):
                    copied.append("input.inp")
                # Every sidecar (.txt / .json)
                for fname in _SIDECARS:
                    src = os.path.join(work_dir, fname)
                    if not os.path.exists(src):
                        continue
                    if same_folder:
                        copied.append(fname)          # already there
                        continue
                    if _copy_one(src, os.path.join(dst, fname)):
                        copied.append(fname)
                # v106: also save the model RESULT files (.out + the
                # dashboard JSON) so the saved folder can later be opened
                # with §11 "Visualize Saved Results".  These only exist
                # after a Run Model; skip silently if the user hasn't run
                # yet.  obs_well*.out is globbed.
                try:
                    import glob as _glob
                    from functions import run_model as _rm
                    _result_names = list(_rm.DASHBOARD_RESULT_FILES)
                    for _pat in _rm.DASHBOARD_RESULT_GLOBS:
                        _result_names += [os.path.basename(p) for p in
                                          _glob.glob(os.path.join(work_dir, _pat))]
                    for fname in _result_names:
                        src = os.path.join(work_dir, fname)
                        if not os.path.exists(src):
                            continue
                        if same_folder:
                            if fname not in copied:
                                copied.append(fname)
                            continue
                        if _copy_one(src, os.path.join(dst, fname)) \
                                and fname not in copied:
                            copied.append(fname)
                except Exception as exc:
                    print(f"[Save_Data] result-file copy failed: {exc}")
                # v106: adopt the saved folder as the ACTIVE working folder
                # so subsequent Run Model writes its input.inp + .out results
                # here, and Visualize Results reads from here — one folder
                # per model.  generate_input_file / popups / run_model all
                # key off state.work_dir, so this is the single switch.
                try:
                    state.work_dir = dst
                except Exception as exc:
                    print(f"[Save_Data] could not set work_dir: {exc}")
                # User-facing summary
                same_note = ("\n(destination is the same as the work "
                             "folder — existing files left in place)"
                             if same_folder else "")
                messagebox.showinfo(
                    "Save Data",
                    f"Saved {len(copied)} file(s) to:\n{dst}{same_note}\n\n"
                    + "\n".join(f"  {f}" for f in copied)
                    + "\n\nThis is now your active model folder — Run Model "
                      "will write results here, and Visualize Results will "
                      "read from here."
                )
            except Exception as exc:
                messagebox.showerror("Save Data",
                                     f"Could not save:\n{exc}")
            return

        if macro_name == "RunPythonScript":
            # Full pipeline: build input.inp, run remfluor_v8a.exe with
            # shell redirection (< input.inp > output.out), show a runtime
            # clock, and launch the Plotly-Dash dashboard in the user's
            # default browser when the model finishes.
            try:
                run_model.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("Run Model",
                                     f"Run Model failed:\n{exc}")
            return

    if macro_name == "Save_Data_Calibration":
        # Pure-Python: build input.inp from current state, then write a
        # calibration_inputs.txt sidecar with the chosen parameter
        # ranges + weighting factors.  No xlwings / .exe required.
        if _FUNCS_LOADED and _app_ref is not None:
            try:
                ok = generate_input_file.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("Save Calibration Data",
                                     f"input.inp build failed:\n{exc}")
                return
            if not ok:
                return
            try:
                _save_calibration_inputs(_app_ref)
                messagebox.showinfo("Save Calibration Data",
                    "input.inp + calibration_inputs.txt written to:\n"
                    f"{BASE_DIR}")
            except Exception as exc:
                messagebox.showerror("Save Calibration Data",
                                     f"calibration_inputs.txt write "
                                     f"failed:\n{exc}")
        return

    if macro_name == "OpenAppendix_2_1_Relative":
        _open_html(HTML_APPENDIX[macro_name])
        if _FUNCS_LOADED and _app_ref is not None:
            try:
                popups_cellsize.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("Cell Size", f"Popup failed:\n{exc}")
        return

    if macro_name in HTML_APPENDIX:
        _open_html(HTML_APPENDIX[macro_name])
        return

    if macro_name in HTML_CHICKLETS:
        html_file, anchor = HTML_CHICKLETS[macro_name]
        url = _html("data_chicklets", html_file)
        if anchor:
            url = url + "#" + anchor
        _open_html(url)
        return

    if macro_name == "Show_Visualization":
        if _app_ref:
            _app_ref.show_calibration_panel()
        return

    if macro_name == "Show_MainInterface":
        if _app_ref:
            _app_ref.show_main_panel()
        return

    if macro_name == "ScrollToCalibration":
        # §11 "Run Model with Auto-Calibration" button — DON'T launch
        # the optimizer here.  The user wants to be taken to the
        # §calibration quadrant first so they can fill out Step 2-4
        # before kicking it off (the optimizer button lives there
        # too: "2. Run Machine Based Calibration").
        if _app_ref:
            try:
                _app_ref.scroll_to_calibration()
            except Exception:
                pass
        return

    if macro_name == "SourceOption1":
        # v96: previously only copied PFAA-1 / PFAA-2 values; the
        # precursor columns (Detailed-only) were left at whatever the
        # user had entered.  Now ALL four §7 columns get the
        # "constant" treatment when this button is clicked.
        if _app_ref:
            for col_attr in ("v_src_pfaa1", "v_src_pfaa2",
                             "v_src_pre1",  "v_src_pre2"):
                lst = getattr(_app_ref, col_attr, None)
                if not lst:
                    continue
                try:
                    first = lst[0].get()
                except Exception:
                    continue
                for v in lst[1:]:
                    try: v.set(first)
                    except Exception: pass
        return

    if macro_name == "See_Calibration_Data":
        csv_path = os.path.join(BASE_DIR, "run_history.csv")
        if os.path.exists(csv_path):
            _open_html(csv_path)
        else:
            messagebox.showinfo("Not Found",
                                f"run_history.csv not found in:\n{BASE_DIR}")
        return

    # ── Calibration helper buttons (DefaultRanges, Run_Optimal_Model,
    # Save_Optimal_Model, Help_Calibration) — pure-Python stubs.  They
    # don't have dist/popups_*.exe equivalents in the legacy codebase.
    if macro_name == "DefaultRanges":
        if _app_ref is None:
            return
        # Default literature low/high = mid * 0.5 / mid * 2 (per the
        # button caption: "Experience (default +/- x2)")
        try:
            for cv, lo, mid, hi in zip(_app_ref.v_calib_chk,
                                       _app_ref.v_calib_low,
                                       _app_ref.v_calib_mid,
                                       _app_ref.v_calib_high):
                try:
                    m = float(str(mid.get()).replace(",", ""))
                except (ValueError, TypeError):
                    continue
                lo.set(f"{m * 0.5:g}")
                hi.set(f"{m * 2.0:g}")
        except Exception as exc:
            messagebox.showerror("Default Ranges",
                                 f"Could not fill defaults:\n{exc}")
        return

    if macro_name == "Run_Optimal_Model":
        # Run the model with the current (presumed-optimal) inputs —
        # equivalent to clicking the main "Run Model" button.
        if _FUNCS_LOADED and _app_ref is not None:
            try:
                run_model.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("Run Optimal Model",
                                     f"Run failed:\n{exc}")
        return

    if macro_name == "Run_Machine_Based_Calibration":
        # Standalone Python calibration loop — DDS optimizer over the
        # checked rows in §Step 4.  Replaces the legacy cali_1.exe
        # pipeline (which depended on xlwings + an Excel template +
        # the pandas/sklearn stack).  See functions/cali_1.py for the
        # full implementation.
        if _FUNCS_LOADED and _app_ref is not None:
            # First save the calibration inputs sidecar so the user
            # can inspect / version the parameter ranges they ran.
            try:
                ok = generate_input_file.run(_app_ref)
                if ok:
                    _save_calibration_inputs(_app_ref)
            except Exception as exc:
                print(f"[cali] sidecar save failed: {exc}")
            # Now kick off the optimizer.  Returns immediately; the
            # actual loop runs in a worker thread inside cali_1.run().
            try:
                cali_1.run(_app_ref)
            except Exception as exc:
                import traceback
                traceback.print_exc()
                messagebox.showerror(
                    "Run Machine Based Calibration",
                    f"Calibration failed to start:\n{exc}")
        return

    if macro_name == "Load_Optimal_Data":
        # Pure-Python: read optimal_model.txt (written by Save Optimal
        # Model) and push the saved state back into the app.  If the
        # file doesn't exist, tell the user.
        if _app_ref is None:
            return
        path = os.path.join(BASE_DIR, "optimal_model.txt")
        if not os.path.exists(path):
            messagebox.showinfo(
                "Load Optimal Data",
                f"No optimal_model.txt found in:\n{BASE_DIR}\n\n"
                "Click '6. Save Optimal Model' first to create it.")
            return
        # Delete sidecar files first.  Otherwise _src_K / _src_i /
        # _src_volfrac / _src_difflen would re-read those files on
        # the next refresh and clobber the values we're about to
        # load.  Optimal_model.txt now carries the source-cell
        # snapshot directly, so the sidecars are no longer needed.
        for sidecar in ("gwvelocity_inputs.txt",
                        "heterogeneity_inputs.txt"):
            sp = os.path.join(BASE_DIR, sidecar)
            if os.path.exists(sp):
                try: os.remove(sp)
                except Exception: pass
        try:
            n = _load_optimal_model(_app_ref, path)
            # Order matters here:
            #  (1) Write gwvelocity_inputs.txt FIRST from the freshly-
            #      loaded mid[K] + mid[i].  If we let _push run before
            #      this, _push's trailing refresh would call _src_K
            #      which falls back to v_darcy — and v_darcy at that
            #      moment is the K×i product (a small number), so
            #      mid[K] would be clobbered to the velocity instead
            #      of the K.  By writing the sidecar first, _src_K
            #      reads K from the sidecar and the refresh stays
            #      consistent.
            try:
                label_to_idx = {lbl: i for i, lbl in
                                enumerate(_CALIB_PARAMS)}
                k_idx = label_to_idx.get("Hydraulic Conductivity (k)")
                i_idx = label_to_idx.get("Hydraulic Gradient (i)")
                mid = _app_ref.v_calib_mid
                K_str = mid[k_idx].get().strip() if k_idx is not None else ""
                i_str = mid[i_idx].get().strip() if i_idx is not None else "1"
                if K_str:
                    K = float(K_str.replace(",", ""))
                    i_val = float(i_str.replace(",", "")) if i_str else 1.0
                    vd_my = K * i_val
                    vd_fy = vd_my / 0.3048
                    gw_path = os.path.join(BASE_DIR, "gwvelocity_inputs.txt")
                    with open(gw_path, "w", encoding="utf-8") as fp:
                        fp.write("Groundwater BulkDarcy Velocity Calculator Results\n")
                        fp.write(f"Bulk Hydraulic Conductivity Value: {K:g}\n")
                        fp.write(f"Bulk Hydraulic Conductivity Unit: m/year\n")
                        fp.write(f"Bulk Hydraulic Gradient: {i_val:g}\n")
                        fp.write(f"Bulk Darcy Velocity (m/year): {vd_my:.6f}\n")
                        fp.write(f"Bulk Darcy Velocity (ft/year): {vd_fy:.6f}\n")
            except Exception as exc:
                print(f"[load_optimal] gwvelocity_inputs.txt write failed: {exc}")

            # (2) Push Mid → §1-§7 inputs.  Now that the sidecar
            #     exists with the correct K, _push's trailing
            #     refresh will pull K from sidecar (not v_darcy
            #     fallback) and Mid stays correct.
            cb = getattr(_app_ref, "_push_calib_mids_to_inputs", None)
            if callable(cb):
                try: cb()
                except Exception: pass

            # NOTE: we deliberately DO NOT call
            # _apply_calib_multipliers_to_s7 here.  optimal_model.txt
            # v3 stores §7 source values directly (post-multiplier),
            # and _load_optimal_model already restored them above.
            # Re-applying the multiplier on top of those values
            # would double the reduction on every Load click.
            # The Mid column shows the optimal multiplier for the
            # user's reference; §7 already reflects that multiplier.
            #
            # (3) Recompute retardation from the loaded PFAA species
            #     + porosity so v_ret_trans* line up with the
            #     restored source-cell snapshot.
            try: _app_ref._on_pfaa_change()
            except Exception: pass
            messagebox.showinfo("Load Optimal Data",
                                f"Restored {n} calibration row(s) from:\n"
                                f"{path}\n\n"
                                "§3 v_darcy = K × i  (computed jointly).\n"
                                "§5 retardation, §6 alphax, etc. restored.\n"
                                "Click '5. Run Optimal Model' to run.")
        except Exception as exc:
            messagebox.showerror("Load Optimal Data",
                                 f"Could not parse {path}:\n{exc}")
        return

    if macro_name == "Save_Optimal_Model":
        # Save the calibration ranges + the §3 source values used to
        # produce them so Load Optimal Data can deterministically
        # restore the model state.  Format: one block per row, each
        # block tagged with the row label so re-ordering rows in
        # main.py won't desync past saves.
        #
        # v90: prefer best_calib.json (written by cali_1 right after
        # the optimizer finishes) over v_calib_mid for the calibration
        # rows.  v_calib_mid can drift after the post-calibration
        # callbacks (push to source cells, apply multipliers, trace
        # refreshes) — best_calib.json is the canonical record of the
        # iteration with the lowest RMSLE.
        if _app_ref is None:
            return
        best_overrides = {}
        best_rmsle_str = ""
        try:
            best_path = os.path.join(BASE_DIR, "best_calib.json")
            if os.path.exists(best_path):
                import json as _json
                with open(best_path, "r", encoding="utf-8") as fp:
                    bp = _json.load(fp)
                for lbl, val in zip(bp.get("labels", []),
                                    bp.get("best_x", [])):
                    best_overrides[lbl] = val
                if bp.get("best_rmsle") is not None:
                    best_rmsle_str = f"{bp['best_rmsle']:.6g}"
        except Exception as exc:
            print(f"[Save Optimal] best_calib.json read failed: {exc}")
        try:
            path = os.path.join(BASE_DIR, "optimal_model.txt")
            with open(path, "w", encoding="utf-8") as fp:
                fp.write("REMFluor-MD Optimal Model snapshot v4\n")
                if best_rmsle_str:
                    fp.write(f"Best RMSLE: {best_rmsle_str}\n")
                fp.write(f"Iterations: {_app_ref.v_n_iter.get()}\n")
                # Snapshot the §3 / §5 / §6 source cells —
                # Load Optimal Data uses these to restore inputs
                # exactly (and to set v_darcy for the K row).
                src_snapshot = {
                    "v_yr_start":   _app_ref.v_yr_start,
                    "v_darcy":      _app_ref.v_darcy,
                    "v_porf":       _app_ref.v_porf,
                    "v_alpha_l":    _app_ref.v_alpha_l,
                    "v_ret_trans1": _app_ref.v_ret_trans1,
                    "v_ret_trans2": _app_ref.v_ret_trans2,
                    "v_ret_trans3": _app_ref.v_ret_trans3,
                    "v_ret_trans4": _app_ref.v_ret_trans4,
                }
                fp.write("# Source cell snapshot (§3 / §5 / §6)\n")
                for k, v in src_snapshot.items():
                    try: fp.write(f"src.{k}={v.get()}\n")
                    except Exception: pass
                # v3 ADDITION: §7 source-zone concentration grid.
                # Loading these directly avoids the compounding bug
                # where re-applying the multiplier on top of an
                # already-multiplied §7 produced v × m² values.
                fp.write("\n# §7 source zone (post-multiplier values)\n")
                for col, lst in (("pfaa1", _app_ref.v_src_pfaa1),
                                  ("pfaa2", _app_ref.v_src_pfaa2),
                                  ("pre1",  _app_ref.v_src_pre1),
                                  ("pre2",  _app_ref.v_src_pre2)):
                    for i, var in enumerate(lst):
                        try:
                            fp.write(f"s7.{col}[{i}]={var.get()}\n")
                        except Exception:
                            pass
                fp.write("\n# Step 4 calibration rows\n")
                for i, label in enumerate(_CALIB_PARAMS):
                    if i >= min(len(_app_ref.v_calib_chk),
                                 len(_app_ref.v_calib_low),
                                 len(_app_ref.v_calib_mid),
                                 len(_app_ref.v_calib_high)):
                        break
                    # v90: prefer best_calib.json's value over v_calib_mid
                    # for this row when the optimizer wrote one.
                    if label in best_overrides:
                        mid_val = f"{best_overrides[label]:g}"
                    else:
                        mid_val = _app_ref.v_calib_mid[i].get()
                    fp.write(f"row[{i}]\n")
                    fp.write(f"  label={label}\n")
                    fp.write(f"  use={_app_ref.v_calib_chk[i].get()}\n")
                    fp.write(f"  lo={_app_ref.v_calib_low[i].get()}\n")
                    fp.write(f"  mid={mid_val}\n")
                    fp.write(f"  hi={_app_ref.v_calib_high[i].get()}\n")
            messagebox.showinfo("Save Optimal Model",
                                f"Saved to:\n{path}")
        except Exception as exc:
            messagebox.showerror("Save Optimal Model",
                                 f"Save failed:\n{exc}")
        return

    if macro_name == "Help_Calibration":
        messagebox.showinfo(
            "Calibration Panel — Help",
            "REMFluor-MD MACHINE-CALIBRATION (Singh et al., 2025)\n\n"
            "Step 1) Enter monitoring well data in Section 9.\n"
            "Step 2) Tick which parameters to calibrate against.\n"
            "Step 3) Adjust per-well weighting factors (downgradient\n"
            "        wells are typically weighted higher).\n"
            "Step 4) Tick the parameters to vary, set Lowest /\n"
            "        Highest Likely Values.  Mid-Range pulls from\n"
            "        the input sheet automatically.\n"
            "Step 5) Click '2. Run Machine Based Calibration'.\n\n"
            "After the run:\n"
            "  '3. See All Calibration Data' — opens run_history.csv\n"
            "  '4. Load Optimal Data'         — pulls best-fit values\n"
            "                                  back into the inputs\n"
            "  '5. Run Optimal Model'         — re-runs the solver\n"
            "                                  with the optimal set\n"
            "  '6. Save Optimal Model'        — exports the snapshot")
        return

    if macro_name == "Authors":
        messagebox.showinfo(
            "REMFluor-MD Authors",
            "REMFluor-MD v2.6\n\n"
            "Singh et al. (2025)\n"
            "Falta et al. (2025)\n\n"
            "ESTCP – Environmental Security Technology Certification Program"
        )
        return

    # ── Pure-Python ports of the dist/ EXEs (no Excel install needed) ──
    if _FUNCS_LOADED and _app_ref is not None:
        if macro_name == "SourceRemediation":
            # v101: prefer the baseline-aware _apply_s8_to_s7 (defined
            # in App.__init__) so repeated clicks don't compound — it
            # always re-derives §7 from the captured baseline using the
            # CURRENT §8 % and start year.  Fall back to the popup port
            # (popups_source_remediation.run multiplies in place — fine
            # for first click, compounds on repeat clicks).
            try:
                cb = getattr(_app_ref, "_apply_s8_to_s7", None)
                if callable(cb):
                    cb()
                    # Light confirmation so the user sees something
                    # happened.  Read §8 values for the message.
                    try:
                        red = _app_ref.v_src_conc_red.get().strip()
                        yr  = _app_ref.v_src_rem_yr.get().strip()
                        messagebox.showinfo(
                            "Apply Remediation",
                            f"Source Concentration Reduction = {red}%\n"
                            f"Source Treatment Start Year = {yr}\n\n"
                            f"§7 PFAA-1 / PFAA-2 (and precursors in "
                            f"Detailed) updated from baseline.")
                    except Exception:
                        pass
                else:
                    popups_source_remediation.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("Apply Remediation",
                                     f"Could not apply remediation:\n{exc}")
            return
        if macro_name == "ModelingTransformationLowK":
            try:
                popups_transformation.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("Modeling Transformation Low-K",
                                     f"Popup failed:\n{exc}")
            return
        if macro_name == "ChangeNumericalParameters":
            try:
                popups_numerical.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("Change Numerical Parameters",
                                     f"Popup failed:\n{exc}")
            return
        if macro_name == "OpenAppendix_2_1_Relative_EXE":
            try:
                popups_cellsize.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("Cell Size",
                                     f"Popup failed:\n{exc}")
            return
        if macro_name == "GWVelocityCalculator":
            try:
                popups_GWvelocity.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("GW Velocity Calculator",
                                     f"Popup failed:\n{exc}")
            # Refresh §calibration Step 4 K/i Mid cells now that the
            # GW Velocity Calculator may have just written new
            # gwvelocity_inputs.txt — _src_K / _src_i re-read the
            # file so a refresh picks up fresh values immediately.
            cb = getattr(_app_ref, "_refresh_calib_mids", None)
            if callable(cb):
                try: cb()
                except Exception: pass
            return
        if macro_name == "LongevityTool":
            try:
                popups_longevity.run(_app_ref)
            except Exception as exc:
                import traceback
                tb = traceback.format_exc()
                print("[LongevityTool] popup failed:\n" + tb)
                messagebox.showerror("Longevity Tool",
                                     f"Popup failed:\n{exc}\n\n"
                                     f"Traceback (last lines):\n"
                                     f"{tb.strip().splitlines()[-1] if tb else ''}")
            return
        if macro_name == "SourceOption2":
            try:
                popups_mass_discharge_import.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("Source Concentrations",
                                     f"Popup failed:\n{exc}")
            return
        if macro_name == "CalibrationDataLoader":
            try:
                popups_calibration.run(_app_ref)
            except Exception as exc:
                messagebox.showerror("Calibration Data Loader",
                                     f"Popup failed:\n{exc}")
            return
        if macro_name == "HeterogeneityCalculator_Unconsolidated_Media":
            try:
                popups_heterogeneity.run(_app_ref,
                                        media_type="Unconsolidated Media")
            except Exception as exc:
                messagebox.showerror("Heterogeneity Calculator",
                                     f"Popup failed:\n{exc}")
            # Refresh §calibration Step 4 — Heterogeneity Calculator
            # writes heterogeneity_inputs.txt which feeds volfrac /
            # difflen Mid cells.
            cb = getattr(_app_ref, "_refresh_calib_mids", None)
            if callable(cb):
                try: cb()
                except Exception: pass
            return
        if macro_name == "HeterogeneityCalculator_Fractured_Rock":
            try:
                popups_heterogeneity.run(_app_ref,
                                        media_type="Fractured Rock")
            except Exception as exc:
                messagebox.showerror("Heterogeneity Calculator",
                                     f"Popup failed:\n{exc}")
            cb = getattr(_app_ref, "_refresh_calib_mids", None)
            if callable(cb):
                try: cb()
                except Exception: pass
            return

    messagebox.showinfo(
        "Not Mapped",
        f"No action mapped for macro '{macro_name}'.\n"
        "Add a handler in run_script() or register it in HTML_CHICKLETS / HTML_APPENDIX."
    )


def open_quarto(macro_name):
    run_script(macro_name)

# ─────────────────────────────────────────────────────────────────────────────
# HELPER WIDGETS
# ─────────────────────────────────────────────────────────────────────────────
class _RoundButton(tk.Canvas):
    """
    Canvas-based rounded-corner button — Excel storyboard buttons all
    use rounded rectangles with a 1-px gray outline.  tk.Button cannot
    draw rounded corners; this class does, while exposing the same
    `command` / event semantics.
    """
    def __init__(self, parent, text, command, *, fg=FG_BTN_NAVY,
                 font=None, bg=BTN_FILL, padx=10, pady=4, width=None,
                 radius=8, anchor="center"):
        import tkinter.font as _tkFont
        f = _tkFont.nametofont(font) if font else _tkFont.nametofont("TkDefaultFont")
        lines = text.split("\n")
        text_w = max((f.measure(l) for l in lines), default=20)
        text_h = f.metrics("linespace") * len(lines) + 2 * (len(lines) - 1)
        if width:
            text_w = max(text_w, int(f.measure("0") * width))
        w = text_w + padx * 2
        h = text_h + pady * 2
        super().__init__(parent, width=w, height=h, bg=BG_MAIN,
                         highlightthickness=0, bd=0, cursor="hand2")
        # NOTE: avoid overwriting tk.Misc internal attributes (_w / _h
        # are used by Tk for the widget path).  Use _cw / _ch instead.
        self._cmd  = command
        self._fg   = fg
        self._bg   = bg
        self._cw   = w
        self._ch   = h
        self._r    = radius
        self._bg_id  = self._draw_rrect(self._bg)
        self._txt_id = self.create_text(w/2, h/2, text=text, fill=fg,
                                        font=font, justify="center",
                                        anchor=anchor)
        # Click + hover bindings
        for ev, h_fn in [("<Button-1>",        self._on_press),
                         ("<ButtonRelease-1>", self._on_release),
                         ("<Enter>",           self._on_enter),
                         ("<Leave>",           self._on_leave)]:
            self.bind(ev, h_fn)

    # rounded rectangle as smoothed polygon
    def _draw_rrect(self, fill):
        x1, y1, x2, y2, r = 1, 1, self._cw-1, self._ch-1, self._r
        pts = [
            x1+r, y1, x2-r, y1, x2, y1, x2, y1+r,
            x2,   y2-r, x2,   y2, x2-r, y2, x1+r, y2,
            x1,   y2, x1,   y2-r, x1,   y1+r, x1, y1,
        ]
        return self.create_polygon(pts, smooth=True, fill=fill,
                                   outline="#777")

    def _on_enter(self, _):
        self.itemconfig(self._bg_id, fill="#C0C0C0")
    def _on_leave(self, _):
        self.itemconfig(self._bg_id, fill=self._bg)
    def _on_press(self, _):
        self.itemconfig(self._bg_id, fill="#A8A8A8")
    def _on_release(self, _):
        self.itemconfig(self._bg_id, fill=self._bg)
        if self._cmd:
            self._cmd()


def make_btn(parent, text, macro, quarto=False, fg=FG_BTN_NAVY,
             font=FONT_BTN, width=None, padx=4, pady=2, bg=BTN_FILL,
             anchor="center"):
    cmd = (lambda m=macro: open_quarto(m)) if quarto else (lambda m=macro: run_script(m))
    return _RoundButton(parent, text, cmd, fg=fg, font=font, bg=bg,
                        padx=max(padx, 8), pady=max(pady, 3),
                        width=width, radius=8, anchor=anchor)


def make_entry(parent, var, width=10, bg=BG_INPUT_BLUE, justify="right"):
    # Auto-pick readable text color: white on dark backgrounds, black on light.
    is_locked = bg.lower() in ("#000000", "#000", "black")
    fg = "#FFFFFF" if is_locked else FG_INPUT
    e = tk.Entry(parent, textvariable=var, width=width, font=FONT_INPUT,
                 bg=bg, fg=fg, relief="solid", bd=1, justify=justify,
                 insertbackground=fg,
                 disabledbackground=bg, disabledforeground=fg,
                 readonlybackground=bg)
    if is_locked:
        # Black cells = calculated values (per Excel storyboard).
        # Make them read-only so users can't accidentally overwrite.
        e.configure(state="readonly")
    return e


def make_label(parent, text, font=FONT_LABEL, fg=FG_INPUT, bg=BG_MAIN,
               anchor="w", wraplength=0):
    kw = dict(text=text, font=font, fg=fg, bg=bg, anchor=anchor)
    if wraplength:
        kw["wraplength"] = wraplength
    return tk.Label(parent, **kw)


def help_link(parent, macro, bg=BG_MAIN):
    """
    Red '?' help link inside a TRUE square box, sized to match the
    height of an adjacent tk.Entry so the box appears flush-adjacent
    to its input cell (the tkinter equivalent of the PySide6
    HelpLink button).  Win9x raised look: white top/left, gray
    bottom/right.

    Tk Labels are sized in characters x lines, which gives a tall
    rectangle for a single character.  To force a square we wrap the
    '?' Label inside a Frame with explicit pixel width=height and
    pack_propagate / grid_propagate disabled so the children can't
    push the Frame out of square.
    """
    # Compute the square side from FONT_HELP linespace so the box
    # scales with the chosen font size (and zoom). Use a slightly
    # smaller multiplier so the box hugs the entry instead of
    # towering above it.
    try:
        import tkinter.font as _tkFont
        ls = _tkFont.nametofont(FONT_HELP).metrics("linespace")
        side = max(13, int(ls * 1.05))
    except Exception:
        side = 14

    # "raised" relief on tk.Frame draws a Win9x-style bevel — white
    # top/left, dark bottom/right — which mirrors the PySide6 HelpLink
    # styling.
    box = tk.Frame(parent, bg=BTN_FILL, width=side, height=side,
                   bd=1, relief="raised", highlightthickness=0)
    box.pack_propagate(False)
    box.grid_propagate(False)

    # Prefer the Excel "?" icon (Figures/QuestionMark.png) when present.
    # Fall back to a red text glyph if Pillow / image not available.
    img = _help_qmark_image(side - 4)
    if img is not None:
        lbl = tk.Label(box, image=img, bg=BTN_FILL,
                       cursor="hand2",
                       bd=0, padx=0, pady=0,
                       takefocus=0, highlightthickness=0)
        lbl.image = img        # keep ref so Tk doesn't GC
    else:
        lbl = tk.Label(
            box, text="?",
            font=FONT_HELP,
            fg=FG_HELP,
            bg=BTN_FILL,
            cursor="hand2",
            bd=0, padx=0, pady=0,
            relief="flat",
            takefocus=0,
            highlightthickness=0,
        )
    lbl.place(relx=0.5, rely=0.5, anchor="center")

    handler = lambda e, m=macro: run_script(m)
    box.bind("<Button-1>", handler)
    lbl.bind("<Button-1>", handler)
    return box


# Cache the QuestionMark icon at multiple sizes (Tk PhotoImage refs
# must stay alive — store on a module-level dict keyed by pixel side).
_QMARK_CACHE: dict = {}

def _help_qmark_image(side: int):
    """Return Tk PhotoImage of Figures/QuestionMark.png scaled to
    `side x side`, or None if the file/Pillow is unavailable."""
    side = max(8, int(side))
    if side in _QMARK_CACHE:
        return _QMARK_CACHE[side]
    path = os.path.join(FIGURES_DIR, "QuestionMark.png")
    if not os.path.isfile(path):
        return None
    try:
        from PIL import Image, ImageTk
        img = Image.open(path).convert("RGBA")
        try:
            resample = Image.Resampling.LANCZOS
        except AttributeError:
            resample = Image.LANCZOS
        img = img.resize((side, side), resample)
        photo = ImageTk.PhotoImage(img)
    except Exception:
        try:
            photo = tk.PhotoImage(file=path)
            f = max(1, photo.height() // side)
            if f > 1:
                photo = photo.subsample(f, f)
        except Exception:
            return None
    _QMARK_CACHE[side] = photo
    return photo


def big_radio(parent, text, variable, value, *,
              text_font=None, bullet_font=None, bg=BG_MAIN):
    """Radio with a Canvas-drawn circle so the size is OS-independent.

    Windows locks tk.Radiobutton's indicator to the OS-theme size regardless
    of font, so we draw our own circle on a Canvas and wire up click/trace
    events to replicate full radio behaviour.

    RADIO_R controls the circle radius in pixels (default 9 → 18 px diameter).
    """
    RADIO_R = 9  # radius in pixels — change here to resize all circles
    if text_font is None:
        text_font = FONT_LABEL

    fr  = tk.Frame(parent, bg=bg)
    dia = RADIO_R * 2 + 4          # canvas size with a small margin
    cv  = tk.Canvas(fr, width=dia, height=dia, bg=bg,
                    highlightthickness=0, bd=0)
    cv.pack(side="left")

    lbl = tk.Label(fr, text=text, font=text_font, bg=bg, anchor="w")
    lbl.pack(side="left", padx=(3, 0))

    def _draw(*_):
        cv.delete("all")
        # Outer ring
        cv.create_oval(2, 2, RADIO_R * 2 + 2, RADIO_R * 2 + 2,
                       outline="#555555", width=1.5)
        # Inner filled dot when selected
        if variable.get() == value:
            inner = max(RADIO_R // 2, 3)
            cx = RADIO_R + 2
            cv.create_oval(cx - inner, cx - inner, cx + inner, cx + inner,
                           fill="#333333", outline="#333333")

    def _select(_event=None):
        variable.set(value)

    cv.bind("<Button-1>", _select)
    lbl.bind("<Button-1>", _select)
    variable.trace_add("write", _draw)
    _draw()  # initial render
    return fr


def big_check(parent, text, variable, *, text_font=None, bg=BG_MAIN):
    """Checkbox with a Canvas-drawn circle indicator — matches big_radio
    (used in §1 / §6) so the look is consistent across the whole app.

    Like big_radio, the size is OS-independent (Windows locks the
    native tk.Checkbutton indicator regardless of font).  Set RADIO_R
    in big_radio() to scale both helpers together.

    The indicator toggles `variable` (a BooleanVar) on click.  Pass
    text="" to omit the label.
    """
    RADIO_R = 9
    if text_font is None:
        text_font = FONT_LABEL

    fr  = tk.Frame(parent, bg=bg)
    dia = RADIO_R * 2 + 4
    cv  = tk.Canvas(fr, width=dia, height=dia, bg=bg,
                    highlightthickness=0, bd=0)
    cv.pack(side="left")

    if text:
        lbl = tk.Label(fr, text=text, font=text_font, bg=bg,
                       anchor="w", cursor="hand2")
        lbl.pack(side="left", padx=(3, 0))
    else:
        lbl = None

    def _draw(*_):
        cv.delete("all")
        cv.create_oval(2, 2, RADIO_R * 2 + 2, RADIO_R * 2 + 2,
                       outline="#555555", width=1.5)
        try:
            on = bool(variable.get())
        except Exception:
            on = False
        if on:
            inner = max(RADIO_R // 2, 3)
            cx = RADIO_R + 2
            cv.create_oval(cx - inner, cx - inner,
                           cx + inner, cx + inner,
                           fill="#333333", outline="#333333")

    def _toggle(_event=None):
        try:
            variable.set(not bool(variable.get()))
        except Exception:
            pass

    cv.config(cursor="hand2")
    cv.bind("<Button-1>", _toggle)
    if lbl:
        lbl.bind("<Button-1>", _toggle)
    variable.trace_add("write", _draw)
    _draw()
    return fr


def section_header(parent, num, text, bg=BG_MAIN):
    return tk.Label(parent, text=f"{num}.  {text}",
                    font=FONT_SECTION, fg=FG_SECTION, bg=bg, anchor="w")


def dropdown(parent, var, choices, width=10, bg=None):
    cb = ttk.Combobox(parent, textvariable=var, values=choices, width=width,
                      font=FONT_INPUT, state="readonly")
    cb.set(var.get())
    if bg:
        # ttk widgets ignore plain `bg=`; build a per-color style instead.
        # A readonly Combobox displays its current value using the
        # SELECTION colors, so we have to override those too — otherwise
        # the field stays the OS default.
        style_name = f"C{bg.replace('#','').upper()}.TCombobox"
        s = ttk.Style()
        s.configure(style_name,
                    fieldbackground=bg, background=bg,
                    foreground=FG_INPUT, arrowcolor=FG_INPUT,
                    selectbackground=bg, selectforeground=FG_INPUT)
        s.map(style_name,
              fieldbackground=[("readonly", bg), ("active", bg),
                               ("focus", bg), ("!focus", bg)],
              background=[("readonly", bg), ("active", bg)],
              foreground=[("readonly", FG_INPUT)],
              selectbackground=[("readonly", bg), ("focus", bg),
                                ("!focus", bg)],
              selectforeground=[("readonly", FG_INPUT),
                                ("focus", FG_INPUT),
                                ("!focus", FG_INPUT)])
        cb.configure(style=style_name)
    return cb


# ─────────────────────────────────────────────────────────────────────────────
# LOGO LOADER  (looks in BASE_DIR/Figures/ for the ESTCP image)
# ─────────────────────────────────────────────────────────────────────────────
def _load_figure(filename: str, target_height: int = None, target_width: int = None):
    """
    Load BASE_DIR/Figures/<filename> and return a Tk PhotoImage (or None
    if the file isn't found / can't be decoded).

    If only one of target_height/target_width is given, the other is
    computed to preserve aspect ratio.  Prefers Pillow for resizing
    quality; falls back to native tk.PhotoImage if Pillow is missing.
    """
    if not filename:
        return None
    path = os.path.join(FIGURES_DIR, filename)
    if not os.path.isfile(path):
        return None
    # Pillow path (preferred)
    try:
        from PIL import Image, ImageTk
        img = Image.open(path)
        if img.mode not in ("RGB", "RGBA"):
            img = img.convert("RGBA")
        w, h = img.size
        if target_height and not target_width:
            target_width = max(1, int(w * target_height / h))
        elif target_width and not target_height:
            target_height = max(1, int(h * target_width / w))
        if target_height and target_width:
            try:
                resample = Image.Resampling.LANCZOS
            except AttributeError:
                resample = Image.LANCZOS
            img = img.resize((target_width, target_height), resample)
        return ImageTk.PhotoImage(img)
    except Exception:
        pass
    # Fallback: native Tk PhotoImage (PNG / GIF only, no smooth resize)
    try:
        return tk.PhotoImage(file=path)
    except Exception:
        return None


def _load_logo_image(target_height: int = 50):
    """
    Returns a Tk-compatible image object (PhotoImage / ImageTk.PhotoImage)
    or None if no suitable image can be found.

    Search order:
      1) Files in Figures/ whose name contains "estcp" (PNG/GIF/JPG/BMP)
      2) Files in Figures/ whose name contains "logo"
      3) Any image file in Figures/
    Prefers PIL (Pillow) for resampling/JPG support; falls back to native
    Tk PhotoImage with subsample() if PIL is unavailable.
    """
    if not os.path.isdir(FIGURES_DIR):
        return None

    exts = (".png", ".gif", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff")
    all_files = [f for f in os.listdir(FIGURES_DIR) if f.lower().endswith(exts)]

    def score(name):
        low = name.lower()
        s = 0
        if "estcp" in low: s += 100
        if "logo"  in low: s += 50
        if low.endswith(".png"): s += 5  # prefer PNG
        return s

    candidates = sorted(all_files, key=score, reverse=True)

    for fname in candidates:
        path = os.path.join(FIGURES_DIR, fname)
        # Try Pillow first (handles all formats + smooth resize)
        try:
            from PIL import Image, ImageTk
            pil_img = Image.open(path)
            if pil_img.mode not in ("RGB", "RGBA"):
                pil_img = pil_img.convert("RGBA")
            w, h = pil_img.size
            if h <= 0:
                continue
            new_w = max(1, int(w * target_height / h))
            try:
                resample = Image.Resampling.LANCZOS
            except AttributeError:
                resample = Image.LANCZOS
            pil_img = pil_img.resize((new_w, target_height), resample)
            return ImageTk.PhotoImage(pil_img)
        except Exception:
            pass
        # Fallback: native Tk PhotoImage (PNG / GIF only)
        try:
            img = tk.PhotoImage(file=path)
            h = img.height()
            if h > target_height:
                factor = max(1, h // target_height)
                img = img.subsample(factor, factor)
            return img
        except Exception:
            continue

    return None


# ─────────────────────────────────────────────────────────────────────────────
# MAIN APPLICATION
# ─────────────────────────────────────────────────────────────────────────────
class REMFluorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("REMFluor-MD Model Input Screen  v2.6")
        self.configure(bg=BG_MAIN)
        self.resizable(True, True)

        # Match Tk's internal scaling to the actual display DPI so fonts/widgets
        # render at their intended physical size. Without this, the high-DPI
        # awareness call shrinks every widget to ~50-66% of its expected size.
        self._apply_tk_dpi_scaling()

        # ttk's default Windows theme ('vista') draws Combobox fields with the
        # native OS colors and ignores `fieldbackground` style settings.
        # Switching to 'clam' makes ttk honor our color customizations so the
        # PFAA / Clay dropdowns can actually appear in Pull-Down peach.
        try:
            s = ttk.Style()
            if "clam" in s.theme_names():
                s.theme_use("clam")
        except Exception:
            pass

        # Create the Tk named fonts (lets us zoom in/out at runtime).
        self._init_named_fonts()

        # Bind zoom shortcuts: Ctrl++ / Ctrl+- / Ctrl+0 / Ctrl+wheel
        self.bind_all("<Control-plus>",        self._zoom_in)
        self.bind_all("<Control-equal>",       self._zoom_in)   # plus w/o shift
        self.bind_all("<Control-KP_Add>",      self._zoom_in)
        self.bind_all("<Control-minus>",       self._zoom_out)
        self.bind_all("<Control-KP_Subtract>", self._zoom_out)
        self.bind_all("<Control-0>",           self._zoom_reset)
        self.bind_all("<Control-MouseWheel>",  self._zoom_wheel)

        self.active_sheet = "Simple"
        self._detailed_only_frames = []  # frames hidden in Simple mode
        self._simple_only_frames   = []  # frames hidden in Detailed mode
        global _app_ref
        _app_ref = self

        # v105: clean-start guard.  Building the dropdowns calls
        # cb.set(var.get()) which re-writes each StringVar and therefore
        # FIRES its write-trace — that's what made the dropdown change
        # handlers (media → porosity, PFAA → retardation, heterogeneity →
        # dispersivity) auto-fill the calculated cells at startup even
        # though their defaults are blank.  While this flag is True those
        # handlers skip their auto-fill, so the form opens genuinely empty.
        # The flag is cleared once the UI is fully built.
        self._building = True

        self._build_vars()

        outer = tk.Frame(self, bg=BG_MAIN)
        outer.pack(fill="both", expand=True)

        # yscrollincrement / xscrollincrement = 1 means yview_scroll/xview_scroll
        # operate in single-pixel units, giving us pixel-precise control over
        # how far each wheel tick scrolls (see _scroll_y_pixels / _scroll_x_pixels).
        self.canvas = tk.Canvas(outer, bg=BG_MAIN, bd=0, highlightthickness=0,
                                yscrollincrement=1, xscrollincrement=1)
        # Classic Win9x-style scrollbars via ttk + clam theme:
        # light blue thumb on a gray track with dark arrow buttons and a
        # dotted grip in the middle of the thumb — matches user's
        # reference image.
        try:
            _ssty = ttk.Style()
            if "clam" in _ssty.theme_names() and _ssty.theme_use() != "clam":
                _ssty.theme_use("clam")
            for orient in ("Vertical", "Horizontal"):
                _ssty.configure(f"Big.{orient}.TScrollbar",
                                background="#A3C1E0",       # blue thumb
                                troughcolor="#E4E8EE",      # light gray track
                                bordercolor="#7B7B7B",
                                lightcolor="#FFFFFF",
                                darkcolor="#7B7B7B",
                                arrowcolor="#000000",
                                arrowsize=22,
                                gripcount=4)
                _ssty.map(f"Big.{orient}.TScrollbar",
                          background=[("active",  "#7AA8D0"),
                                      ("pressed", "#5C8AA0")])
        except Exception:
            pass

        vsb = ttk.Scrollbar(outer, orient="vertical",
                            style="Big.Vertical.TScrollbar",
                            command=self.canvas.yview)
        hsb = ttk.Scrollbar(outer, orient="horizontal",
                            style="Big.Horizontal.TScrollbar",
                            command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.canvas.pack(side="left", fill="both", expand=True)

        self.inner = tk.Frame(self.canvas, bg=BG_MAIN)
        # Use natural height (inner expands to fit all content); width is
        # synced to the visible canvas in _on_canvas_configure so the
        # vertical scrollbar shows up when content overflows the window.
        self.canvas_win = self.canvas.create_window(
            (0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)
        # Vertical wheel
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        # Horizontal wheel — Shift+wheel on Win/macOS, dedicated on Linux
        self.canvas.bind_all("<Shift-MouseWheel>", self._on_mousewheel_x)
        # Linux X11 fallback (delta isn't reported; use Button-4 / 5)
        self.canvas.bind_all("<Button-4>",
                             lambda e: self._wheel_dir(e, dy=-1))
        self.canvas.bind_all("<Button-5>",
                             lambda e: self._wheel_dir(e, dy=+1))
        self.canvas.bind_all("<Shift-Button-4>",
                             lambda e: self._wheel_dir(e, dx=-1))
        self.canvas.bind_all("<Shift-Button-5>",
                             lambda e: self._wheel_dir(e, dx=+1))

        self._build_ui()
        # Scale the initial window with DPI so it isn't a tiny postage stamp
        # on a 4K display. _dpi_scale is 1.0 at 96 DPI, 1.5 at 144 DPI, etc.
        scale = getattr(self, "_dpi_scale", 1.0)
        # Open at a comfortable size; the surrounding canvas exposes
        # vertical and horizontal scrollbars when content overflows.
        scr_w = self.winfo_screenwidth()
        scr_h = self.winfo_screenheight()
        win_w = min(int(1500 * scale), scr_w - 60)
        win_h = min(int(950  * scale), scr_h - 100)
        self.geometry(f"{win_w}x{win_h}")

        # v105: UI fully built — allow the dropdown change handlers to do
        # their normal auto-fill from here on (genuine user selections).
        self._building = False

    # ── Named fonts + zoom ───────────────────────────────────────────────
    def _init_named_fonts(self):
        """Create the Tk named fonts described by _FONT_DEFS."""
        import tkinter.font as tkFont
        self._tk_fonts = {}
        self._base_font_sizes = {}
        self._zoom = 1.0
        for name, (family, size, weight, slant) in _FONT_DEFS.items():
            try:
                fnt = tkFont.nametofont(name)
                fnt.config(family=family, size=size, weight=weight, slant=slant)
            except tk.TclError:
                fnt = tkFont.Font(name=name, family=family, size=size,
                                  weight=weight, slant=slant)
            self._tk_fonts[name] = fnt
            self._base_font_sizes[name] = size

    def _apply_zoom(self, factor):
        """Resize every named font to (base_size * factor)."""
        self._zoom = max(0.5, min(3.5, factor))
        for name, base in self._base_font_sizes.items():
            try:
                self._tk_fonts[name].config(
                    size=max(4, int(round(base * self._zoom))))
            except Exception:
                pass

    def _zoom_in(self,    _=None): self._apply_zoom(self._zoom * 1.10)
    def _zoom_out(self,   _=None): self._apply_zoom(self._zoom / 1.10)
    def _zoom_reset(self, _=None): self._apply_zoom(1.0)

    def _zoom_wheel(self, event):
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()
        return "break"

    # ── Section divider helper ───────────────────────────────────────────
    def _hsep(self, parent, color="#000000", thickness=1, pady=8):
        """Thin horizontal black line, used between sections in a column.
        Default pady bumped 3 → 8 so each section in the right column
        has more breathing room above/below its divider line."""
        sep = tk.Frame(parent, bg=color, height=thickness)
        sep.pack(fill="x", padx=2, pady=pady)
        return sep

    # ── DPI / scaling ─────────────────────────────────────────────────────
    def _apply_tk_dpi_scaling(self):
        """
        Scale Tk's font/widget sizes to match the display DPI.

        Tk's default 'scaling' assumes 72 DPI; on a real Windows monitor
        that's typically 96 DPI (= scaling 1.333). On a 125% scaled display
        the OS reports 120 DPI (scaling 1.667); 150% reports 144 DPI
        (scaling 2.0); 175% reports 168 DPI (scaling 2.333). We compute the
        ratio from winfo_fpixels('1i') and apply it.
        """
        try:
            dpi = self.winfo_fpixels("1i")  # pixels per inch on the display
            if dpi and dpi > 0:
                self.tk.call("tk", "scaling", dpi / 72.0)
                # Cache the scale factor so geometry calls can use it too
                self._dpi_scale = max(1.0, dpi / 96.0)
            else:
                self._dpi_scale = 1.0
        except Exception:
            self._dpi_scale = 1.0

    # ── Layout helpers ────────────────────────────────────────────────────
    def _on_frame_configure(self, _):
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        # When canvas is wider than natural content width, stretch inner
        # to fill (no h-scrollbar needed). When canvas is narrower, keep
        # inner at its natural width so the horizontal scrollbar appears.
        try:
            inner_w = max(self.inner.winfo_reqwidth(), 1300)
        except Exception:
            inner_w = 1300
        self.canvas.itemconfig(self.canvas_win,
                               width=max(event.width, inner_w))

    # ── Mouse-wheel scrolling ────────────────────────────────────────
    # Tk's default "1 unit per 120-delta-tick" feels chunky on modern
    # high-resolution wheels.  We drop to ~3 px / tick and use the
    # canvas's pixel-based yview_scroll(..., "units") with a small
    # _SCROLL_PX scratchpad so consecutive ticks coalesce smoothly.
    _SCROLL_PX_PER_TICK = 24    # vertical pixels per 120-delta tick
    _SCROLL_PX_PER_TICK_X = 36  # horizontal pixels per 120-delta tick

    def _on_mousewheel(self, event):
        # event.delta is a multiple of 120 on Windows/macOS.
        ticks = -event.delta / 120.0
        self._scroll_y_pixels(int(ticks * self._SCROLL_PX_PER_TICK))
        return "break"

    def _on_mousewheel_x(self, event):
        ticks = -event.delta / 120.0
        self._scroll_x_pixels(int(ticks * self._SCROLL_PX_PER_TICK_X))
        return "break"

    def _wheel_dir(self, event, dx=0, dy=0):
        """Linux X11 wheel fallback (Button-4 = up, Button-5 = down)."""
        if dy:
            self._scroll_y_pixels(dy * self._SCROLL_PX_PER_TICK)
        if dx:
            self._scroll_x_pixels(dx * self._SCROLL_PX_PER_TICK_X)
        return "break"

    def _scroll_y_pixels(self, dy):
        # The canvas was created with yscrollincrement=1 so 1 "unit" =
        # 1 pixel.  Pass the pixel delta straight through — gives the
        # smoothest possible step on every wheel tick.
        if dy == 0:
            return
        try:
            self.canvas.yview_scroll(int(dy), "units")
        except Exception:
            pass

    def _scroll_x_pixels(self, dx):
        if dx == 0:
            return
        try:
            self.canvas.xview_scroll(int(dx), "units")
        except Exception:
            pass

    def show_calibration_panel(self):
        run_script("RunPythonScript")  # Auto-cal: just run model (Simple version)

    def show_main_panel(self):
        # "Go Back to Main Interface" — reset the scroll position to
        # top-left so the user lands where the app first opened
        # (top of §1, all the way scrolled left).  yview_moveto alone
        # was vertical-only; xview_moveto resets horizontal too.
        try:
            self.canvas.yview_moveto(0.0)
            self.canvas.xview_moveto(0.0)
        except Exception:
            pass
        try:
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def scroll_to_calibration(self):
        """Scroll the canvas so the §calibration quadrant fills the
        viewable area.  Wired to §11 'Run Model with Auto-Calibration'
        — user clicks expecting to fill in Step 2-4 first, not to kick
        off the optimizer.

        The panel lives in the bottom-right of a 2-column body, so
        BOTH x and y need to scroll: x to the panel's left edge
        (otherwise we show the empty left column), y to the panel's
        top edge.  Coordinates are walked up the widget hierarchy
        from self._calib_panel_frame to self.inner so the math
        survives any layout reshuffle.  Tk's yview_moveto / xview_moveto
        take a fraction of the total scrollable extent, and they
        auto-clamp so the panel doesn't scroll past the bottom-right
        corner — i.e. on a wide-enough window the panel can land
        exactly at the top-left of the visible canvas."""
        try:
            self.update_idletasks()
            panel = getattr(self, "_calib_panel_frame", None)
            if panel is not None and panel.winfo_exists():
                # Sum winfo_x / winfo_y up the parent chain until we
                # hit self.inner (the scrolled frame).
                x = y = 0
                w = panel
                while w is not None and w is not self.inner:
                    try:
                        x += w.winfo_x()
                        y += w.winfo_y()
                    except Exception:
                        break
                    w = w.master
                inner_w = max(self.inner.winfo_width(),  1)
                inner_h = max(self.inner.winfo_height(), 1)
                # Nudge upward by 4 px so the panel border isn't
                # sliced off by the scrollbar bevel.
                fy = max(0.0, min(1.0, max(0, y - 4) / inner_h))
                fx = max(0.0, min(1.0, max(0, x - 4) / inner_w))
                self.canvas.yview_moveto(fy)
                self.canvas.xview_moveto(fx)
            else:
                # Fallback if the panel ref isn't ready yet.
                self.canvas.yview_moveto(0.55)
                self.canvas.xview_moveto(0.5)
        except Exception:
            pass
        try:
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def _on_model_version_change(self, *_):
        val = self.v_model_version.get()
        self.active_sheet = "Simple" if val == "Simple Version" else "Detailed_2"
        is_detailed = (self.active_sheet == "Detailed_2")
        self._toggle_widgets(
            getattr(self, "_detailed_only_frames", []), show=is_detailed)
        self._toggle_widgets(
            getattr(self, "_simple_only_frames", []), show=not is_detailed)
        # Recompute retardation for new version context
        self._on_pfaa_change()
        # The Detailed-only toggle above will have re-shown the MW row's
        # extra Detailed cells (mw_e3/mw_e4) — but those must stay
        # hidden when the current Kf unit isn't mol-based.  Re-run the
        # §9 visibility logic so the final state honours BOTH rules.
        cb = getattr(self, "_on_psb_kf_unit_change", None)
        if callable(cb):
            cb()
        # §7 cells whose color differs between Simple and Detailed —
        # recolor now that active_sheet is up-to-date.
        if hasattr(self, "_apply_s7_version_colors"):
            try:
                self._apply_s7_version_colors()
            except Exception:
                pass

    def _toggle_widgets(self, widgets, show):
        """Show or hide widgets based on the manager kind tagged on
        the widget at registration time.

        Each registered widget carries a ``_toggle_kind`` attribute:
        ``"grid"`` (default) or ``"pack"``.  No runtime detection is
        attempted — Tk's introspection (``winfo_manager``, ``grid_info``)
        is unreliable after a widget has been hidden.
        """
        for w in widgets:
            kind = getattr(w, "_toggle_kind", "grid")
            try:
                if kind == "pack":
                    if show:
                        info = getattr(w, "_pack_info", None)
                        if info is not None:
                            w.pack(**info)
                    else:
                        try:
                            if not hasattr(w, "_pack_info"):
                                w._pack_info = w.pack_info()
                        except Exception:
                            pass
                        w.pack_forget()
                else:  # "grid"
                    if show:
                        w.grid()         # restores grid_remove'd state
                    else:
                        w.grid_remove()
            except Exception:
                pass

    # Default retardation factor lookup per PFAA — matches what the
    # Excel CalculrateRetardationFactors macro produces with default
    # Section 4 inputs (Clay, porosity 0.48, tortuosity 0.56).  Used as
    # a placeholder until the popups_retardation.exe pipeline is wired.
    # ── PFAA Koc / diffusion lookup (verbatim from Sheet3.cls/Sheet7.cls
    #    HandlePFAS_Generic) — source of truth for Section 5 retardation.
    #    Tuple is (Koc, molecular diffusion m²/sec or None).
    # v96: molecular diffusion exponents corrected from e-6 to e-10
    # — the original Excel macro had 3.52e-6 but that's m²/yr units;
    # the Fortran solver and §5 input both use m²/s, where PFOS
    # diffuses at ≈3.52e-10 m²/s.  Other species' values were shifted
    # by the same 10000× factor so their relative magnitudes remain
    # consistent with the Excel reference.
    PFAA_KOC = {
        "PFOS":           (631, 3.52e-10),
        "PFOA":           (200, 5.79e-10),
        "PFHxS":          (126, 4.5e-10),
        "PFHxA":          (40,  7.8e-10),
        "PFBS":           (35,  1.1e-9),
        "PFNA":           (398, None),
        "User-Specified": (None, None),
        "None":           (None, None),
    }
    # Section 4 K26 dropdown (Clay/Silt). Tortuosity formulas pulled
    # from Sheet3.cls Worksheet_Change handler:
    #   Clay : K27=0.48, K28 = 0.77 * exp(0.04 * ln(2.2 ** -10)) ≈ 0.56
    #   Silt : K27=0.48, K28 = 0.77 * exp(0.04 * ln(1.4 ** -7))  ≈ 0.70
    LOWK_DEFAULTS = {
        "Clay": (0.48, 0.56),
        "Silt": (0.48, 0.70),
    }

    def _read_retardation_file(self):
        """Read foc/rho_b/Koc list from retardation_inputs.txt.

        v102: search state.work_dir FIRST (that's where popups_retardation
        writes to), then fall back to BASE_DIR.  Previously the BASE_DIR-
        only lookup missed the fresh file the popup just wrote, so
        clicking "Calculate Retardation Factors" had no visible effect
        — the §5 Koc fell through to PFAA_KOC table defaults and the
        user's custom Koc / foc / bulk density input was ignored.

        Returns (rho_b, foc_t, foc_l, koc_list) or defaults if missing."""
        import os
        rho_b, foc_t, foc_l, koc_list = 1.7, 0.001, 0.002, []
        # v102: prefer the live work_dir (where popups write); fall back
        # to BASE_DIR (legacy hardcoded location).
        candidates = []
        try:
            from functions.state import get_state
            st = get_state()
            wd = getattr(st, "work_dir", "") or ""
            if wd:
                candidates.append(os.path.join(wd, "retardation_inputs.txt"))
        except Exception:
            pass
        candidates.append(os.path.join(BASE_DIR, "retardation_inputs.txt"))
        ret_file = None
        for c in candidates:
            if c and os.path.exists(c):
                ret_file = c; break
        if not ret_file or not os.path.exists(ret_file):
            return rho_b, foc_t, foc_l, koc_list
        try:
            with open(ret_file) as f:
                lines = [l.strip() for l in f if l.strip()]
            in_trans = in_lowk = False
            for line in lines:
                if "Transmissive Zone" in line:
                    in_trans, in_lowk = True, False
                elif "Low-K Zone" in line:
                    in_trans, in_lowk = False, True
                elif in_trans:
                    if line.startswith("Soil Bulk Density"):
                        try: rho_b = float(line.split(",")[1])
                        except Exception: pass
                    elif line.startswith("foc (-)"):
                        try: foc_t = float(line.split(",")[1])
                        except Exception: pass
                    elif line.startswith("Koc (L/kg)"):
                        for p in line.split(",")[1:]:
                            try: koc_list.append(float(p.strip()))
                            except Exception: koc_list.append(None)
                elif in_lowk:
                    if line.startswith("foc (-)"):
                        try: foc_l = float(line.split(",")[1])
                        except Exception: pass
        except Exception as e:
            print(f"Warning reading retardation_inputs.txt: {e}")
        return rho_b, foc_t, foc_l, koc_list

    # ── §5 General Molecular Diffusion Coefficient (v_mol_diff / E44) ──
    def _set_mol_diff(self, value):
        """Set v_mol_diff programmatically WITHOUT tripping the
        user-edit flag (see _on_mol_diff_edit)."""
        self._mol_diff_programmatic = True
        try:
            self.v_mol_diff.set(value)
        finally:
            self._mol_diff_programmatic = False

    def _on_mol_diff_edit(self, *_):
        """Trace on v_mol_diff — a real keystroke (not one of our own
        programmatic sets) flags a manual override that survives until
        the PFAA-1 species changes."""
        if getattr(self, "_building", False):
            return
        if getattr(self, "_mol_diff_programmatic", False):
            return
        self._mol_diff_user_edited = True

    def _update_mol_diff(self, species, diff, has_porosity):
        """Refill the §5 diffusion cell from the lookup table for the
        current PFAA-1 *species*, EXCEPT keep a manual override.  When
        the species changes the override is dropped and the new
        tabulated value (``diff``) is written."""
        if species != getattr(self, "_mol_diff_last_species", None):
            # PFAA-1 species changed → revert to the table value.
            self._mol_diff_last_species = species
            self._mol_diff_user_edited = False
        if getattr(self, "_mol_diff_user_edited", False):
            return                       # preserve the user's typed value
        if has_porosity and diff is not None:
            self._set_mol_diff(f"{diff:.2E}")
        else:
            self._set_mol_diff("")

    def _on_pfaa_change(self, *_):
        """Recompute Section 5 retardation.
        v105: skipped while the UI is being built so the §5 black cells
        stay empty on a clean-start form (the §5 dropdowns' cb.set fires
        this trace at startup).
        Uses retardation_inputs.txt (Koc, foc, rho_b) when present;
        falls back to PFAA_KOC table defaults otherwise.
        Formula: R = 1 + (rho_b * foc * Koc) / porosity"""
        if getattr(self, "_building", False):
            return
        try:
            G22 = float(self.v_porf.get())
        except (ValueError, TypeError):
            G22 = 0.0
        try:
            K27 = float(self.v_lowk_por.get())
        except (ValueError, TypeError):
            K27 = 0.0

        rho_b, foc_t, foc_l, koc_list = self._read_retardation_file()

        pairs = [
            (self.v_pfaa1, self.v_ret_trans1, self.v_ret_lowk1),
            (self.v_pfaa2, self.v_ret_trans2, self.v_ret_lowk2),
        ]
        if self.active_sheet == "Detailed_2":
            pairs += [
                (self.v_pfaa3, self.v_ret_trans3, self.v_ret_lowk3),
                (self.v_pfaa4, self.v_ret_trans4, self.v_ret_lowk4),
            ]
        for idx, (src_var, ret_t, ret_l) in enumerate(pairs):
            # Koc: prefer retardation_inputs.txt, fall back to lookup table
            koc = None
            if idx < len(koc_list) and koc_list[idx] and koc_list[idx] > 0:
                koc = koc_list[idx]
            species = src_var.get()
            # v96: precursor placeholders ("PFAA 1-able" / "PFAA 2-able")
            # don't have their own Koc entries in PFAA_KOC.  Resolve
            # them to the sibling PFAA's species so the §5 black cells
            # populate (Koc/diffusion of the precursor are taken to be
            # the same as the PFAA it converts to).
            if species == "PFAA 1-able":
                species = self.v_pfaa1.get()
            elif species == "PFAA 2-able":
                species = self.v_pfaa2.get()
            _, diff = self.PFAA_KOC.get(species, (None, None))
            if koc is None:
                koc_lookup, _ = self.PFAA_KOC.get(species, (None, None))
                koc = koc_lookup

            if koc is None or koc == 0:
                ret_t.set(""); ret_l.set("")
                # v102: also blank v_mol_diff when PFAA-1 is unset so
                # Clear All Data wipes the diffusion coefficient cell.
                # v106: route through _update_mol_diff so a manual
                # override is still honored / reverted on species change.
                if idx == 0:
                    try: self._update_mol_diff(species, diff, G22 > 0)
                    except Exception: pass
                continue
            # v105: these are CALCULATED (black) display cells.  Retardation
            # has the zone porosity in its denominator, so until the user has
            # entered that porosity (e.g. on a clean-start form) leave the
            # cell BLANK rather than writing a misleading R = 1.0.  Earlier
            # code used `... if G22 > 0 else 1.0`, which is why §5 showed 1.0
            # everywhere on startup and filled more cells when switching to
            # Detailed.  They now populate automatically once porosity exists.
            if G22 > 0:
                ret_t.set(f"{1.0 + (rho_b * foc_t * koc / G22):.1f}")
            else:
                ret_t.set("")
            if K27 > 0:
                ret_l.set(f"{1.0 + (rho_b * foc_l * koc / K27):.1f}")
            else:
                ret_l.set("")
            # Molecular diffusion (PFAA-1 only) — also a §5 black cell.  Keep
            # it blank until §5 is actually being configured (transmissive
            # porosity entered) so the clean-start form stays empty.
            if idx == 0:
                # v106: tabulated value unless the user typed an override
                # (which survives until the PFAA-1 species changes).
                self._update_mol_diff(species, diff, G22 > 0)

    # ── Decimal-formatting helpers ──────────────────────────────
    @staticmethod
    def _format_decimal_var(var, n: int):
        """Reformat the StringVar's value to *n* decimal places.
        Leaves the value untouched if it can't be parsed as a float
        or if it's empty."""
        s = (var.get() or "").strip()
        if not s:
            return
        try:
            f = float(s)
        except ValueError:
            return
        var.set(f"{f:.{n}f}")

    def _bind_decimal_format(self, entry, var, n: int):
        """Reformat *var* to *n* decimals when the user leaves the
        entry (FocusOut) or presses Return."""
        cb = lambda _=None, v=var, k=n: self._format_decimal_var(v, k)
        try:
            entry.bind("<FocusOut>", cb, add="+")
            entry.bind("<Return>",   cb, add="+")
        except Exception:
            pass

    def _on_lowk_media_change(self, *_):
        """K26 → K27 (porosity) + K28 (tortuosity), then cascade to
        retardation factors."""
        # v105: don't auto-fill while the UI is still being built (the
        # dropdown's cb.set fires this trace at startup) — clean start.
        if getattr(self, "_building", False):
            return
        por, tort = self.LOWK_DEFAULTS.get(self.v_lowk_media.get(),
                                           ("", ""))
        self.v_lowk_por.set(f"{por:.2f}" if isinstance(por, float) else str(por))
        self.v_lowk_tort.set(f"{tort:.2f}" if isinstance(tort, float) else str(tort))
        self._on_pfaa_change()

    def _on_porf_change(self, *_):
        """Trans. porosity G22 changed → recalc retardation."""
        self._on_pfaa_change()

    # Section 6 – dispersivity presets (matches popups_dispersivity.py values)
    _DISPERSIVITY_PRESETS = {
        "High":   (7.5,  0.04, 0.004),
        "Medium": (3.2,  0.04, 0.004),
        "Weak":   (1.1,  0.04, 0.004),
    }

    def _on_het_change(self, *_):
        """Heterogeneity radio changed → update longitudinal/transverse/vertical."""
        # v105: don't auto-fill dispersivity during UI build (clean start).
        if getattr(self, "_building", False):
            return
        # v106: suppressed while Load Data restores custom dispersivity
        # from dispersivity_inputs.txt, so the loaded alpha cells aren't
        # re-derived from a preset / cleared by this trace.
        if getattr(self, "_disp_loading", False):
            return
        preset = self._DISPERSIVITY_PRESETS.get(self.v_het.get())
        if preset is None:
            # "Enter Your Own Value Below" — clear cells so the user
            # types fresh values.  Cell colors stay sky blue; only
            # the text content is reset.
            self.v_alpha_l.set("")
            self.v_alpha_t.set("")
            self.v_alpha_v.set("")
            return
        al, at, av = preset
        # v105: presets are defined in METRES (matching popups_dispersivity).
        # If §1 units are feet, convert so the value matches the unit label
        # shown beside the cell.
        if self.v_units.get() == "feet":
            FT_PER_M = 1.0 / 0.3048
            al, at, av = al * FT_PER_M, at * FT_PER_M, av * FT_PER_M
        self.v_alpha_l.set(f"{al:.3f}")
        self.v_alpha_t.set(f"{at:.3f}")
        self.v_alpha_v.set(f"{av:.4f}")

    def _unit_len(self):
        """Return current length unit suffix string: 'm' or 'ft'."""
        try:
            return "ft" if self.v_units.get() == "feet" else "m"
        except Exception:
            return "m"

    def _register_unit_label(self, label_widget, fmt):
        """Register a label whose text contains a length-unit placeholder.

        fmt should contain '{u}' which gets replaced with 'm' or 'ft'.
        Example: '(${u})' or 'Distance from Source ({u})'.
        Label text is set immediately and then updated whenever the
        user toggles §1 feet/meters.
        """
        try:
            label_widget.config(text=fmt.format(u=self._unit_len()))
            self._unit_labels.append((label_widget, fmt))
        except Exception:
            pass

    def _register_length_var(self, var, kind="length"):
        """Register a StringVar that holds a length-valued number so
        its value converts when the user toggles §1 units.

        kind:
          'length' — meters ↔ feet  (×3.281 / ×0.3048)
          'rate'   — m/yr  ↔ ft/yr (same factor, length per time)
        """
        try:
            self._unit_length_vars.append((var, kind))
        except Exception:
            pass

    def _apply_units(self, *_):
        """v_units trace callback: refresh every registered length-unit
        label AND convert every registered numeric value to the new
        unit so what the user sees on screen stays internally consistent
        with the selected unit system.
        """
        try:
            new_unit = self.v_units.get()
        except Exception:
            return
        prev = getattr(self, "_prev_units", new_unit)
        # Compute conversion factor for the toggle
        # v102: use the EXACT reciprocal of FT2M=0.3048 so that
        # toggling units twice (m → ft → m) is bit-exactly idempotent.
        # Previous code used 3.28084 (rounded), so 500 m → 1640.42 ft
        # → 499.9999...m and rounding/formatting made the round-trip
        # drift visibly.  1 / 0.3048 = 3.280839895013123... — applying
        # factor and then 1/factor cancels exactly in IEEE-754.
        FT2M = 0.3048
        factor = None
        if prev == "meters" and new_unit == "feet":
            factor = 1.0 / FT2M        # m → ft
        elif prev == "feet" and new_unit == "meters":
            factor = FT2M              # ft → m
        # Update labels first (labels are visible immediately)
        u = self._unit_len()
        for w, fmt in list(self._unit_labels):
            try:
                w.config(text=fmt.format(u=u))
            except Exception:
                pass
        # Convert registered numeric values when toggling between units.
        # v102: format with `g`-style precision that preserves enough
        # digits so m → ft → m round-trips to the same string the user
        # started with.  Previous 2-decimal formatting caused visible
        # drift on the second toggle.
        def _fmt_round_trip(val):
            a = abs(val)
            if a == 0:
                return "0"
            # Use up to 6 significant digits, strip trailing zeros.
            s = f"{val:.6g}"
            return s
        if factor is not None:
            # Direction of this toggle, used to interpret the user's
            # current text and to render the canonical in the new unit.
            FT_PER_M = 1.0 / FT2M
            for var, kind in list(self._unit_length_vars):
                try:
                    raw = str(var.get()).strip()
                    if not raw:
                        continue
                    num = float(raw.replace(",", ""))
                except (ValueError, TypeError, AttributeError):
                    continue
                key = id(var)
                last = self._unit_last.get(key)
                if last is not None and raw == last and key in self._unit_canon:
                    # Untouched since we last set it → reuse the stored
                    # canonical so the round-trip is drift-free.
                    canon_m = self._unit_canon[key]
                else:
                    # First time, or user edited the field → derive the
                    # canonical (metres) from what's currently shown,
                    # interpreted in the PREVIOUS unit.
                    canon_m = num * FT2M if prev == "feet" else num
                # Render canonical in the NEW unit.
                disp = canon_m * FT_PER_M if new_unit == "feet" else canon_m
                s = _fmt_round_trip(disp)
                try:
                    var.set(s)
                except Exception:
                    pass
                self._unit_canon[key] = canon_m
                self._unit_last[key]  = s
        self._prev_units = new_unit

    # ── Run-Model validation gate ─────────────────────────────────────────
    def _collect_missing_inputs(self):
        """Return [(section_title, [missing-field labels]), ...] for every
        REQUIRED input left blank, taking the user's current options into
        account (Simple vs Detailed, 1 vs 2 PFAS, precursors, PSB, units).
        An empty list means the form is complete enough to run.

        The required set mirrors exactly the numeric fields that the
        clean-start blanks; advanced popup/sidecar parameters keep their
        own values and are not gated here.  Conditions match the mode
        logic in functions/generate_input_file.py and are grounded in the
        shipped Example inputs.
        """
        u = self._unit_len()   # 'm' or 'ft'

        def blank(v):
            try:
                return not str(v.get()).strip()
            except Exception:
                return True

        def chosen(var, *placeholders):
            """True if a dropdown holds a real selection (not blank/None)."""
            try:
                s = str(var.get()).strip()
            except Exception:
                s = ""
            return bool(s) and s.lower() not in [p.lower() for p in placeholders]

        detailed  = (getattr(self, "active_sheet", "Simple") == "Detailed_2")
        ncomp2    = chosen(self.v_pfaa2, "", "None")
        precursor = detailed and chosen(self.v_pfaa3, "", "None")
        try:
            psb_on = bool(self.v_model_psb.get())
        except Exception:
            psb_on = False

        missing = []

        def need(section, pairs):
            miss = [lbl for var, lbl in pairs if blank(var)]
            if miss:
                missing.append((section, miss))

        # §2 — Model domain & timeframe
        need("Section 2 — Model Domain & Timeframe", [
            (self.v_x_size,   f"Model length X ({u})"),
            (self.v_y_size,   f"Model width Y ({u})"),
            (self.v_z_size,   f"Model depth Z ({u})"),
            (self.v_sw_width, f"Source zone width ({u})"),
            (self.v_sw_thick, f"Source zone thickness ({u})"),
            (self.v_yr_start, "Simulation start year"),
            (self.v_yr_end,   "Simulation end year"),
        ])
        # §3 — Groundwater flow
        need("Section 3 — Groundwater Flow", [
            (self.v_darcy, f"Darcy velocity ({u}/yr)"),
            (self.v_porf,  "Transmissive-zone porosity"),
        ])
        # §4 — Low-permeability zone
        need("Section 4 — Low-Permeability Zone", [
            (self.v_lowk_por,  "Low-k porosity"),
            (self.v_lowk_tort, "Low-k tortuosity"),
        ])
        # §5 — Retardation & diffusion: the retardation factors and the
        # molecular-diffusion value are CALCULATED (black, read-only) cells,
        # NOT user inputs — they auto-fill from the PFAA species dropdowns
        # plus the §3/§4 porosities (both already required above).  Listing
        # them here would tell the user to "fill" a cell they can't type
        # into, so §5 is validated via its real inputs instead: the PFAA-1
        # species must be selected, and — in Detailed precursor mode — the
        # editable transformation-rate / yield-factor cells must be filled.
        s5 = [(self.v_pfaa1, "PFAA-1 species")]
        if precursor:
            s5 += [
                (self.v_trans_rate_3,   "Precursor 1 transformation rate (yr)"),
                (self.v_yield_factor_3, "Precursor 1 yield factor"),
            ]
            if ncomp2:
                s5 += [
                    (self.v_trans_rate_4,   "Precursor 2 transformation rate (yr)"),
                    (self.v_yield_factor_4, "Precursor 2 yield factor"),
                ]
        need("Section 5 — PFAS & Precursor Inputs", s5)
        # §6 — Dispersivity
        need("Section 6 — Dispersivity", [
            (self.v_alpha_l, f"Longitudinal dispersivity ({u})"),
            (self.v_alpha_t, f"Transverse dispersivity ({u})"),
            (self.v_alpha_v, f"Vertical dispersivity ({u})"),
        ])
        # §7 — Source loading table: every row that HAS a year needs its
        # concentration(s) for the active component set.
        s7 = []
        if all(blank(v) for v in self.v_src_years):
            s7.append("Source loading table is empty (enter years & concentrations)")
        else:
            for i in range(len(self.v_src_years)):
                if blank(self.v_src_years[i]):
                    continue
                yr = str(self.v_src_years[i].get()).strip()
                if blank(self.v_src_pfaa1[i]):
                    s7.append(f"Row {i+1} (year {yr}): PFAA-1 source concentration")
                if ncomp2 and blank(self.v_src_pfaa2[i]):
                    s7.append(f"Row {i+1} (year {yr}): PFAA-2 source concentration")
                if precursor and blank(self.v_src_pre1[i]):
                    s7.append(f"Row {i+1} (year {yr}): Precursor 1 source concentration")
                if precursor and ncomp2 and blank(self.v_src_pre2[i]):
                    s7.append(f"Row {i+1} (year {yr}): Precursor 2 source concentration")
        if s7:
            missing.append(("Section 7 — Source Loading", s7))
        # §11 — Output cadence
        need("Section 11 — Output", [
            (self.v_see_every, "See results every (yr)"),
        ])
        # §9 — PSB: required ONLY when the PSB option is turned on.
        if psb_on:
            psb = [
                (self.v_psb_yr,    "Year PSB installed"),
                (self.v_psb_dist,  f"PSB distance from source ({u})"),
                (self.v_psb_width, f"PSB width ({u})"),
                (self.v_psb_cells, "Number of cells in PSB"),
                (self.v_psb_load,  "PSB loading"),
                (self.v_psb_a_1,   "PSB Freundlich exponent a (PFAA-1)"),
                (self.v_psb_kf_1,  "PSB Freundlich Kf (PFAA-1)"),
            ]
            if ncomp2:
                psb += [(self.v_psb_a_2,  "PSB Freundlich exponent a (PFAA-2)"),
                        (self.v_psb_kf_2, "PSB Freundlich Kf (PFAA-2)")]
            need("Section 9 — Permeable Sorptive Barrier (PSB)", psb)

        return missing

    def _format_missing_message(self, missing):
        """Render the missing-input list into a clear, sectioned message."""
        total = sum(len(items) for _, items in missing)
        lines = [
            f"Cannot run the model yet — {total} required "
            f"input{'s' if total != 1 else ''} still need a value.",
            "These are the fields required for your current setup "
            "(version, number of PFAS, precursors, PSB, units):",
            "",
        ]
        for section, items in missing:
            lines.append(section)
            for it in items:
                lines.append(f"    • {it}")
            lines.append("")
        lines.append("Fill these in, or click \"Paste Example\" to load a "
                     "complete worked example, then Run Model again.")
        return "\n".join(lines)

    # ── Variable initialisation ───────────────────────────────────────────
    def _build_vars(self):
        self.v_model_version = tk.StringVar(value="Simple Version")
        self.v_model_version.trace_add('write', self._on_model_version_change)
        self.v_units         = tk.StringVar(value="meters")
        # v102: unit-aware label registry + value converters.
        # Every length-unit label that needs to flip between "(m)" and
        # "(ft)" registers itself via _register_unit_label(label, fmt).
        # Every numeric StringVar that should convert when the user
        # toggles feet/meters registers via _register_length_var(var,
        # kind), where kind is "length" (m↔ft) or "rate" (m/yr↔ft/yr).
        self._unit_labels = []        # list of (widget, fmt) tuples
        self._unit_length_vars = []   # list of (StringVar, kind)
        # v105: no-drift unit conversion.  For each registered length var
        # we remember (a) its canonical value in METRES and (b) the exact
        # string we last wrote into it.  On a feet<->meters toggle we
        # derive the new display FROM the canonical instead of chaining
        # conversions on already-rounded display text — so m->ft->m always
        # returns to the original number.  If the user has edited the field
        # since we last set it, we refresh the canonical from what they
        # typed (interpreted in the previous unit) before converting.
        self._unit_canon = {}   # id(var) -> canonical value in metres
        self._unit_last  = {}   # id(var) -> last string we set into var
        self._prev_units = self.v_units.get()
        self.v_units.trace_add("write", self._apply_units)
        self.v_site          = tk.StringVar(value="Test Case 1")
        self.v_date          = tk.StringVar(value=datetime.now().strftime("%b-%y"))

        # Section 2
        # v105: start CLEAN — numeric inputs blank so the user fills them
        # deliberately (pre-filled defaults previously caused confusion /
        # accidental runs on example values).  Dropdowns & text labels keep
        # their sensible defaults; only numbers are blanked.  Run Model and
        # all live recompute callbacks fall back to safe internal defaults
        # when a field is empty, so nothing crashes on a blank form.
        self.v_x_size   = tk.StringVar(value="")
        self.v_y_size   = tk.StringVar(value="")
        self.v_z_size   = tk.StringVar(value="")
        self.v_sw_width = tk.StringVar(value="")
        self.v_sw_thick = tk.StringVar(value="")
        self.v_yr_start = tk.StringVar(value="")
        self.v_yr_end   = tk.StringVar(value="")
        self.v_run_time = tk.StringVar(value="")

        # ── §2 "approx. run time" — live recompute ──────────────────────
        # Was a static "0.0" cell.  Now computed from the model
        # geometry + duration.  Empirical constant 5e-7 s per
        # (timestep × cell × component) tuned against Example/1_Simple
        # (~10 s for 1000×100×10×10×1 = 1e7 evals).  Multiply by 1.4
        # safety factor — slow machines / Detailed runs benefit.
        # Re-runs whenever any driver var changes.
        SECS_PER_EVAL = 5e-7 * 1.4

        def _safe_f(s, default=0.0):
            try:
                return float(str(s).replace(",", "").strip())
            except (ValueError, TypeError):
                return default

        def _recompute_run_time(*_):
            try:
                # v105: clean-start guard — if the essential geometry /
                # duration fields are blank, show nothing rather than an
                # estimate computed from hidden fallback numbers.
                _essential = (self.v_yr_start.get(), self.v_yr_end.get(),
                              self.v_x_size.get(), self.v_y_size.get(),
                              self.v_z_size.get())
                if any(not str(v).strip() for v in _essential):
                    self.v_run_time.set("")
                    return
                yr0 = _safe_f(self.v_yr_start.get(), 1977)
                yr1 = _safe_f(self.v_yr_end.get(),   yr0 + 100)
                xs  = _safe_f(self.v_x_size.get(), 500)
                ys  = _safe_f(self.v_y_size.get(),  50)
                zs  = _safe_f(self.v_z_size.get(),  10)
                # v105: make the estimate UNIT-INDEPENDENT.  The domain
                # sizes are shown in the user's selected unit (m or ft),
                # but the cell sizes below are metric.  Convert the domain
                # to meters first so the grid-cell count — and therefore
                # the estimate — is identical whether the form is in feet
                # or meters (it's only an estimate; consistency matters
                # more than the display unit).
                try:
                    _to_m = 0.3048 if self.v_units.get() == "feet" else 1.0
                except Exception:
                    _to_m = 1.0
                xs *= _to_m; ys *= _to_m; zs *= _to_m
                # Default cell sizes (metres) — match the cellsize_input.txt
                # defaults shipped with the examples.
                dx, dy, dz, dt = 5.0, 5.0, 1.0, 0.1
                nt  = max(1, int(round((yr1 - yr0) / dt)))
                nx  = max(1, int(round(xs / dx)))
                ny  = max(1, int(round(ys / dy)))
                nz  = max(1, int(round(zs / dz)))
                # ncomp + ipre — based on §5 PFAA dropdowns.  PFAA-2
                # != "None" → ncomp=2.  Detailed sheet & K38 set →
                # ipre=1.
                pfaa2 = str(getattr(self.v_pfaa2, "get", lambda: "")()).strip()
                ncomp = 1 if pfaa2 in ("", "None", "none") else 2
                ipre  = 1 if (self.active_sheet == "Detailed_2"
                              and str(self.v_pfaa3.get()).strip()
                                  not in ("", "None", "none")) else 0
                comps = max(1, ncomp + ipre)
                secs = nt * nx * ny * nz * comps * SECS_PER_EVAL
                mins = secs / 60.0
                if mins < 0.1:
                    self.v_run_time.set(f"{mins:.2f}")
                elif mins < 10:
                    self.v_run_time.set(f"{mins:.1f}")
                else:
                    self.v_run_time.set(f"{int(round(mins))}")
            except Exception:
                self.v_run_time.set("?")
        self._recompute_run_time = _recompute_run_time

        # Section 3  (v105: numeric inputs blank for clean start)
        self.v_darcy   = tk.StringVar(value="")
        self.v_porf    = tk.StringVar(value="")
        self.v_porf.trace_add("write",
                              lambda *_: self._on_porf_change())

        # Section 4  (dropdown keeps default; derived numbers start blank —
        # they auto-fill only when the user actively changes the media
        # dropdown, since constructor defaults don't fire the trace)
        self.v_lowk_media = tk.StringVar(value="Clay")
        self.v_lowk_media.trace_add("write",
                                    lambda *_: self._on_lowk_media_change())
        self.v_lowk_por   = tk.StringVar(value="")
        self.v_lowk_por.trace_add("write",
                                  lambda *_: self._on_pfaa_change())
        self.v_lowk_tort  = tk.StringVar(value="")

        # Section 5
        self.v_pfaa1      = tk.StringVar(value="PFOS")
        self.v_pfaa2      = tk.StringVar(value="None")
        self.v_pfaa3      = tk.StringVar(value="PFAA 1-able")  # K38 Precursor 1 (Detailed)
        self.v_pfaa4      = tk.StringVar(value="PFAA 2-able")  # M38 Precursor 2 (Detailed)
        self.v_ret_trans1 = tk.StringVar(value="")
        self.v_ret_lowk1  = tk.StringVar(value="")
        # PFAA-2 retardation factors (matching PFAA-1 columns)
        self.v_ret_trans2 = tk.StringVar(value="")
        self.v_ret_lowk2  = tk.StringVar(value="")
        # Precursor 1 & 2 retardation (Detailed mode only)
        self.v_ret_trans3 = tk.StringVar(value="")
        self.v_ret_lowk3  = tk.StringVar(value="")
        self.v_ret_trans4 = tk.StringVar(value="")
        self.v_ret_lowk4  = tk.StringVar(value="")
        self.v_mol_diff   = tk.StringVar(value="")
        # v106: §5 General Molecular Diffusion Coefficient may be manually
        # overridden by the user.  The override sticks until the PFAA-1
        # species changes, then reverts to that species' tabulated value.
        #   _mol_diff_user_edited  — True once the user types a value
        #   _mol_diff_last_species — PFAA-1 species the cell was filled for
        #   _mol_diff_programmatic — guard so OUR .set() calls don't count
        #                            as a user edit
        self._mol_diff_user_edited = False
        self._mol_diff_last_species = None
        self._mol_diff_programmatic = False
        self.v_mol_diff.trace_add("write", self._on_mol_diff_edit)
        # When user picks a different PFAA, recompute the retardation
        # cells from a default lookup (mirrors what the Excel
        # CalculrateRetardationFactors macro does — until we wire the
        # real EXE).
        self.v_pfaa1.trace_add("write", lambda *_: self._on_pfaa_change())
        self.v_pfaa2.trace_add("write", lambda *_: self._on_pfaa_change())
        self.v_pfaa3.trace_add("write", lambda *_: self._on_pfaa_change())
        self.v_pfaa4.trace_add("write", lambda *_: self._on_pfaa_change())

        # Display-mirror StringVars for §5 PFAA dropdowns: when the
        # underlying var is empty (user cleared the dropdown or loaded
        # blank state), these mirrors render "None" instead of "" so
        # places like calibration Step 2 always show a meaningful label
        # next to the radio button.  v_pfaa1 has no "None" choice so we
        # mirror it verbatim — but provide it for symmetry.
        def _mk_disp(src_var, fallback="None"):
            disp = tk.StringVar(value=(src_var.get() or fallback))
            def _sync(*_):
                v = src_var.get()
                disp.set(v if str(v).strip() else fallback)
            src_var.trace_add("write", _sync)
            return disp
        self._v_pfaa1_disp = _mk_disp(self.v_pfaa1, fallback="PFOS")
        self._v_pfaa2_disp = _mk_disp(self.v_pfaa2, fallback="None")
        self._v_pfaa3_disp = _mk_disp(self.v_pfaa3, fallback="None")
        self._v_pfaa4_disp = _mk_disp(self.v_pfaa4, fallback="None")

        # Section 6 – Dispersivity (top header)
        # v105: dropdown keeps its default; the dispersivity numbers start
        # blank (they auto-fill if the user re-selects a heterogeneity
        # preset, since the constructor value doesn't fire the trace).
        self.v_het     = tk.StringVar(value="Medium")
        self.v_alpha_l = tk.StringVar(value="")
        self.v_alpha_t = tk.StringVar(value="")
        self.v_alpha_v = tk.StringVar(value="")
        self.v_het.trace_add("write", self._on_het_change)

        # Section 7  (v105: numeric source grids blank for clean start)
        self.v_src_years = [tk.StringVar(value="") for _ in range(11)]
        self.v_src_pfaa1 = [tk.StringVar(value="") for _ in range(11)]
        self.v_src_pfaa2 = [tk.StringVar(value="") for _ in range(11)]
        # Detailed-only Precursor 1 / Precursor 2 year-grid source values
        # (XLSM cols Z and AB on Detailed_2 sheet)
        self.v_src_pre1  = [tk.StringVar(value="") for _ in range(11)]
        self.v_src_pre2  = [tk.StringVar(value="") for _ in range(11)]
        self.v_total_mass      = tk.StringVar(value="")
        self.v_total_mass_p2   = tk.StringVar(value="")
        self.v_total_mass_pre1 = tk.StringVar(value="")
        self.v_total_mass_pre2 = tk.StringVar(value="")

        # Section 8
        self.v_src_rem_yr   = tk.StringVar(value="")
        self.v_src_conc_red = tk.StringVar(value="")

        # Auto-apply §8 → §7: when the user types into §8 % or start
        # year, multiply §7 PFAA-1 / PFAA-2 (and Detailed precursor)
        # rows from the start year onward by (1 - %/100), pulled from
        # a per-row BASELINE so editing % is reversible (typing 5 then
        # 0 restores the original values).  Manual edits in §7 update
        # the baseline for that row so user overrides aren't fought.
        # Baselines snapshot the row's value as of the LAST non-§8
        # write — they're refreshed in restore_from_example /
        # restore_from_saved (see those modules for the call into
        # _refresh_s7_baseline).
        self._s7_baseline = {
            "pfaa1": [None] * 11,
            "pfaa2": [None] * 11,
            "pre1":  [None] * 11,
            "pre2":  [None] * 11,
        }
        self._s8_applying = False    # guard: programmatic §7 writes

        # Section 9
        self.v_model_psb   = tk.BooleanVar(value=False)
        self.v_psb_a_1     = tk.StringVar(value="")
        self.v_psb_a_2     = tk.StringVar(value="")
        self.v_psb_a_3     = tk.StringVar(value="")
        self.v_psb_a_4     = tk.StringVar(value="")
        self.v_psb_kf_1    = tk.StringVar(value="")
        self.v_psb_kf_2    = tk.StringVar(value="")
        self.v_psb_kf_3    = tk.StringVar(value="")
        self.v_psb_kf_4    = tk.StringVar(value="")
        self.v_psb_kf_conv  = tk.StringVar(value="")
        self.v_psb_kf_conv2 = tk.StringVar(value="")
        self.v_psb_kf_conv3 = tk.StringVar(value="")
        self.v_psb_kf_conv4 = tk.StringVar(value="")
        # v102: default Kf unit is (ug/kg)(ug/L)^(-a) — matches the
        # default concentration unit (µg/L) used everywhere else in
        # the app, so the freshly-loaded form is self-consistent.
        self.v_psb_kf_unit  = tk.StringVar(value="(ug/kg)(ug/L)^(-a)")
        # S molecular weight (g/mol) — only visible when v_psb_kf_unit is mol-based
        self.v_psb_mw_1     = tk.StringVar(value="")
        self.v_psb_mw_2     = tk.StringVar(value="")
        self.v_psb_mw_3     = tk.StringVar(value="")
        self.v_psb_mw_4     = tk.StringVar(value="")
        # §5 Detailed-only Transformation Rate + Yield Factor
        # (XLSM K41/M41 = Precursor 1/2 transformation rates,
        #  K42/M42 = yield factors)
        self.v_trans_rate_3   = tk.StringVar(value="")   # v105: clean start
        self.v_trans_rate_4   = tk.StringVar(value="")
        self.v_yield_factor_3 = tk.StringVar(value="")   # v105: clean start
        self.v_yield_factor_4 = tk.StringVar(value="")
        self.v_psb_yr      = tk.StringVar(value="")
        self.v_psb_width   = tk.StringVar(value="")      # v105: clean start
        self.v_psb_load    = tk.StringVar(value="")
        self.v_psb_dist    = tk.StringVar(value="")
        self.v_psb_cells   = tk.StringVar(value="")

        # Auto-enable "Model PSB?" the moment the user types ANY value
        # into a §9 cell.  Excel parity: in the workbook the checkbox
        # is decorative — values alone are enough to drive iwall=1.  In
        # our app the checkbox has been a hard gate, which surprised
        # users who filled in geometry but forgot to tick it.  This
        # trace closes the gap.
        def _maybe_enable_psb(*_):
            for v in (self.v_psb_dist, self.v_psb_width, self.v_psb_load,
                      self.v_psb_cells, self.v_psb_yr,
                      self.v_psb_a_1, self.v_psb_a_2,
                      self.v_psb_kf_1, self.v_psb_kf_2):
                try:
                    s = str(v.get()).strip()
                except Exception:
                    s = ""
                if s and s not in ("0", "0.0", "None"):
                    if not self.v_model_psb.get():
                        self.v_model_psb.set(True)
                    return
        for v in (self.v_psb_dist, self.v_psb_width, self.v_psb_load,
                  self.v_psb_cells, self.v_psb_yr,
                  self.v_psb_a_1, self.v_psb_a_2,
                  self.v_psb_kf_1, self.v_psb_kf_2):
            v.trace_add("write", _maybe_enable_psb)

        # Section 10  (v105: numeric calibration inputs blank; well NAME
        # labels kept as a convenient template since they're text, not
        # numbers — the user can overwrite them)
        self.v_sample_yr = tk.StringVar(value="")
        # v106: start blank — previously seeded with example well names
        # ("MW-504", "FS-MW504", …) which made §10 look pre-populated on a
        # clean start.  User types their own monitoring-well names.
        self.v_mw_names  = [tk.StringVar(value="") for _ in range(7)]
        self.v_mw_conc   = [tk.StringVar(value="") for _ in range(7)]
        self.v_mw_dist   = [tk.StringVar(value="") for _ in range(7)]
        # v102: monitoring-well distances are lengths → convert on unit toggle
        for _d in self.v_mw_dist:
            self._register_length_var(_d, "length")
        self.v_mw_conc2  = [tk.StringVar(value="") for _ in range(7)]

        # Section 11
        # v105: blank for clean start — the user fills it (or pastes an
        # example).  If left blank at Run Model, generate_input_file falls
        # back to V47=10 internally so the run still produces output.
        self.v_see_every = tk.StringVar(value="")

        # Image cache (Tk PhotoImage refs must be held to avoid GC)
        self._figures = {}

        # Wire §2 approx-run-time trace now that every driver var
        # exists (years, model size, PFAA dropdowns).
        for _drv in (self.v_yr_start, self.v_yr_end,
                     self.v_x_size,   self.v_y_size, self.v_z_size,
                     self.v_pfaa1,    self.v_pfaa2,
                     self.v_pfaa3,    self.v_pfaa4):
            try:
                _drv.trace_add("write", self._recompute_run_time)
            except Exception:
                pass
        # Initial paint
        try:
            self._recompute_run_time()
        except Exception:
            pass

        # v102: §7 source year cells (U8..U18) auto-fill from §2 start
        # and end years by linear interpolation across 11 rows so the
        # user only has to edit §2.  Set a guard so paste-example /
        # load-data writes don't trigger re-interpolation mid-restore.
        self._s7_years_filling = False

        def _refill_s7_years(*_):
            if getattr(self, "_s7_years_filling", False):
                return
            try:
                yr_start = float(str(self.v_yr_start.get()).strip())
                yr_end   = float(str(self.v_yr_end.get()).strip())
            except (ValueError, TypeError, AttributeError):
                return
            if yr_end <= yr_start:
                return
            self._s7_years_filling = True
            try:
                n = len(self.v_src_years)
                if n < 2:
                    return
                step = (yr_end - yr_start) / (n - 1)
                for i, var in enumerate(self.v_src_years):
                    val = yr_start + i * step
                    # Integer years for whole-year steps, else one decimal
                    if abs(val - round(val)) < 1e-6:
                        var.set(str(int(round(val))))
                    else:
                        var.set(f"{val:.1f}")
            finally:
                self._s7_years_filling = False

        self._refill_s7_years = _refill_s7_years
        self.v_yr_start.trace_add("write", _refill_s7_years)
        self.v_yr_end.trace_add("write", _refill_s7_years)
        # Initial fill so §7's defaults match §2's defaults.
        try:
            _refill_s7_years()
        except Exception:
            pass

    # ── UI builder ───────────────────────────────────────────────────────
    def _build_ui(self):
        p = self.inner

        self._build_top_bar(p)          # 1. Dark teal banner with logo

        # 2. Two-column body with a thick black vertical bar between them.
        # Header strip (Site/Date + Sec1 + Legend) lives INSIDE the left
        # column so the legend's right edge snaps to the central divider.
        body = tk.Frame(p, bg=BG_MAIN)
        body.pack(fill="both", expand=True, padx=8, pady=4)

        left    = tk.Frame(body, bg=BG_MAIN)
        # 0.05" thick black vertical divider between the two halves.
        try:
            divider_w = max(2, int(self.winfo_pixels("0.05i")))
        except Exception:
            divider_w = 5
        divider = tk.Frame(body, bg="#000000", width=divider_w)
        right   = tk.Frame(body, bg=BG_MAIN)

        # padx around the central divider — small breathing room so
        # left/right content never touches the black line.
        left.grid(   row=0, column=0, sticky="new",  padx=(0, 14))
        divider.grid(row=0, column=1, sticky="ns")
        right.grid(  row=0, column=2, sticky="new",  padx=(14, 0))
        body.columnconfigure(0, weight=1, uniform="halves")
        body.columnconfigure(1, minsize=divider_w)
        body.columnconfigure(2, weight=1, uniform="halves")

        # ── LEFT HALF ──────────────────────────────────────────────────
        # Header strip (Site/Date + Section 1 + Legend) sits at top of
        # the left column.  Legend's right edge therefore snaps to the
        # central divider.
        self._build_header_strip(left)
        self._hsep(left)
        self._build_s2_model_config(left)
        self._hsep(left)
        self._build_s3_gw_velocity(left)
        self._hsep(left)
        self._build_s4_hydrogeologic(left)
        self._hsep(left)
        self._build_s5_transport(left)
        self._hsep(left)   # closing line at the end of Section 5

        # ── RIGHT HALF ─────────────────────────────────────────────────
        # Section 6 (Heterogeneity) spans the top of the right half
        self._build_s6_heterogeneity(right)
        self._hsep(right)   # between Section 6 and Section 7/8

        # Section 7 + thick black vertical divider + Section 8
        s78 = tk.Frame(right, bg=BG_MAIN)
        s78.pack(fill="x", anchor="nw", pady=(0, 4))

        s7_col = tk.Frame(s78, bg=BG_MAIN)
        s7_col.grid(row=0, column=0, sticky="nw", padx=(0, 20))
        self._build_s7_source(s7_col)

        # Vertical bar between Section 7 and Section 8 (same thickness
        # as the body's half-divider).
        divider78 = tk.Frame(s78, bg="#000000", width=divider_w)
        divider78.grid(row=0, column=1, sticky="ns")

        s8_col = tk.Frame(s78, bg=BG_MAIN)
        s8_col.grid(row=0, column=2, sticky="nw", padx=(20, 0))
        self._build_s8_source_rem(s8_col)

        s78.columnconfigure(1, minsize=divider_w)
        s78.columnconfigure(2, weight=1)

        # Sections 9-11 stacked beneath, each filling the full half-width.
        self._hsep(right)   # between Section 7/8 row and Section 9
        self._build_s9_psb(right)
        self._build_s10_field_data(right)
        self._hsep(right)   # between Section 10 and Section 11
        self._build_s11_output(right)

        # Action cluster is now built inline inside _build_s11_output
        # (see the blue panel there) — no separate _build_action_row call.

        # ── Mid horizontal black bar ─────────────────────────────────
        midbar = tk.Frame(p, bg="#000000", height=6)
        midbar.pack(fill="x", padx=0, pady=(2, 2))

        # ── Bottom-body 2-col: empty | calibration panel ─────────────
        bot = tk.Frame(p, bg=BG_MAIN)
        bot.pack(fill="both", expand=True, padx=8, pady=4)

        bot_left  = tk.Frame(bot, bg=BG_MAIN)
        bot_div   = tk.Frame(bot, bg="#000000", width=divider_w)
        bot_right = tk.Frame(bot, bg=BG_MAIN)

        bot_left .grid(row=0, column=0, sticky="new",  padx=(0, 8))
        bot_div  .grid(row=0, column=1, sticky="ns")
        bot_right.grid(row=0, column=2, sticky="new",  padx=(8, 0))
        bot.columnconfigure(0, weight=1, uniform="halves")
        bot.columnconfigure(1, minsize=divider_w)
        bot.columnconfigure(2, weight=1, uniform="halves")

        self._build_calibration_panel(bot_right)

    # ─────────────────────────────────────────────────────────────────────
    # 1.  TOP DARK-TEAL BANNER  (#074F69)
    # ─────────────────────────────────────────────────────────────────────
    def _build_top_bar(self, parent):
        bar = tk.Frame(parent, bg=BG_HEADER_BAR, pady=6)
        bar.pack(fill="x")

        # Title styled per the Excel storyboard reference: "REMFluor-MD"
        # (the product name) is greenish (#DAF2D0); " Model Input Screen"
        # and "Vers. 2.6" are white.  All three labels use the bold
        # AppTitle / AppVersion fonts.
        title_fr = tk.Frame(bar, bg=BG_HEADER_BAR)
        title_fr.pack(side="left", padx=18)
        tk.Label(title_fr, text="REMFluor-MD",
                 font=FONT_TITLE, fg=FG_TITLE_GREEN, bg=BG_HEADER_BAR
                 ).pack(side="left")
        tk.Label(title_fr, text=" Model Input Screen",
                 font=FONT_TITLE, fg=FG_TITLE, bg=BG_HEADER_BAR
                 ).pack(side="left")

        tk.Label(bar, text="Vers. 2.6",
                 font=FONT_VERSION, fg=FG_TITLE, bg=BG_HEADER_BAR
                 ).pack(side="left", padx=10)

        # ── Zoom hint + buttons (between version and logo) ─────────────
        zoom_fr = tk.Frame(bar, bg=BG_HEADER_BAR)
        zoom_fr.pack(side="right", padx=12)
        tk.Label(zoom_fr, text="Zoom:",
                 font=FONT_LABEL_SM, fg=FG_TITLE, bg=BG_HEADER_BAR
                 ).pack(side="left", padx=(0, 4))
        for txt, cmd in [("−", self._zoom_out),
                         ("⟲", self._zoom_reset),
                         ("+", self._zoom_in)]:
            tk.Button(zoom_fr, text=txt, command=cmd,
                      font=FONT_BTN_CALIB, fg=FG_BTN_NAVY, bg=BTN_FILL,
                      width=2, padx=2, pady=0, bd=1,
                      relief="raised", cursor="hand2"
                      ).pack(side="left", padx=1)
        tk.Label(zoom_fr, text="(Ctrl±wheel)",
                 font=FONT_LABEL_XS, fg=FG_TITLE_GREEN, bg=BG_HEADER_BAR
                 ).pack(side="left", padx=(4, 0))

        # ── ESTCP logo from Figures/ ───────────────────────────────────
        # Render at the display's native pixel resolution so the logo stays
        # crisp on high-DPI screens. target_height is in pixels, not points.
        scale = getattr(self, "_dpi_scale", 1.0)
        self._logo_img = _load_logo_image(target_height=int(54 * scale))
        if self._logo_img is not None:
            tk.Label(bar, image=self._logo_img,
                     bg=BG_HEADER_BAR, bd=0).pack(side="right", padx=18)
        else:
            # Fallback "ESTCP" text badge if no image was found
            badge = tk.Label(bar, text="  ESTCP  ",
                             font=("Arial Black", 14), fg="#0F8B4C",
                             bg="#FFFFFF", padx=10, pady=2,
                             bd=2, relief="raised")
            badge.pack(side="right", padx=18)

    # ─────────────────────────────────────────────────────────────────────
    # 2.  HEADER STRIP – Site/ID + Date  (Legend & Heterogeneity moved into
    #                                    the body columns)
    # ─────────────────────────────────────────────────────────────────────
    def _build_header_strip(self, parent):
        # Top band of upper-left quadrant — TWO columns per Excel:
        #   LEFT  : Site/Date row + Section 1 STARTING INFO row stacked
        #   RIGHT : Legend block (rows 1-5)
        strip = tk.Frame(parent, bg=BG_MAIN, pady=4)
        strip.pack(fill="x", padx=8, pady=(2, 4))

        # ── LEFT column: Site/Date on top, Section 1 below ───────────
        left_col = tk.Frame(strip, bg=BG_MAIN)
        left_col.grid(row=0, column=0, sticky="nw")

        # --- Site / Date row ---
        sd = tk.Frame(left_col, bg=BG_MAIN)
        sd.pack(anchor="w", fill="x")
        tk.Label(sd, text="Site Location and ID:",
                 font=FONT_LABEL, bg=BG_MAIN).grid(row=0, column=0, sticky="w")
        tk.Label(sd, text="Date:", font=FONT_LABEL, bg=BG_MAIN
                 ).grid(row=0, column=1, sticky="w", padx=(14, 0))
        make_entry(sd, self.v_site, width=22, justify="left").grid(
            row=1, column=0, sticky="w", pady=2)
        make_entry(sd, self.v_date, width=10, justify="left").grid(
            row=1, column=1, sticky="w", padx=(14, 0), pady=2)

        # --- Section 1 STARTING INFO row (header + bordered radios) ---
        s1 = tk.Frame(left_col, bg=BG_MAIN)
        s1.pack(anchor="w", fill="x", pady=(8, 0))
        tk.Label(s1, text="1.  STARTING INFORMATION",
                 font=FONT_SECTION, fg=FG_SECTION, bg=BG_MAIN, anchor="w"
                 ).grid(row=0, column=0, sticky="w", columnspan=2)
        tk.Label(s1, text="Units?", font=FONT_LABEL_BI, bg=BG_MAIN
                 ).grid(row=0, column=2, sticky="w", padx=(40, 0))

        # Bordered Simple/Detailed radios block (Excel rectangle) —
        # thick black border per Excel reference.
        ver_box = tk.Frame(s1, bg=BG_MAIN, bd=2, relief="solid",
                           highlightbackground="#000000",
                           highlightthickness=0,
                           padx=8, pady=3)
        ver_box.grid(row=1, column=0, sticky="w", pady=2)
        big_radio(ver_box, "Simple Version",
                  self.v_model_version, "Simple Version").pack(anchor="w")
        big_radio(ver_box, "Detailed Version",
                  self.v_model_version, "Detailed Version").pack(anchor="w")
        help_link(s1, "OpenTable1").grid(row=1, column=1,
                                         sticky="w", padx=(2, 0))

        # Bordered feet/meters block — thick black border per Excel.
        u_box = tk.Frame(s1, bg=BG_MAIN, bd=2, relief="solid",
                         highlightbackground="#000000",
                         highlightthickness=0,
                         padx=8, pady=3)
        u_box.grid(row=1, column=2, sticky="w", padx=(40, 0), pady=2)
        big_radio(u_box, "feet",   self.v_units, "feet").pack(anchor="w")
        big_radio(u_box, "meters", self.v_units, "meters").pack(anchor="w")
        help_link(s1, "OpenTable1").grid(row=1, column=3,
                                         sticky="w", padx=(2, 0))

        # ── Legend block — left edge aligned with the central black
        #    divider that separates upper-left from upper-right halves.
        #    Spacer column 1 expands so legend snaps to the far right
        #    of the upper-left half.
        strip.columnconfigure(1, weight=1)
        legend_holder = tk.Frame(strip, bg=BG_MAIN)
        legend_holder.grid(row=0, column=2, sticky="ne")
        self._build_legend(legend_holder)

    # ─────────────────────────────────────────────────────────────────────
    # LEGEND BLOCK  (top-right of upper-left quadrant, per Excel storyboard)
    # Layout matches the Excel reference:
    #   Row 0: "Legend" title  ............................  ☐ Check Box
    #   Row 1: [white  ] Enter value directly             ..  [Button]
    #   Row 2: [blue   ] Cell with formula or default, but ok to overwrite**
    #   Row 3: [peach  ] Pull Down Menu
    #   Row 4: [black  ] Calculated value or taken from other cell.
    # ─────────────────────────────────────────────────────────────────────
    def _build_legend(self, parent):
        outer = tk.Frame(parent, bg=BG_LEGEND_BAR, bd=1, relief="solid")
        outer.pack(anchor="nw")

        # ── Title bar (dark teal, white italic) ──────────────────────
        title = tk.Frame(outer, bg=BG_HEADER_BAR, padx=4, pady=1)
        title.pack(fill="x")
        tk.Label(title, text="Legend",
                 font=FONT_LABEL_BI, bg=BG_HEADER_BAR, fg=FG_TITLE
                 ).pack(side="left", padx=(2, 12))
        tk.Label(title, text="☑", font=FONT_LABEL_B,
                 bg=BG_HEADER_BAR, fg="#7FCB7B"
                 ).pack(side="left")
        tk.Label(title, text="Check Box", font=FONT_LABEL_I,
                 bg=BG_HEADER_BAR, fg=FG_TITLE
                 ).pack(side="left", padx=(2, 12))
        tk.Label(title, text="Button", font=FONT_LABEL_BI,
                 bg=BG_HEADER_BAR, fg="#9CC3E8"
                 ).pack(side="left", padx=(2, 4))

        # ── Body (white) — 4 swatch rows ─────────────────────────────
        body = tk.Frame(outer, bg=BG_LEGEND_BAR, padx=4, pady=2)
        body.pack(fill="x")

        legend_items = [
            (BG_INPUT_BLUE, "",             "Enter value directly"),
            (BG_FORMULA,    "",             "Cell with formula or default, but ok to overwrite**"),
            (BG_PULLDOWN,   "Pull Down",    "Pull Down Menu"),
            (BG_LOCKED,     "Can't change", "Calculated value or taken from other cell."),
        ]
        for i, (color, swatch_text, descr) in enumerate(legend_items):
            fg_swatch = "#FFFFFF" if color == "#000000" else FG_INPUT
            sw = tk.Label(body, text=swatch_text,
                          font=FONT_LABEL_BI, bg=color, fg=fg_swatch,
                          bd=1, relief="solid", width=12, anchor="center")
            sw.grid(row=i, column=0, sticky="w", padx=2, pady=1)
            tk.Label(body, text=descr, font=FONT_LABEL_I,
                     bg=BG_LEGEND_BAR, fg=FG_INPUT, anchor="w"
                     ).grid(row=i, column=1, sticky="w", padx=6)

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 6 – GEOLOGIC HETEROGENEITY  (right column, top)
    # ─────────────────────────────────────────────────────────────────────
    def _build_s6_heterogeneity(self, parent):
        _hdr6 = tk.Frame(parent, bg=BG_MAIN)
        _hdr6.pack(anchor="w", pady=(2, 1), fill="x")
        section_header(_hdr6, "6",
                       "PLUME TRANSPORT – DISPERSIVITY"
                       ).pack(side="left")
        help_link(_hdr6, "OpenTable6_0").pack(side="left", padx=(6, 0))

        # All controls live in a single grid:
        #   row 0  – heterogeneity radios + "Enter Your Own Value Below"
        #   row 1  – column titles (Longitudinal / Transverse / Vertical)
        #   row 2  – "Values:" label + the three cells
        #
        # Column layout:
        #   col 0   – row labels  ("Geologic Heterogeneity is:" / "Values:")
        #   col 1-3 – content (radios on row 0, titles on row 1, cells on row 2)
        #   col 4   – "Enter Your Own Value Below" hint  (row 0 only)
        #   col 5   – How to Decide?  (rowspan=3 → vertically centered)
        #   col 6   – Section5_1.png  (rowspan=3 → vertically centered)
        form = tk.Frame(parent, bg=BG_MAIN)
        form.pack(fill="x", pady=(0, 4))

        # ── Row 0 – Heterogeneity radios ───────────────────────────
        tk.Label(form, text="Geologic Heterogeneity is:",
                 font=FONT_LABEL, bg=BG_MAIN
                 ).grid(row=0, column=0, sticky="e", padx=(0, 4))

        het_fr = tk.Frame(form, bg=BG_MAIN)
        het_fr.grid(row=0, column=1, columnspan=4, sticky="w")
        for val in ["High", "Medium", "Weak", "Enter Your Own Value Below"]:
            big_radio(het_fr, val, self.v_het, val
                      ).pack(side="left", padx=4)

        # ── Row 1 – Titles ABOVE each cell ─────────────────────────
        # v102: dispersivity values are lengths → labels flip m↔ft.
        title_fmts = ["Longitudinal ({u})", "Transverse ({u})", "Vertical ({u})"]
        for i, fmt in enumerate(title_fmts):
            _dlbl = tk.Label(form, text=fmt.format(u=self._unit_len()),
                             font=FONT_LABEL_SM, bg=BG_MAIN)
            _dlbl.grid(row=1, column=1+i, sticky="w", padx=(8, 4))
            self._register_unit_label(_dlbl, fmt)

        # ── Row 2 – Values label + entry cells (start at same X as
        #             the heterogeneity radios — both in column 1). ──
        tk.Label(form, text="Values:", font=FONT_LABEL, bg=BG_MAIN
                 ).grid(row=2, column=0, sticky="e", padx=(0, 4))
        for i, var in enumerate([self.v_alpha_l,
                                  self.v_alpha_t,
                                  self.v_alpha_v]):
            make_entry(form, var, width=8, bg=BG_FORMULA
                       ).grid(row=2, column=1+i, sticky="w", padx=(8, 4))
            # v102: dispersivity is a length — convert when units toggle
            self._register_length_var(var, "length")

        # ── How to Decide? button — vertically centered between
        #     heterogeneity row and values row (rowspan=3 with no
        #     sticky lets the grid auto-center it). ──────────────
        make_btn(form, "How to Decide?",
                 "OpenAppendix_6_1_Relative", quarto=True,
                 font=FONT_BTN_SM, width=14, bg=BTN_FILL
                 ).grid(row=0, column=5, rowspan=3, padx=(12, 0))

        # ── Section5_1.png next to the How to Decide button ───────
        scale = getattr(self, "_dpi_scale", 1.0)
        img = _load_figure("Section5_1.png", target_height=int(80 * scale))
        if img is None:
            # Fallback: previous filename in case the file is still
            # named Section6_1.png in the Figures/ folder.
            img = _load_figure("Section6_1.png",
                               target_height=int(80 * scale))
        self._figures["s5_1"] = img
        if img is not None:
            tk.Label(form, image=img, bg=BG_MAIN, bd=0
                     ).grid(row=0, column=6, rowspan=3, padx=(8, 0))

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 1 – STARTING INFORMATION
    # ─────────────────────────────────────────────────────────────────────
    def _build_s1_starting(self, parent):
        # One grid frame holds everything so the "Units" title can sit on
        # row 0 — the same vertical line as "STARTING INFORMATION".
        #
        # Column map:
        #   0 : "STARTING INFORMATION" header
        #   1 : 1-inch spacer (forced via columnconfigure minsize)
        #   2 : Simple / Detailed radios          (rows 1-2)
        #   3 : '?' for versions  (rowspan=2  -> vertically centered)
        #   4 : 2-inch spacer (forced via columnconfigure minsize)
        #   5 : "Units" title  (row 0)  +  feet / meters  (rows 1-2)
        #   6 : '?' for Units    (rowspan=2  -> vertically centered)
        f = tk.Frame(parent, bg=BG_MAIN)
        f.pack(fill="x", pady=(2, 4))

        # Force the spacer columns to be exactly 0.5" / 0.5" wide.
        f.columnconfigure(1, minsize="0.5i")
        f.columnconfigure(4, minsize="0.5i")

        # Row 0 – titles on the same horizontal line
        tk.Label(f, text="1.  STARTING INFORMATION",
                 font=FONT_SECTION, fg=FG_SECTION, bg=BG_MAIN, anchor="w"
                 ).grid(row=0, column=0, sticky="w")
        tk.Label(f, text="Units", font=FONT_LABEL_BI, bg=BG_MAIN
                 ).grid(row=0, column=5, sticky="w")

        # Rows 1-2 – Simple / Detailed radios in col 2
        big_radio(f, "Simple Version",
                  self.v_model_version, "Simple Version"
                  ).grid(row=1, column=2, sticky="w")
        big_radio(f, "Detailed Version",
                  self.v_model_version, "Detailed Version"
                  ).grid(row=2, column=2, sticky="w")

        # '?' for versions – col 3, rowspan covers feet/meters rows
        help_link(f, "OpenTable1").grid(
            row=1, column=3, rowspan=2, padx=(2, 0))

        # Rows 1-2 – feet / meters in col 5 (same column as the Units title)
        big_radio(f, "feet",   self.v_units, "feet"
                  ).grid(row=1, column=5, sticky="w")
        big_radio(f, "meters", self.v_units, "meters"
                  ).grid(row=2, column=5, sticky="w")

        # '?' for Units – col 6, rowspan covers feet/meters rows
        help_link(f, "OpenTable1").grid(
            row=1, column=6, rowspan=2, padx=(2, 0))

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 2 – MODEL CONFIGURATION
    # ─────────────────────────────────────────────────────────────────────
    def _build_s2_model_config(self, parent):
        section_header(parent, "2", "MODEL CONFIGURATION").pack(
            anchor="w", pady=(4, 1))

        container = tk.Frame(parent, bg=BG_MAIN)
        container.pack(fill="x", pady=(0, 4))

        # ── LEFT: just the form. Labels on a SINGLE line (no wraplength)
        #         so the form takes its natural full width and pushes the
        #         images further to the right. ─────────────────────────
        left_col = tk.Frame(container, bg=BG_MAIN)
        left_col.pack(side="left", anchor="n")

        f = tk.Frame(left_col, bg=BG_MAIN)
        f.pack(fill="x", pady=(0, 4))

        # "Model Size" sub-label sits on row 0, in the column where the
        # value cells live (column 1) — so it labels the input column.
        tk.Label(f, text="Model Size", font=FONT_LABEL_I,
                 bg=BG_MAIN, fg=FG_GREY
                 ).grid(row=0, column=1, sticky="w", padx=2)

        # v102: "(m)" entries use "({u})" so _apply_units flips them to
        # "(ft)" when the user picks feet in §1.  Year units stay literal.
        rows = [
            ("Model Size in Direction of Groundwater Flow (X Direction)", self.v_x_size, "({u})",     "OpenTable2_1_XDirection"),
            ("Model Width Perpendicular to Flow (Y Direction)",           self.v_y_size, "({u})",     "OpenTable2_1_YDirection"),
            ("Model Depth Below Water Table (Z Direction)",               self.v_z_size, "({u})",     "OpenTable2_1_ZDirection"),
            ("Source Width (REMChlor-MD will round to nearest whole cell)", self.v_sw_width, "({u})", "OpenTable2_3"),
            ("Thickness of Source Below Water Table",                     self.v_sw_thick, "({u})",   "OpenTable2_4"),
            ("Starting Year of Simulation (year the source started)",     self.v_yr_start, "(YYYY year)", "OpenTable2_5"),
            ("Ending Year of Simulation",                                 self.v_yr_end,   "(YYYY year)", "OpenTable2_6"),
        ]
        # v102: register §2 length-valued StringVars so values convert
        # when user toggles feet/meters.
        for _len_var in (self.v_x_size, self.v_y_size, self.v_z_size,
                         self.v_sw_width, self.v_sw_thick):
            self._register_length_var(_len_var, "length")
        for i, (lbl, var, unit, helpm) in enumerate(rows):
            # Extra top-padding (~3× previous) to visually separate the
            # three logical groups (model dims | source | sim years):
            #   i = 3 → above "Source Width"   (after Model Depth)
            #   i = 5 → above "Starting Year"  (after Thickness)
            pady_top = 30 if i in (3, 5) else 1
            pady     = (pady_top, 1)

            # Right-align the descriptive label so all entries line up.
            tk.Label(f, text=lbl, font=FONT_LABEL, bg=BG_MAIN, anchor="e"
                     ).grid(row=i+1, column=0, sticky="e",
                            pady=pady, padx=(0, 4))
            make_entry(f, var, width=8).grid(
                row=i+1, column=1, sticky="w", pady=pady)
            # v102: if the unit string contains "{u}" placeholder,
            # register the label so it flips m↔ft when units toggle.
            _ulbl = tk.Label(f, text=unit, font=FONT_LABEL_SM, bg=BG_MAIN)
            _ulbl.grid(row=i+1, column=2, sticky="w", padx=2, pady=pady)
            if "{u}" in unit:
                self._register_unit_label(_ulbl, unit)
            help_link(f, helpm).grid(row=i+1, column=3, sticky="w",
                                      pady=pady)

        # ── RIGHT: image area + controls, all positioned with place()
        #          so each control can sit RIGHT of the image it
        #          relates to (Optional next to Section2_1, run-time +
        #          Estimating next to Section2_2). Because Section2_2
        #          is narrower than Section2_1, run-time / Estimating
        #          start further LEFT than Optional. ────────────────
        right_area = tk.Frame(container, bg=BG_MAIN)
        right_area.pack(side="left", anchor="n", padx=(14, 0))

        scale = getattr(self, "_dpi_scale", 1.0)
        img_h = int(110 * scale)
        try:
            overlap_px = int(self.winfo_pixels("0.5i"))
            gap_px     = int(self.winfo_pixels("0.15i"))
            shift_px   = int(self.winfo_pixels("0.2i"))   # vertical nudge
        except Exception:
            overlap_px = int(48 * scale)
            gap_px     = int(14 * scale)
            shift_px   = int(19 * scale)

        self._figures["s2_1"] = _load_figure("Section2_1.png", target_height=img_h)
        self._figures["s2_2"] = _load_figure("Section2_2.png", target_height=img_h)

        img1 = self._figures.get("s2_1")
        img2 = self._figures.get("s2_2")
        w1 = img1.width()  if img1 else 200
        h1 = img1.height() if img1 else img_h
        w2 = img2.width()  if img2 else 150
        h2 = img2.height() if img2 else img_h

        # Container needs explicit width / height for place() to work.
        # Reserve ~380 px on the right of the wider image for buttons —
        # the rounded "Optional: Enter user defined size of grid cells"
        # button is wider than the previous flat tk.Button.
        ctrl_w = 380
        box_w  = max(w1, w2) + gap_px + ctrl_w
        box_h  = h1 + h2 - overlap_px          # vertical overlap of 0.5"

        img_box = tk.Frame(right_area, bg=BG_MAIN, width=box_w, height=box_h)
        img_box.pack(anchor="nw")
        img_box.pack_propagate(False)

        # Section2_1 at top-left
        if img1 is not None:
            tk.Label(img_box, image=img1, bg=BG_MAIN, bd=0
                     ).place(x=0, y=0)

        # Section2_2 below Section2_1, overlapping by 0.5", drawn on top
        if img2 is not None:
            lbl2 = tk.Label(img_box, image=img2, bg=BG_MAIN, bd=0)
            lbl2.place(x=0, y=h1 - overlap_px)
            lbl2.lift()

        # ── Optional button: top-right, next to Section2_1 ───────────
        # Starts at x = w1 + gap (right edge of Section2_1 + gap).
        opt_btn = make_btn(img_box,
                           "Optional: Enter user defined\nsize of grid cells",
                           "OpenAppendix_2_1_Relative", quarto=True,
                           font=FONT_BTN_SM, width=24, bg=BTN_FILL)
        opt_btn.place(x=w1 + gap_px, y=8 + shift_px)   # 0.2" lower

        # ── approx run time: next to Section2_2.  Because Section2_2
        #    is narrower, this starts further LEFT than the Optional
        #    button does (x = w2 + gap). ──────────────────────────
        rt = tk.Frame(img_box, bg=BG_MAIN)
        tk.Label(rt, text="approx. run time:", font=FONT_LABEL, bg=BG_MAIN,
                 fg=FG_GREY).pack(side="left")
        tk.Label(rt, textvariable=self.v_run_time, font=FONT_LABEL,
                 bg=BG_LOCKED, fg=FG_LOCKED, width=14, relief="solid", bd=1
                 ).pack(side="left", padx=4)
        tk.Label(rt, text="minutes", font=FONT_LABEL_SM, bg=BG_MAIN,
                 fg=FG_GREY).pack(side="left")
        # 0.2" lower than its previous position
        rt_y = h1 - overlap_px + 4 + shift_px
        rt.place(x=w2 + gap_px, y=rt_y)

        # ── Estimating the Source Start Time: under run-time, also at
        #    x = w2 + gap (next to Section2_2). Single line, no wrap.
        #    0.2" gap between the run-time label and this button.
        est_btn = make_btn(img_box, "Estimating the Source Start Time",
                           "OpenAppendix_2_2_Relative", quarto=True,
                           font=FONT_BTN_SM, width=30, bg=BTN_FILL)
        est_btn.place(x=w2 + gap_px, y=rt_y + 22 + shift_px)

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 3 – GW DARCY VELOCITY
    # ─────────────────────────────────────────────────────────────────────
    def _build_s3_gw_velocity(self, parent):
        # v102: header includes a unit suffix that flips with §1 units.
        # section_header returns a tk.Label directly; the text format
        # is "{num}.  {text}".  Register the label so its full text
        # updates when units toggle.
        _s3_hdr = section_header(parent, "3",
            f"GROUNDWATER DARCY VELOCITY (Vd) ({self._unit_len()}/yr)")
        _s3_hdr.pack(anchor="w", pady=(6, 1))
        self._register_unit_label(_s3_hdr,
            "3.  GROUNDWATER DARCY VELOCITY (Vd) ({u}/yr)")

        f = tk.Frame(parent, bg=BG_MAIN)
        f.pack(fill="x", pady=(0, 4))

        # ── Row 0: short titles ON TOP of the cells, with '?' next ──
        # Col 0: empty (long Vd description sits here on row 1)
        _vd_lbl = tk.Label(f, text=f"Vd ({self._unit_len()}/yr)",
                 font=FONT_LABEL, bg=BG_MAIN, anchor="w")
        _vd_lbl.grid(row=0, column=1, sticky="w")
        self._register_unit_label(_vd_lbl, "Vd ({u}/yr)")
        # v102: v_darcy is a velocity (length/time) — converts like length
        self._register_length_var(self.v_darcy, "rate")
        help_link(f, "OpenTable3_1").grid(row=0, column=2,
                                          sticky="w", padx=(2, 0))

        tk.Label(f, text="Trans. Effective Porosity (-)",
                 font=FONT_LABEL, bg=BG_MAIN, anchor="w"
                 ).grid(row=0, column=4, sticky="w", padx=(20, 0))
        help_link(f, "OpenTable3_2").grid(row=0, column=5,
                                          sticky="w", padx=(2, 0))

        # ── Row 1: long Vd description (LEFT of cell), Vd cell,
        #          Trans cell, OR, Calculator button ───────────────
        tk.Label(f, text="Bulk Groundwater Darcy Velocity (Vd)  (Vd = K·dh/dx)",
                 font=FONT_LABEL, bg=BG_MAIN, anchor="e"
                 ).grid(row=1, column=0, sticky="e", padx=(0, 6),
                        pady=(2, 0))

        # Vd cell sized to roughly match the "Vd (m/yr)" title above it
        make_entry(f, self.v_darcy, width=12).grid(
            row=1, column=1, columnspan=2, sticky="w", pady=(2, 0))

        # Trans. Effective Porosity cell — doubled width
        make_entry(f, self.v_porf, width=20).grid(
            row=1, column=4, columnspan=2, sticky="w",
            padx=(20, 0), pady=(2, 0))

        # OR between Trans. Effective Porosity cell and Calculator button
        tk.Label(f, text="OR", font=FONT_LABEL_BI,
                 bg=BG_MAIN, fg=FG_GREY
                 ).grid(row=1, column=6, padx=8, pady=(2, 0))

        # Groundwater Velocity Calculator button (right of OR)
        make_btn(f, "Groundwater Velocity Calculator",
                 "GWVelocityCalculator", font=FONT_BTN_SM, width=30,
                 bg=BTN_FILL).grid(row=1, column=7, padx=(0, 0),
                                    pady=(2, 0))

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 4 – HYDROGEOLOGIC / MATRIX DIFFUSION
    # ─────────────────────────────────────────────────────────────────────
    def _build_s4_hydrogeologic(self, parent):
        # v105: flush-left (padx 0) so Section 4 aligns with Sections 3 and 5,
        # which pack their headers directly into the column with no indent.
        # (Was indented 0.25" from the left edge, which made it look offset.)
        s4 = tk.Frame(parent, bg=BG_MAIN)
        s4.pack(fill="x", padx=0, pady=(0, 4))

        section_header(s4, "4",
                       "HYDROGEOLOGIC SETTING AND MATRIX DIFFUSION"
                       ).pack(anchor="w", pady=(2, 1))

        f = tk.Frame(s4, bg=BG_MAIN)
        f.pack(fill="x", pady=(0, 4))

        # ── LEFT: heterogeneity calculator buttons, vertically
        #          centered against the image's height. The two
        #          weight=1 spacer rows (above & below the buttons)
        #          push the button pair to the middle of bcol. ────
        bcol = tk.Frame(f, bg=BG_MAIN)
        bcol.grid(row=0, column=0, sticky="ns")
        bcol.rowconfigure(0, weight=1)   # top spacer
        bcol.rowconfigure(3, weight=1)   # bottom spacer

        make_btn(bcol, "Use Unconsolidated Media\nHeterogeneity Calculator",
                 "HeterogeneityCalculator_Unconsolidated_Media",
                 font=FONT_BTN_SM, width=30, bg=BTN_FILL, padx=4, pady=4
                 ).grid(row=1, column=0, sticky="w")
        # 0.1" gap between the two "Use" buttons
        make_btn(bcol, "Use Fractured Rock\nHeterogeneity Calculator",
                 "HeterogeneityCalculator_Fractured_Rock",
                 font=FONT_BTN_SM, width=30, bg=BTN_FILL, padx=4, pady=4
                 ).grid(row=2, column=0, sticky="w", pady=("0.1i", 0))

        # ── MIDDLE: Section4_1.png with 1.5" gaps on either side
        #            (6× the previous 0.25" — separating it from the
        #            buttons on the left and the Low-k Media block on
        #            the right). ──────────────────────────────────
        mcol = tk.Frame(f, bg=BG_MAIN)
        mcol.grid(row=0, column=1, sticky="n",
                  padx=("1.5i", "1.5i"))

        scale = getattr(self, "_dpi_scale", 1.0)
        self._figures["s4_1"] = _load_figure("Section4_1.png",
                                             target_height=int(140 * scale))
        if self._figures["s4_1"] is not None:
            tk.Label(mcol, image=self._figures["s4_1"],
                     bg=BG_MAIN, bd=0).pack(anchor="n")

        # ── RIGHT: Low-k Media Details + "What is a Low-k" button ──
        rcol = tk.Frame(f, bg=BG_MAIN)
        rcol.grid(row=0, column=2, sticky="nw")

        tk.Label(rcol, text="Low-k Media Details",
                 font=FONT_LABEL_BI,
                 bg=BG_SECTION_HDR, anchor="w", padx=4
                 ).grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 2))

        rows = [
            ("Low-k Zone Media",          self.v_lowk_media, None,  "OpenTable4_1"),
            ("Low-k Zone Total Porosity", self.v_lowk_por,   "(-)", "OpenTable4_2"),
            ("Low-k Zone Pore Tortuosity",self.v_lowk_tort,  "(-)", "OpenTable4_3"),
        ]
        for i, (lbl, var, unit, helpm) in enumerate(rows):
            tk.Label(rcol, text=lbl, font=FONT_LABEL, bg=BG_MAIN
                     ).grid(row=i+1, column=0, sticky="w", pady=1)
            if i == 0:
                # Excel data-validation for K26: Clay or Silt only.
                dropdown(rcol, var,
                         ["Clay", "Silt"],
                         width=10, bg=BG_PULLDOWN
                         ).grid(row=i+1, column=1, padx=4, sticky="w")
            else:
                # Total Porosity / Pore Tortuosity – Formula legend color
                e_s4 = make_entry(rcol, var, width=8, bg=BG_FORMULA)
                e_s4.grid(row=i+1, column=1, padx=4, sticky="w")
                # Section 4: enforce 2-decimal display on focus-out / Return.
                self._bind_decimal_format(e_s4, var, 2)
            if unit:
                tk.Label(rcol, text=unit, font=FONT_LABEL_SM, bg=BG_MAIN
                         ).grid(row=i+1, column=2, sticky="w")
            help_link(rcol, helpm).grid(row=i+1, column=3, sticky="w", padx=2)

        # 0.1" gap between Low-k Zone Pore Tortuosity (row 3) and the
        # "What is a Low-k Geologic Unit?" button (row 4).
        make_btn(rcol, 'What is a "Low-k" Geologic Unit?',
                 "OpenAppendix_4_2_Relative", quarto=True,
                 font=FONT_BTN_SM, width=32, bg=BTN_FILL).grid(
            row=4, column=0, columnspan=4, sticky="w", pady=("0.1i", 0))

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 5 – PFAS TRANSPORT PROPERTIES
    # ─────────────────────────────────────────────────────────────────────
    def _build_s5_transport(self, parent):
        section_header(parent, "5", "PFAS TRANSPORT PROPERTIES").pack(
            anchor="w", pady=(4, 1))

        # Centered form — labels right-aligned in column 0, entries
        # share columns 1 + 2.  Per Excel storyboard the whole block is
        # centered horizontally, not flush-left.
        f = tk.Frame(parent, bg=BG_MAIN)
        f.pack(anchor="center", pady=(0, 4))

        # Excel data-validation lists for E38 / G38 (PFAA-1 / PFAA-2)
        PFAA1_CHOICES = ["PFOS", "PFOA", "PFHxS", "PFHxA", "PFBS", "PFNA", "User-Specified"]
        PFAA2_CHOICES = ["None"] + PFAA1_CHOICES

        # ── Column headers ────────────────────────────────────────────────
        tk.Label(f, text="", bg=BG_MAIN).grid(row=0, column=0)
        tk.Label(f, text="PFAA 1", font=FONT_LABEL_B,
                 bg=BG_MAIN).grid(row=0, column=1, padx=4)
        tk.Label(f, text="PFAA 2", font=FONT_LABEL_B,
                 bg=BG_MAIN).grid(row=0, column=2, padx=4)
        # Detailed-only headers
        lbl_pre1 = tk.Label(f, text="Precursor 1", font=FONT_LABEL_B, bg=BG_MAIN)
        lbl_pre1.grid(row=0, column=3, padx=4)
        lbl_pre2 = tk.Label(f, text="Precursor 2", font=FONT_LABEL_B, bg=BG_MAIN)
        lbl_pre2.grid(row=0, column=4, padx=4)
        self._detailed_only_frames += [lbl_pre1, lbl_pre2]

        # ── PFAA dropdowns ─────────────────────────────────────────────
        tk.Label(f, text="PFAA (use dropdown menu)",
                 font=FONT_LABEL, bg=BG_MAIN, anchor="e"
                 ).grid(row=1, column=0, sticky="e", padx=(0, 6))
        # Width 14 fits the longest option ("User-Specified") with a
        # comfortable margin so the full caption is visible at rest.
        dropdown(f, self.v_pfaa1, PFAA1_CHOICES, width=14, bg=BG_PULLDOWN
                 ).grid(row=1, column=1, padx=4)
        dropdown(f, self.v_pfaa2, PFAA2_CHOICES, width=14, bg=BG_PULLDOWN
                 ).grid(row=1, column=2, padx=4)
        # Precursor 1 → only "PFAA 1-able" or "None"
        # Precursor 2 → only "PFAA 2-able" or "None"
        # (Per Excel storyboard — precursors transform into their
        # paired PFAA, so they can't take any other species value.)
        # Width bumped 8 → 14 so the full "PFAA 1-able" caption is
        # visible at rest with a little room to spare.
        PRE1_CHOICES = ["PFAA 1-able", "None"]
        PRE2_CHOICES = ["PFAA 2-able", "None"]
        dd3 = dropdown(f, self.v_pfaa3, PRE1_CHOICES, width=14,
                       bg=BG_PULLDOWN)
        dd3.grid(row=1, column=3, padx=4)
        dd4 = dropdown(f, self.v_pfaa4, PRE2_CHOICES, width=14,
                       bg=BG_PULLDOWN)
        dd4.grid(row=1, column=4, padx=4)
        self._detailed_only_frames += [dd3, dd4]
        help_link(f, "OpenTable5_1").grid(row=1, column=6)

        # ── Retardation rows ───────────────────────────────────────────
        ret_rows = [
            ("Retardation Factor for Transmissive Zones",
             self.v_ret_trans1, self.v_ret_trans2,
             self.v_ret_trans3, self.v_ret_trans4, "(-)", "OpenTable5_3"),
            ("Retardation Factor for Low-k Units",
             self.v_ret_lowk1,  self.v_ret_lowk2,
             self.v_ret_lowk3,  self.v_ret_lowk4,  "(-)", "OpenTable5_4"),
        ]
        for i, (lbl, v1, v2, v3, v4, unit, helpm) in enumerate(ret_rows):
            tk.Label(f, text=lbl, font=FONT_LABEL, bg=BG_MAIN, anchor="e"
                     ).grid(row=2+i, column=0, sticky="e", pady=1, padx=(0, 6))
            # Width 14 — matches the §5 dropdowns above so the column
            # of cells reads as a single visual block.
            e1 = make_entry(f, v1, width=14, bg=BG_LOCKED)
            e1.grid(row=2+i, column=1, padx=4)
            e2 = make_entry(f, v2, width=14, bg=BG_LOCKED)
            e2.grid(row=2+i, column=2, padx=4)
            e3 = make_entry(f, v3, width=14, bg=BG_LOCKED)
            e3.grid(row=2+i, column=3, padx=4)
            e4 = make_entry(f, v4, width=14, bg=BG_LOCKED)
            e4.grid(row=2+i, column=4, padx=4)
            # Section 5 (black retardation cells): enforce 1-decimal display.
            for _e, _v in ((e1, v1), (e2, v2), (e3, v3), (e4, v4)):
                self._bind_decimal_format(_e, _v, 1)
            self._detailed_only_frames += [e3, e4]
            tk.Label(f, text=unit, font=FONT_LABEL_SM, bg=BG_MAIN
                     ).grid(row=2+i, column=5, sticky="w")
            help_link(f, helpm).grid(row=2+i, column=6, padx=2)

        # Start in correct state for current version
        if self.active_sheet != "Detailed_2":
            for w in [lbl_pre1, lbl_pre2, dd3, dd4]:
                w.grid_remove()
            # entry widgets for pre cols already added above, remove them too
            for w in self._detailed_only_frames[-4:]:
                w.grid_remove()

        # ── Detailed-only: Transformation Rate + Yield Factor rows ─
        # XLSM J41/K41/M41/N41 (Trans. Rate) and J42/K42/M42/N42 (Yield).
        # White cells with red text in Precursor 1/2 columns.
        for r_off, lbl, v3, v4, unit in [
            (4, "Transformation Rate",
             self.v_trans_rate_3, self.v_trans_rate_4, "(years)"),
            (5, "Yield Factor",
             self.v_yield_factor_3, self.v_yield_factor_4, "(-)"),
        ]:
            tr_lbl = tk.Label(f, text=lbl, font=FONT_LABEL,
                              bg=BG_MAIN, anchor="e")
            tr_lbl.grid(row=r_off, column=0, sticky="e",
                        pady=1, padx=(0, 6))
            # Width 14 — matches the §5 dropdowns + retardation cells
            # above so the precursor column reads as a single block.
            tr_e3 = tk.Entry(f, textvariable=v3, width=14, font=FONT_INPUT,
                             bg="#FFFFFF", fg="#C00000",
                             relief="solid", bd=1, justify="right")
            tr_e3.grid(row=r_off, column=3, padx=4)
            tr_e4 = tk.Entry(f, textvariable=v4, width=14, font=FONT_INPUT,
                             bg="#FFFFFF", fg="#C00000",
                             relief="solid", bd=1, justify="right")
            tr_e4.grid(row=r_off, column=4, padx=4)
            tr_unit = tk.Label(f, text=unit, font=FONT_LABEL_SM,
                               bg=BG_MAIN, fg=FG_GREY)
            tr_unit.grid(row=r_off, column=5, sticky="w")
            self._detailed_only_frames += [tr_lbl, tr_e3, tr_e4, tr_unit]
            # Hide on init if currently Simple
            if self.active_sheet != "Detailed_2":
                for w in (tr_lbl, tr_e3, tr_e4, tr_unit):
                    try: w.grid_remove()
                    except Exception: pass

        # "Calculate Retardation Factors" + "Modeling Transformation Low K"
        # buttons live in the same row.  Use grid layout inside cb so the
        # detailed-only Modeling Transformation button can be hidden via
        # grid_remove (consistent with the existing toggle).
        cb = tk.Frame(parent, bg=BG_MAIN); cb.pack(anchor="center", pady=(3, 3))
        make_btn(cb, "Calculate Retardation Factors",
                 "CalculrateRetardationFactors",
                 font=FONT_BTN_SM, width=24, bg=BTN_FILL
                 ).grid(row=0, column=0, padx=2)
        help_link(cb, "OpenTable5_2").grid(row=0, column=1, padx=4)
        mt_btn = make_btn(cb, "Modeling Transformation Low K",
                          "ModelingTransformationLowK",
                          font=FONT_BTN_SM, width=28, bg=BTN_FILL)
        mt_btn.grid(row=0, column=2, padx=(12, 2))
        self._detailed_only_frames.append(mt_btn)
        if self.active_sheet != "Detailed_2":
            try: mt_btn.grid_remove()
            except Exception: pass

        # General Molecular Diffusion Coefficient on its own row — centered
        df = tk.Frame(parent, bg=BG_MAIN)
        df.pack(anchor="center", pady=(2, 4))
        tk.Label(df, text="General molecular Diffusion Coefficient for all Constituents",
                 font=FONT_LABEL, bg=BG_MAIN, anchor="e"
                 ).grid(row=0, column=0, sticky="e", padx=(0, 6))
        make_entry(df, self.v_mol_diff, width=10, bg=BG_FORMULA).grid(
            row=0, column=1, padx=4)
        tk.Label(df, text="(m²/sec)", font=FONT_LABEL_SM, bg=BG_MAIN
                 ).grid(row=0, column=2, sticky="w")
        help_link(df, "OpenTable5_7").grid(row=0, column=3)

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 7 – PFAS SOURCE TERM (middle column)
    #
    # XLSM-authoritative layout (Detailed_2 sheet):
    #   col Q   – Option 1:/2:/3: labels + button frames
    #   col T   – "Go to Section 2 to Change Years" rotated label
    #             (BOTH versions; spans years 1977…2077)
    #   col U   – Years (1977…2077)
    #   col V   – PFAA 1 values   (header V6, species V7=E38)
    #   col X   – PFAA 2 values   (header X6, species X7=G38)
    #   col Y   – "(ug/L)" units / "(kg)" on totals row
    #   col Z   – Precursors 1 values  (DETAILED ONLY, header Z6, Z7=K38)
    #   col AB  – Precursors 2 values  (DETAILED ONLY, header AB6, AB7=M38)
    #   row 19  – Total PFAS Mass Out of Source per column
    # ─────────────────────────────────────────────────────────────────────
    def _build_s7_source(self, parent):
        section_header(parent, "7", "PFAS SOURCE TERM").pack(
            anchor="w", pady=(2, 1))

        f = tk.Frame(parent, bg=BG_MAIN)
        f.pack(fill="x", pady=(0, 4))

        PFAA1_CHOICES = ["PFOS", "PFOA", "PFHxS", "PFHxA", "PFBS", "PFNA",
                         "User-Specified"]
        PFAA2_CHOICES = ["None"] + PFAA1_CHOICES
        PRE_CHOICES   = ["PFAA 1-able", "PFAA 2-able"] + PFAA1_CHOICES

        # Track Detailed-only widgets locally for reliable hide-on-init.
        det_only = []

        def _mark_detailed(*widgets):
            for w in widgets:
                det_only.append(w)
                self._detailed_only_frames.append(w)

        # ── Row 0 – Column headers (italic gray text).  Cell bg is
        #    version-specific: gray (BG_MAIN) for Simple, white
        #    (BG_INPUT_BLUE) for Detailed — see _apply_s7_version_colors.
        hdr_pfaa1 = tk.Label(f, text="PFAA 1", font=FONT_LABEL_BI,
                             bg=BG_MAIN, fg=FG_GREY)
        hdr_pfaa1.grid(row=0, column=3, padx=4)
        hdr_pfaa2 = tk.Label(f, text="PFAA 2", font=FONT_LABEL_BI,
                             bg=BG_MAIN, fg=FG_GREY)
        hdr_pfaa2.grid(row=0, column=4, padx=4)
        hdr_pre1 = tk.Label(f, text="Precursors 1", font=FONT_LABEL_BI,
                            bg=BG_MAIN, fg=FG_GREY)
        hdr_pre1.grid(row=0, column=6, padx=4)
        hdr_pre2 = tk.Label(f, text="Precursors 2", font=FONT_LABEL_BI,
                            bg=BG_MAIN, fg=FG_GREY)
        hdr_pre2.grid(row=0, column=7, padx=4)
        _mark_detailed(hdr_pre1, hdr_pre2)
        self._s7_hdr_labels = [hdr_pfaa1, hdr_pfaa2, hdr_pre1, hdr_pre2]

        # ── Row 1 – "How to Get Source Concentration?" + dropdowns ────
        make_btn(f, "How to Get Source Concentration?",
                 "OpenAppendix_7_1_Relative", quarto=True,
                 font=FONT_BTN_SM, width=28, bg=BTN_FILL
                 ).grid(row=1, column=0, sticky="w",
                        padx=(0, 8), pady=(0, 4))
        # §7 species cells are display-only (XLSM V7=E38, X7=G38, Z7=K38,
        # AB7=M38).  Use locked black labels bound via textvariable so
        # they auto-update when the §5 dropdowns change.
        tk.Label(f, textvariable=self.v_pfaa1, font=FONT_LABEL_B,
                 bg=BG_LOCKED, fg="#FFFFFF", width=8,
                 relief="solid", bd=1
                 ).grid(row=1, column=3, padx=4, sticky="ew")
        tk.Label(f, textvariable=self.v_pfaa2, font=FONT_LABEL_B,
                 bg=BG_LOCKED, fg="#FFFFFF", width=8,
                 relief="solid", bd=1
                 ).grid(row=1, column=4, padx=4, sticky="ew")
        help_link(f, "OpenTable5_1").grid(row=1, column=5, padx=(2, 0))
        dd_pre1 = tk.Label(f, textvariable=self.v_pfaa3, font=FONT_LABEL_B,
                           bg=BG_LOCKED, fg="#FFFFFF", width=8,
                           relief="solid", bd=1)
        dd_pre1.grid(row=1, column=6, padx=4, sticky="ew")
        dd_pre2 = tk.Label(f, textvariable=self.v_pfaa4, font=FONT_LABEL_B,
                           bg=BG_LOCKED, fg="#FFFFFF", width=8,
                           relief="solid", bd=1)
        dd_pre2.grid(row=1, column=7, padx=4, sticky="ew")
        _mark_detailed(dd_pre1, dd_pre2)

        # ── Col 1 rowspan – rotated "Go to Section 2 to Change Years" ─
        # XLSM cell T8 — present in BOTH Simple and Detailed sheets.
        # Bigger font + wider canvas for legibility; spans only the 11
        # year rows (rows 2..12), leaving the totals row 13 untouched.
        rot_h = 28 * 11
        rot = tk.Canvas(f, width=26, height=rot_h,
                        bg=BG_MAIN, highlightthickness=0)
        rot.create_text(13, rot_h // 2,
                        text="Go to Section 2 to Change Years",
                        font=("Calibri", 11, "italic"),
                        fill=FG_BTN_NAVY, angle=90)
        rot.grid(row=2, column=1, rowspan=11, sticky="ns",
                 padx=(0, 6), pady=(0, 0))

        # ── Col 0 – Option 1/2/3 stacked frames ───────────────────────
        # Option 1: rowspan=4 covers years 1977…2007 (i 0..3)
        opt1 = tk.Frame(f, bg=BG_MAIN)
        tk.Label(opt1, text="Option 1:", font=FONT_LABEL_I, bg=BG_MAIN
                 ).pack(anchor="w")
        make_btn(opt1, "Assume Constant Source\n(populates same values\n"
                       "from the first row)",
                 "SourceOption1", font=FONT_BTN_SM, width=24, bg=BTN_FILL,
                 padx=4, pady=2
                 ).pack(anchor="w", pady=(2, 0))
        opt1.grid(row=2, column=0, rowspan=4, padx=(0, 8), sticky="n")

        # Option 2: rowspan=5 covers years 2017…2057 (i 4..8) — DETAILED
        opt2 = tk.Frame(f, bg=BG_MAIN)
        tk.Label(opt2, text="Option 2:", font=FONT_LABEL_I, bg=BG_MAIN
                 ).pack(anchor="w")
        make_btn(opt2, "Upload PFAS\nConcentrations\nfrom PFAS-LEACH",
                 "SourceOption2", font=FONT_BTN_SM, width=24,
                 bg=BTN_FILL, padx=4, pady=2
                 ).pack(anchor="w", pady=(2, 0))
        opt2.grid(row=6, column=0, rowspan=5, padx=(0, 8), sticky="n")
        _mark_detailed(opt2)

        # Option 3: rowspan=2 covers years 2067…2077 (i 9..10) — DETAILED
        opt3 = tk.Frame(f, bg=BG_MAIN)
        tk.Label(opt3, text="Option 3: Enter\nconcs. for each year.",
                 font=FONT_LABEL_I, bg=BG_MAIN, justify="left"
                 ).pack(anchor="w")
        opt3.grid(row=11, column=0, rowspan=2, padx=(0, 8), sticky="n")
        _mark_detailed(opt3)

        # ── Year × value rows (i = 0..10  →  grid rows 2..12) ──────────
        # Cell colors:
        #   * Year column            → SKY BLUE in BOTH versions
        #   * PFAA 1/2 conc columns  → version-specific (Simple WHITE,
        #                               Detailed SKY BLUE)
        #   * Precursors 1/2 columns → SKY BLUE (Detailed-only)
        # Wider cells (10 → 14) so 5-6 digit values like "1,600.00"
        # don't get truncated.
        self._s7_conc12_cells = []
        for i in range(11):
            r = 2 + i
            make_entry(f, self.v_src_years[i], width=8,
                       bg=BG_FORMULA, justify="right"
                       ).grid(row=r, column=2, padx=2, pady=1)
            e_p1 = make_entry(f, self.v_src_pfaa1[i], width=14,
                              bg=BG_INPUT_BLUE)
            e_p1.grid(row=r, column=3, padx=2, pady=1)
            e_p2 = make_entry(f, self.v_src_pfaa2[i], width=14,
                              bg=BG_INPUT_BLUE)
            e_p2.grid(row=r, column=4, padx=2, pady=1)
            self._s7_conc12_cells.extend([e_p1, e_p2])
            tk.Label(f, text="(ug/L)", font=FONT_LABEL_SMI,
                     bg=BG_MAIN, fg=FG_GREY
                     ).grid(row=r, column=5, sticky="w", padx=2)
            e_pre1 = make_entry(f, self.v_src_pre1[i], width=14,
                                bg=BG_FORMULA)
            e_pre1.grid(row=r, column=6, padx=2, pady=1)
            e_pre2 = make_entry(f, self.v_src_pre2[i], width=14,
                                bg=BG_FORMULA)
            e_pre2.grid(row=r, column=7, padx=2, pady=1)
            _mark_detailed(e_pre1, e_pre2)

        # ── Row 13 – "Total PFAS Mass Out of Source:" totals ───────────
        tk.Label(f, text="Total PFAS Mass Out of Source:",
                 font=FONT_LABEL, bg=BG_MAIN, fg=FG_BTN_NAVY,
                 anchor="e"
                 ).grid(row=13, column=0, columnspan=3, sticky="e",
                        padx=(0, 4), pady=(6, 0))
        # Width 14 — matches the source-concentration cells above
        # so the totals row visually aligns with each column.
        tk.Label(f, textvariable=self.v_total_mass, font=FONT_LABEL,
                 bg=BG_LOCKED, fg=FG_LOCKED, width=14,
                 relief="solid", bd=1
                 ).grid(row=13, column=3, padx=2, pady=(6, 0))
        tk.Label(f, textvariable=self.v_total_mass_p2, font=FONT_LABEL,
                 bg=BG_LOCKED, fg=FG_LOCKED, width=14,
                 relief="solid", bd=1
                 ).grid(row=13, column=4, padx=2, pady=(6, 0))
        tk.Label(f, text="(kg)", font=FONT_LABEL_SM,
                 bg=BG_MAIN, fg=FG_GREY
                 ).grid(row=13, column=5, sticky="w", padx=2,
                        pady=(6, 0))
        tot_pre1 = tk.Label(f, textvariable=self.v_total_mass_pre1,
                            font=FONT_LABEL, bg=BG_LOCKED, fg=FG_LOCKED,
                            width=14, relief="solid", bd=1)
        tot_pre1.grid(row=13, column=6, padx=2, pady=(6, 0))
        tot_pre2 = tk.Label(f, textvariable=self.v_total_mass_pre2,
                            font=FONT_LABEL, bg=BG_LOCKED, fg=FG_LOCKED,
                            width=14, relief="solid", bd=1)
        tot_pre2.grid(row=13, column=7, padx=2, pady=(6, 0))
        _mark_detailed(tot_pre1, tot_pre2)

        # ── Auto-compute "Total PFAS Mass Out of Source:" per column ──
        # XLSM V19 formula (Detailed):
        #   Total_kg = 1e-6 * Vd * thickness * width * Σ(C_i * Δt_i)
        # where:
        #   Vd        = §3 Bulk Groundwater Darcy Velocity (v_darcy, m/yr)
        #   thickness = §2 Thickness of Source Below WT  (v_sw_thick, m)
        #   width     = §2 Source Width                  (v_sw_width, m)
        #   Δt_i      = year[i+1] - year[i]  for i=0..9
        #               end_year - year[10]  for i=10  (end_year = §2 v_yr_end)
        #   C_i       = column-specific concentration   (ug/L)
        # Verified against XLSM Detailed_2 sample: 1e-6 * 0.011992 * 5 * 60
        # * 112000 = 0.4029 kg, matches V19.
        def _to_float_pt(var, default=None):
            try:
                s = (var.get() or "").replace(",", "").strip()
                return float(s) if s else default
            except (ValueError, AttributeError):
                return default

        def _fmt_kg(v):
            if v is None:
                return "#VALUE!"
            if v == 0:
                return "0.0"
            return f"{v:.4f}"

        def _recompute_total_mass(*_):
            vd    = _to_float_pt(self.v_darcy)
            thick = _to_float_pt(self.v_sw_thick)
            width = _to_float_pt(self.v_sw_width)
            yr_e  = _to_float_pt(self.v_yr_end)
            years = [_to_float_pt(v) for v in self.v_src_years]
            if (vd is None or thick is None or width is None or yr_e is None
                    or any(y is None for y in years)):
                # Missing geometry/year input → mirror Excel's #VALUE! result
                self.v_total_mass.set("#VALUE!")
                self.v_total_mass_p2.set("#VALUE!")
                self.v_total_mass_pre1.set("#VALUE!")
                self.v_total_mass_pre2.set("#VALUE!")
                return
            # Δt: 10 forward differences + (end_year - last_year) for index 10
            dt = [years[i+1] - years[i] for i in range(10)]
            dt.append(yr_e - years[10])
            base = 1e-6 * vd * thick * width

            def _col_total(varlist):
                tot = 0.0
                for i, v in enumerate(varlist):
                    c = _to_float_pt(v)
                    if c is None:
                        return None
                    tot += c * dt[i]
                return base * tot

            self.v_total_mass.set(_fmt_kg(_col_total(self.v_src_pfaa1)))
            self.v_total_mass_p2.set(_fmt_kg(_col_total(self.v_src_pfaa2)))
            self.v_total_mass_pre1.set(_fmt_kg(_col_total(self.v_src_pre1)))
            self.v_total_mass_pre2.set(_fmt_kg(_col_total(self.v_src_pre2)))

        # Wire the trace on every input the formula reads.
        _trace_vars = [self.v_darcy, self.v_sw_thick, self.v_sw_width,
                       self.v_yr_end]
        _trace_vars += list(self.v_src_years)
        _trace_vars += list(self.v_src_pfaa1) + list(self.v_src_pfaa2)
        _trace_vars += list(self.v_src_pre1)  + list(self.v_src_pre2)
        for _v in _trace_vars:
            _v.trace_add("write", _recompute_total_mass)
        self._recompute_total_mass = _recompute_total_mass
        _recompute_total_mass()  # initial value

        # ── §8 % ↔ §7 auto-apply ─────────────────────────────────────
        # When the user types in §8 "Source Concentration Reduction"
        # or "Source Treatment Start Year", recompute §7 PFAA-1 /
        # PFAA-2 (and precursors in Detailed) from baseline ×
        # (1 - %/100) for rows >= start year.
        col_lists = {
            "pfaa1": self.v_src_pfaa1,
            "pfaa2": self.v_src_pfaa2,
            "pre1":  self.v_src_pre1,
            "pre2":  self.v_src_pre2,
        }

        def _parse_num(s):
            try:
                return float(str(s).replace(",", "").strip())
            except (ValueError, TypeError, AttributeError):
                return None

        def _seed_baseline_from_current():
            """Snapshot every §7 cell's current value into baseline,
            regardless of whether it's None or set.  Called once on
            first §8 trace fire AND from restore_from_example /
            restore_from_saved after they push fresh data."""
            for key, lst in col_lists.items():
                for i, var in enumerate(lst):
                    self._s7_baseline[key][i] = _parse_num(var.get())
        self._refresh_s7_baseline = _seed_baseline_from_current

        def _apply_s8_to_s7(*_):
            red = _parse_num(self.v_src_conc_red.get())
            if red is None:
                return                  # cell empty → don't touch §7
            factor = 1.0 - (red / 100.0)
            if factor < 0: factor = 0.0
            elif factor > 1: factor = 1.0

            # Determine start_idx — closest year in §7 to §8 start year.
            start_year = _parse_num(self.v_src_rem_yr.get())
            years = []
            for i, v in enumerate(self.v_src_years):
                y = _parse_num(v.get())
                if y is not None:
                    years.append((i, y))
            if not years:
                return
            if start_year is None:
                start_idx = 0           # blank start year → from top
            else:
                start_idx = min(years, key=lambda t: abs(t[1] - start_year))[0]

            # Lazy-seed baselines: if ALL entries for a column are
            # None, the baselines haven't been captured yet — capture
            # current values now (this happens on first %change after
            # initial paint).
            if all(b is None for b in self._s7_baseline["pfaa1"]):
                _seed_baseline_from_current()

            self._s8_applying = True
            try:
                for key, lst in col_lists.items():
                    base = self._s7_baseline[key]
                    for i, var in enumerate(lst):
                        b = base[i]
                        if b is None:
                            continue
                        if i < start_idx:
                            new = b
                        else:
                            new = b * factor
                        var.set(f"{new:,.2f}")
            finally:
                self._s8_applying = False
        self._apply_s8_to_s7 = _apply_s8_to_s7

        # When user MANUALLY edits a §7 cell, refresh that row's
        # baseline so subsequent §8 changes operate from the new value.
        def _make_baseline_updater(key, idx):
            def _updater(*_):
                if self._s8_applying:
                    return                # programmatic write — ignore
                self._s7_baseline[key][idx] = _parse_num(
                    col_lists[key][idx].get())
            return _updater
        for key, lst in col_lists.items():
            for i, var in enumerate(lst):
                var.trace_add("write", _make_baseline_updater(key, i))

        # v101: NO auto-apply on §8 changes — user explicitly asked
        # that §7 only updates when the "Apply Remediation" button is
        # pressed (avoids unexpected mass loss while editing % or
        # start year).  The SourceRemediation dispatch in run_script()
        # now calls _apply_s8_to_s7 on demand.
        #
        # Previously:
        #   self.v_src_conc_red.trace_add("write", _apply_s8_to_s7)
        #   self.v_src_rem_yr  .trace_add("write", _apply_s8_to_s7)

        # Apply the Simple/Detailed-specific colors to the §7 widgets
        # whose bg differs between versions (Row 0 header labels +
        # PFAA 1/2 conc cells).  Re-runs from _on_model_version_change.
        self._apply_s7_version_colors()

        # In Simple mode, hide every Detailed-only widget registered above.
        if self.active_sheet != "Detailed_2":
            for w in det_only:
                try:
                    w.grid_remove()
                except Exception:
                    pass

    # ─────────────────────────────────────────────────────────────────────
    # §7 cell-color helper — keeps Simple/Detailed visual differences
    # ─────────────────────────────────────────────────────────────────────
    def _apply_s7_version_colors(self):
        """Set §7 cells whose color depends on Simple vs Detailed.

        Simple   → Row 0 hdr bg = BG_MAIN  ; PFAA 1/2 conc bg = WHITE
        Detailed → Row 0 hdr bg = WHITE    ; PFAA 1/2 conc bg = SKY BLUE
        """
        is_det = (getattr(self, "active_sheet", "Simple") == "Detailed_2")
        hdr_bg    = BG_INPUT_BLUE if is_det else BG_MAIN
        conc12_bg = BG_FORMULA    if is_det else BG_INPUT_BLUE
        for w in getattr(self, "_s7_hdr_labels", []):
            try:
                w.configure(bg=hdr_bg)
            except Exception:
                pass
        for w in getattr(self, "_s7_conc12_cells", []):
            try:
                w.configure(bg=conc12_bg, readonlybackground=conc12_bg,
                            disabledbackground=conc12_bg)
            except Exception:
                pass

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 8 – SOURCE REMEDIATION
    # ─────────────────────────────────────────────────────────────────────
    def _build_s8_source_rem(self, parent):
        section_header(parent, "8", "SOURCE REMEDIATION").pack(
            anchor="w", pady=(2, 1))

        f = tk.Frame(parent, bg=BG_MAIN)
        f.pack(fill="x", pady=(0, 4))

        tk.Label(f, text="Source Treatment Start Year",
                 font=FONT_LABEL, bg=BG_MAIN).grid(row=0, column=0, sticky="w")
        make_entry(f, self.v_src_rem_yr, width=8).grid(row=0, column=1, padx=4)
        tk.Label(f, text="(YYYY)", font=FONT_LABEL_SM, bg=BG_MAIN
                 ).grid(row=0, column=2, sticky="w")
        help_link(f, "OpenTable8_2").grid(row=0, column=3)

        tk.Label(f, text="Source Concentration Reduction",
                 font=FONT_LABEL, bg=BG_MAIN).grid(row=1, column=0, sticky="w", pady=2)
        make_entry(f, self.v_src_conc_red, width=8).grid(row=1, column=1, padx=4)
        tk.Label(f, text="(%)", font=FONT_LABEL_SM, bg=BG_MAIN
                 ).grid(row=1, column=2, sticky="w")
        help_link(f, "OpenTable8_1").grid(row=1, column=3)

        bf = tk.Frame(parent, bg=BG_MAIN)
        bf.pack(fill="x", pady=(2, 4))
        make_btn(bf, "Apply Remediation", "SourceRemediation",
                 font=FONT_BTN_SM, width=18, bg=BTN_FILL).pack(side="left", padx=(0, 4))
        make_btn(bf, "HELP", "OpenAppendix_8_1_Relative", quarto=True,
                 font=FONT_BTN_SM, width=8, bg=BTN_FILL).pack(side="left")

        # ── Section8_1.png below the buttons ─────────────────────────
        scale = getattr(self, "_dpi_scale", 1.0)
        self._figures["s8_1"] = _load_figure("Section8_1.png",
                                             target_height=int(120 * scale))
        if self._figures["s8_1"] is not None:
            tk.Label(parent, image=self._figures["s8_1"],
                     bg=BG_MAIN, bd=0).pack(anchor="w", pady=(2, 4))

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 9 – PSB
    # ─────────────────────────────────────────────────────────────────────
    def _build_s9_psb(self, parent):
        _hdr9 = tk.Frame(parent, bg=BG_MAIN)
        _hdr9.pack(anchor="w", pady=(2, 1), fill="x")
        section_header(_hdr9, "9",
                       "PLUME REMEDIATION: INSTALL PERMEABLE SORPTION BARRIER (PSB)"
                       ).pack(side="left")
        help_link(_hdr9, "OpenTable9_0").pack(side="left", padx=(6, 0))

        det_only = []
        def _mark_detailed(*ws):
            for w in ws:
                w._toggle_kind = "grid"
                det_only.append(w)
                self._detailed_only_frames.append(w)

        # ═══ TOP — leftcol(form + buttons) | sep | PSB Dist | rcluster ═══
        top = tk.Frame(parent, bg=BG_MAIN)
        top.pack(fill="x", pady=(0, 2))

        # ── LEFT column: Freundlich grid + the two action buttons
        #    stacked underneath it (so the buttons sit flush against the
        #    bottom of the form, no empty gap between the form cell and
        #    the button row).
        leftcol = tk.Frame(top, bg=BG_MAIN)
        leftcol.pack(side="left", anchor="nw")
        f = tk.Frame(leftcol, bg=BG_MAIN)
        f.pack(anchor="nw")

        # Headers row 0
        big_check(f, "Model PSB?", self.v_model_psb,
                  bg=BG_MAIN
                  ).grid(row=0, column=0, sticky="w", padx=(0, 4))
        tk.Label(f, text="Unit", font=FONT_LABEL_B, bg=BG_PULLDOWN,
                 relief="solid", bd=1, padx=4
                 ).grid(row=0, column=1, padx=2, sticky="ew")
        tk.Label(f, textvariable=self.v_pfaa1, font=FONT_LABEL_B,
                 bg=BG_LOCKED, fg="#FFFFFF",
                 relief="solid", bd=1, width=12
                 ).grid(row=0, column=2, padx=2, sticky="ew")
        tk.Label(f, textvariable=self.v_pfaa2, font=FONT_LABEL_B,
                 bg=BG_LOCKED, fg="#FFFFFF",
                 relief="solid", bd=1, width=12
                 ).grid(row=0, column=3, padx=2, sticky="ew")
        h_pre1 = tk.Label(f, textvariable=self.v_pfaa3,
                          font=FONT_LABEL_B, bg=BG_LOCKED, fg="#FFFFFF",
                          relief="solid", bd=1, width=12)
        h_pre1.grid(row=0, column=4, padx=2, sticky="ew")
        h_pre2 = tk.Label(f, textvariable=self.v_pfaa4,
                          font=FONT_LABEL_B, bg=BG_LOCKED, fg="#FFFFFF",
                          relief="solid", bd=1, width=12)
        h_pre2.grid(row=0, column=5, padx=2, sticky="ew")
        _mark_detailed(h_pre1, h_pre2)

        # Row 1: Freundlich "a" — Unit cell GREY
        tk.Label(f, text='PSB\'s Freundlich "a"', font=FONT_LABEL,
                 bg=BG_MAIN, anchor="e"
                 ).grid(row=1, column=0, sticky="e", pady=1, padx=(0, 4))
        tk.Label(f, text="-", font=FONT_LABEL_SM, bg=BTN_FILL,
                 fg=FG_INPUT, relief="solid", bd=1, padx=2
                 ).grid(row=1, column=1, padx=2, sticky="ew")
        # Sky-blue (BG_FORMULA) data cells per Excel storyboard — these
        # rows are "formula or default, but ok to overwrite" cells.
        for c, var in enumerate([self.v_psb_a_1, self.v_psb_a_2]):
            make_entry(f, var, width=12, bg=BG_FORMULA
                       ).grid(row=1, column=2 + c, padx=2)
        e_a3 = make_entry(f, self.v_psb_a_3, width=12, bg=BG_FORMULA)
        e_a3.grid(row=1, column=4, padx=2)
        e_a4 = make_entry(f, self.v_psb_a_4, width=12, bg=BG_FORMULA)
        e_a4.grid(row=1, column=5, padx=2)
        _mark_detailed(e_a3, e_a4)
        help_link(f, "OpenTable9_1").grid(row=1, column=6, padx=(2, 0))

        # Row 2: Freundlich Kf — Unit cell DROPDOWN (6 options from XLSM)
        tk.Label(f, text="PSB's Freundlich Kf", font=FONT_LABEL,
                 bg=BG_MAIN, anchor="e"
                 ).grid(row=2, column=0, sticky="e", pady=1, padx=(0, 4))
        KF_UNITS = [
            "(ng/kg)(ng/L)^(-a)",
            "(ug/kg)(ug/L)^(-a)",
            "(mg/kg)(mg/L)^(-a)",
            "(nmol/kg)(nmol/L)^(-a)",
            "(umol/kg)(umol/L)^(-a)",
            "(mmol/kg)(mmol/L)^(-a)",
        ]
        dropdown(f, self.v_psb_kf_unit, KF_UNITS,
                 width=20, bg=BG_PULLDOWN
                 ).grid(row=2, column=1, padx=2, sticky="ew")
        # Sky-blue (BG_FORMULA) data cells per Excel storyboard.
        for c, var in enumerate([self.v_psb_kf_1, self.v_psb_kf_2]):
            make_entry(f, var, width=12, bg=BG_FORMULA
                       ).grid(row=2, column=2 + c, padx=2)
        e_kf3 = make_entry(f, self.v_psb_kf_3, width=12, bg=BG_FORMULA)
        e_kf3.grid(row=2, column=4, padx=2)
        e_kf4 = make_entry(f, self.v_psb_kf_4, width=12, bg=BG_FORMULA)
        e_kf4.grid(row=2, column=5, padx=2)
        _mark_detailed(e_kf3, e_kf4)
        help_link(f, "OpenTable9_2").grid(row=2, column=6, padx=(2, 0))

        # Row 3: PFAS molecular weight (mol/g) — only visible when the
        # Kf unit selected on row 2 contains "mol".  Layout matches the
        # XLSM (U25 = "PFAS molecular weight (mol/g)" inside a BLACK
        # locked cell; V25/X25/Z25/AB25 are SKY-BLUE EDITABLE input
        # cells the user types the molecular weight into).
        mw_unit = tk.Label(f, text="PFAS molecular weight (mol/g)",
                           font=FONT_LABEL_SM,
                           bg=BG_LOCKED, fg="#FFFFFF",
                           relief="solid", bd=1, padx=2)
        mw_unit.grid(row=3, column=1, padx=2, sticky="ew")
        # Editable sky-blue (BG_FORMULA) cells — user types MW in g/mol.
        mw_e1 = make_entry(f, self.v_psb_mw_1, width=12, bg=BG_FORMULA)
        mw_e1.grid(row=3, column=2, padx=2)
        mw_e2 = make_entry(f, self.v_psb_mw_2, width=12, bg=BG_FORMULA)
        mw_e2.grid(row=3, column=3, padx=2)
        mw_e3 = make_entry(f, self.v_psb_mw_3, width=12, bg=BG_FORMULA)
        mw_e3.grid(row=3, column=4, padx=2)
        mw_e4 = make_entry(f, self.v_psb_mw_4, width=12, bg=BG_FORMULA)
        mw_e4.grid(row=3, column=5, padx=2)
        _mark_detailed(mw_e3, mw_e4)
        # Track which widgets are mol-only and which are also detailed-only.
        # No left-side label any more — column 0 stays empty for this row,
        # mirroring the XLSM where row 25 has no T25 caption.
        self._psb_mw_widgets     = [mw_unit, mw_e1, mw_e2]
        self._psb_mw_widgets_det = [mw_e3, mw_e4]

        # Row 4: Converted Kf — BLACK Unit + BLACK locked values
        tk.Label(f, text="Converted PSB's Freundlich Kf",
                 font=FONT_LABEL, bg=BG_MAIN, anchor="e"
                 ).grid(row=4, column=0, sticky="e", pady=1, padx=(0, 4))
        tk.Label(f, text="(ug/kg)(ug/L)^(-a)", font=FONT_LABEL_SM,
                 bg=BG_LOCKED, fg="#FFFFFF",
                 relief="solid", bd=1, padx=2
                 ).grid(row=4, column=1, padx=2, sticky="ew")
        for c, var in enumerate([self.v_psb_kf_conv, self.v_psb_kf_conv2]):
            make_entry(f, var, width=12, bg=BG_LOCKED
                       ).grid(row=4, column=2 + c, padx=2)
        e_cv3 = make_entry(f, self.v_psb_kf_conv3, width=12, bg=BG_LOCKED)
        e_cv3.grid(row=4, column=4, padx=2)
        e_cv4 = make_entry(f, self.v_psb_kf_conv4, width=12, bg=BG_LOCKED)
        e_cv4.grid(row=4, column=5, padx=2)
        _mark_detailed(e_cv3, e_cv4)

        # MW row visibility — driven by current Kf unit selection.
        def _on_kf_unit_change(*_):
            unit   = (self.v_psb_kf_unit.get() or "").lower()
            is_mol = "mol" in unit
            is_det = (self.active_sheet == "Detailed_2")
            for w in self._psb_mw_widgets:
                try:
                    w.grid() if is_mol else w.grid_remove()
                except Exception:
                    pass
            for w in self._psb_mw_widgets_det:
                try:
                    w.grid() if (is_mol and is_det) else w.grid_remove()
                except Exception:
                    pass
        self.v_psb_kf_unit.trace_add("write", _on_kf_unit_change)
        self._on_psb_kf_unit_change = _on_kf_unit_change   # exposed so the
        #  model-version toggle handler can re-run it.
        _on_kf_unit_change()  # initial state — hidden by default (mg-based)

        # ── Auto-compute Converted PSB's Freundlich Kf ────────────
        # Conversion to (ug/kg)(ug/L)^(-a):
        #     Kf_ug = Kf × m^(1-a)
        # where m is the unit-prefix factor (mass1 → ug):
        #     ng/kg     m = 1e-3
        #     ug/kg     m = 1.0
        #     mg/kg     m = 1e3
        #     nmol/kg   m = MW × 1e-3
        #     umol/kg   m = MW
        #     mmol/kg   m = MW × 1e3
        # `a` and `MW` come from row 1 / row 3 of the same column.
        def _to_float(var, default=None):
            # v106: accept BOTH US ("1,234.56" / "1,227,951") and European
            # ("1.234,56" / "0,7") number formats so the converted-Kf cell
            # fills regardless of how the user types the value.
            try:
                s = (var.get() or "").strip()
                if not s:
                    return default
                if "." in s and "," in s:
                    # Both separators present → the LAST one is the decimal.
                    if s.rfind(",") > s.rfind("."):     # European 1.234,56
                        s = s.replace(".", "").replace(",", ".")
                    else:                                # US 1,234.56
                        s = s.replace(",", "")
                elif "," in s:
                    # Only commas: grouped thousands (1,600 / 1,227,951) vs
                    # a European decimal (0,7).
                    if re.fullmatch(r"\d{1,3}(,\d{3})+", s):
                        s = s.replace(",", "")           # US thousands
                    else:
                        s = s.replace(",", ".")          # European decimal
                return float(s)
            except (ValueError, AttributeError):
                return default

        def _unit_prefix_factor(unit_text, mw):
            u = (unit_text or "").lower()
            if "ng/kg"   in u: return 1e-3
            if "ug/kg"   in u: return 1.0
            if "mg/kg"   in u: return 1e3
            # Mol-based units need a positive MW (g/mol) to be defined
            if mw is None or mw <= 0:
                return None
            if "nmol/kg" in u: return mw * 1e-3
            if "umol/kg" in u: return mw
            if "mmol/kg" in u: return mw * 1e3
            return None

        def _recompute_psb_conv_kf(*_):
            unit = self.v_psb_kf_unit.get()
            for var_a, var_kf, var_mw, var_out in [
                (self.v_psb_a_1, self.v_psb_kf_1, self.v_psb_mw_1, self.v_psb_kf_conv),
                (self.v_psb_a_2, self.v_psb_kf_2, self.v_psb_mw_2, self.v_psb_kf_conv2),
                (self.v_psb_a_3, self.v_psb_kf_3, self.v_psb_mw_3, self.v_psb_kf_conv3),
                (self.v_psb_a_4, self.v_psb_kf_4, self.v_psb_mw_4, self.v_psb_kf_conv4),
            ]:
                a   = _to_float(var_a)
                kf  = _to_float(var_kf)
                mw  = _to_float(var_mw)
                if kf is None:
                    var_out.set("")
                    continue
                m = _unit_prefix_factor(unit, mw)
                if m is None:
                    var_out.set("")
                    continue
                # v106: the exponent `a` only changes the result when the
                # unit prefix m != 1 (i.e. non-ug units).  For the default
                # (ug/kg)(ug/L)^(-a) unit, m == 1 so the converted Kf equals
                # the entered Kf — fill it as soon as Kf is present, even if
                # the user hasn't typed `a` yet.  Other units still need `a`.
                if a is None:
                    if abs(m - 1.0) < 1e-12:
                        conv = kf
                    else:
                        var_out.set("")
                        continue
                else:
                    try:
                        conv = kf * (m ** (1 - a))
                    except (ValueError, OverflowError, ZeroDivisionError):
                        var_out.set("")
                        continue
                # Format with thousands separators — never scientific.
                # Whole numbers display without decimals ("1,227,951"),
                # otherwise up to 4 decimals are kept and trailing zeros
                # stripped ("0.33", "0.5617", "1,234.56").
                if conv == 0:
                    var_out.set("0")
                elif abs(conv - round(conv)) < 1e-9:
                    var_out.set(f"{int(round(conv)):,}")
                else:
                    var_out.set(f"{conv:,.4f}".rstrip("0").rstrip("."))

        # Wire traces — any of these StringVar changes triggers a recompute.
        for _v in (self.v_psb_kf_unit,
                   self.v_psb_a_1, self.v_psb_a_2, self.v_psb_a_3, self.v_psb_a_4,
                   self.v_psb_kf_1, self.v_psb_kf_2, self.v_psb_kf_3, self.v_psb_kf_4,
                   self.v_psb_mw_1, self.v_psb_mw_2, self.v_psb_mw_3, self.v_psb_mw_4):
            _v.trace_add("write", _recompute_psb_conv_kf)
        self._recompute_psb_conv_kf = _recompute_psb_conv_kf
        _recompute_psb_conv_kf()  # initial value

        # ── Vertical separator ────────────────────────────────────
        sep = tk.Frame(top, bg="#000000", width=1)
        sep.pack(side="left", fill="y", padx=(8, 8))

        # ── PSB Dist From Source — multi-line BLACK ───────────────
        psbd = tk.Frame(top, bg=BG_MAIN)
        psbd.pack(side="left", anchor="n", padx=(0, 8))
        for w in ("PSB", "Dist.", "From", "Source"):
            tk.Label(psbd, text=w, font=FONT_LABEL, bg=BG_MAIN
                     ).pack(anchor="w")
        dist_row = tk.Frame(psbd, bg=BG_MAIN)
        dist_row.pack(anchor="w", pady=(4, 0))
        make_entry(dist_row, self.v_psb_dist, width=6
                   ).pack(side="left")
        # v102: PSB distance is a length → label flips m↔ft and value converts
        _psbd_lbl = tk.Label(dist_row, text=f"({self._unit_len()})",
                             font=FONT_LABEL_SM, bg=BG_MAIN)
        _psbd_lbl.pack(side="left", padx=(2, 0))
        self._register_unit_label(_psbd_lbl, "({u})")
        self._register_length_var(self.v_psb_dist, "length")
        help_link(dist_row, "OpenTable9_3").pack(side="left", padx=(2, 0))

        # ── RIGHT cluster: image + Total Width stack + Year PSB,
        # vertically grouped & centered ───────────────────────────
        rcluster = tk.Frame(top, bg=BG_MAIN)
        rcluster.pack(side="left", anchor="n", padx=(4, 0))

        # Image on a Canvas so we can overlay a curved gray connector
        # line that links the "PSB Dist From Source" help "?" chicklet
        # (the last widget in `psbd.dist_row`, just outside this
        # canvas's bottom-left corner) to the red SOURCE square on the
        # 3D-box figure.  The canvas is intentionally taller than the
        # image — the extra strip beneath the image gives the curve
        # vertical room to drop down to "?"-button height.
        scale = getattr(self, "_dpi_scale", 1.0)
        self._figures["s9_1"] = _load_figure("Section9_1.png",
                                             target_height=int(70 * scale))
        img = self._figures["s9_1"]
        if img is not None:
            iw = img.width()
            ih = img.height()
            # Tighter below-image strip: ~14 px (was 34 px).  Reduces the
            # vertical gap between the figure and the Year PSB / Total
            # Width / fcac / # cells stack so §10 can move up.  The
            # curve still droops below the image enough to read as
            # pointing at the "?" chicklet sitting in psbd.dist_row.
            extra = int(14 * scale)
            ch = ih + extra                     # canvas height
            img_canvas = tk.Canvas(rcluster, width=iw, height=ch,
                                   bg=BG_MAIN, highlightthickness=0, bd=0)
            img_canvas.pack(anchor="nw", pady=(0, 0))
            img_canvas.create_image(0, 0, anchor="nw", image=img)
            # Source rectangle on the original 1165×365 png is at roughly
            # x≈55, y≈140 (left red square).  Scale into the resized
            # image and use that as the curve's terminus.
            sx = max(2, int(iw * 0.06))
            sy = int(ih * 0.42)
            # Curve start = bottom-left of canvas, which now sits at the
            # same vertical level as the "?" help chicklet inside `psbd`.
            img_canvas.create_line(
                0,                  ch - 2,           # start at "?" level
                int(iw * 0.04),     ih + int(extra * 0.4),
                int(iw * 0.06),     int(ih * 0.80),
                sx,                 sy,               # land on source
                smooth=True,
                fill="#7F7F7F", width=2, capstyle="round",
            )

        # Year PSB Installed — packed ABOVE the Total Width stack and
        # left-anchored so it appears slightly to the left of the
        # "Total Width of PSB in X-Direction" label that sits below.
        # pady tightened from (0, 4) → (0, 0) so the cell hugs the
        # bottom of the image canvas above it.
        yr = tk.Frame(rcluster, bg=BG_MAIN)
        yr.pack(anchor="w", pady=(0, 0))
        tk.Label(yr, text="Year PSB Installed", font=FONT_LABEL,
                 bg=BG_MAIN
                 ).pack(side="left")
        make_entry(yr, self.v_psb_yr, width=8
                   ).pack(side="left", padx=(4, 2))
        help_link(yr, "OpenTable9_3").pack(side="left", padx=(2, 0))

        # Total Width / fcac / # cells stack (right-anchored so the
        # right-edge of these long labels stays aligned with the
        # rcluster column; YR above it visually pokes out to the left).
        rg = tk.Frame(rcluster, bg=BG_MAIN)
        rg.pack(anchor="e", pady=(0, 0))
        # v102: PSB width is a length → flips m↔ft; load + cells stay literal.
        for i, (lbl, var, unit, helpm) in enumerate([
            ("Total Width of PSB in X-Direction",
             self.v_psb_width, "({u}) (Typical 4)",  "OpenTable9_4"),
            ("PSB Loading 'fcac'",
             self.v_psb_load,  "(%)",                "OpenTable9_5"),
            ("# of cells in PSB in x direction",
             self.v_psb_cells, "(-)",                ""),
        ]):
            tk.Label(rg, text=lbl, font=FONT_LABEL, bg=BG_MAIN,
                     anchor="e"
                     ).grid(row=i, column=0, sticky="e", padx=(0, 4),
                            pady=1)
            make_entry(rg, var, width=6
                       ).grid(row=i, column=1, padx=2)
            _unit_text = unit.format(u=self._unit_len()) if "{u}" in unit else unit
            _u_lbl = tk.Label(rg, text=_unit_text,
                              font=FONT_LABEL_SM, bg=BG_MAIN)
            _u_lbl.grid(row=i, column=2, sticky="w", padx=(2, 0))
            if "{u}" in unit:
                self._register_unit_label(_u_lbl, unit)
                # only convert the width value (load % + count don't convert)
                if var is self.v_psb_width:
                    self._register_length_var(var, "length")
            if helpm:
                help_link(rg, helpm).grid(row=i, column=3, padx=(2, 0))

        # ═══ Buttons — packed under the form in `leftcol` so they
        # sit flush against the bottom of the Freundlich grid (no
        # empty gap between the form and the button row).
        bf = tk.Frame(leftcol, bg=BG_MAIN)
        bf.pack(anchor="w", pady=(8, 0))
        make_btn(bf, "Where to Get Freundlich\nParameters",
                 "OpenAppendix_9_1_Relative", quarto=True,
                 font=FONT_BTN, width=28, bg=BTN_FILL,
                 padx=10, pady=16
                 ).pack(side="left", padx=(0, 12))
        make_btn(bf, "Simple CAC Barrier\nLongevity Tool",
                 "LongevityTool", font=FONT_BTN, width=28, bg=BTN_FILL,
                 padx=10, pady=16
                 ).pack(side="left")

        # In Simple, hide registered Detailed-only widgets
        if self.active_sheet != "Detailed_2":
            for w in det_only:
                try:
                    w.grid_remove()
                except Exception:
                    pass

    # ─────────────────────────────────────────────────────────────────────
    # SECTION 10 – FIELD DATA TO CALIBRATE
    #
    # Layout depends on model version:
    #   SIMPLE  : 3-column inline table — header + Sample Year + helper
    #             notes | MW table | separator | Location Info + Distance
    #   DETAILED: header + a single big "Enter Monitoring Well Data" button
    # ─────────────────────────────────────────────────────────────────────
    def _build_s10_field_data(self, parent):
        outer = tk.Frame(parent, bg=BG_MAIN)
        # Tighter top padding — pulls §10 closer to §9 so the §9 button
        # row no longer leaves a visible empty band above it.
        outer.pack(fill="x", pady=(0, 4))

        # ═══ Section header (multi-line) — ALWAYS VISIBLE ═══════════════
        leftcol = tk.Frame(outer, bg=BG_MAIN)
        leftcol.pack(side="left", anchor="nw", padx=(0, 12))
        tk.Label(leftcol, text="10. FIELD\nDATA TO\nCALIBRATE",
                 font=FONT_SECTION, fg=FG_SECTION, bg=BG_MAIN,
                 justify="left", anchor="w"
                 ).pack(anchor="w", pady=(0, 12))

        # Sample Year + helper notes — SIMPLE ONLY
        sy_block = tk.Frame(leftcol, bg=BG_MAIN)
        sy_block.pack(anchor="w")
        tk.Label(sy_block, text="Sample Year (XXXX)",
                 font=FONT_LABEL_I, bg=BG_MAIN
                 ).pack(anchor="w")
        sy_row = tk.Frame(sy_block, bg=BG_MAIN); sy_row.pack(anchor="w")
        # Black value text on white bg — matches Excel (red is reserved
        # for the Distance-from-Source column only).
        tk.Entry(sy_row, textvariable=self.v_sample_yr, width=6,
                 font=FONT_INPUT, bg=BG_INPUT_BLUE, fg=FG_INPUT,
                 relief="solid", bd=1, justify="right"
                 ).pack(side="left", padx=(0, 4))
        help_link(sy_row, "OpenTable10_1").pack(side="left")
        helper_fr = tk.Frame(sy_block, bg=BG_MAIN)
        helper_fr.pack(anchor="e", pady=(16, 0))
        # Helper notes — black italic (Excel uses red only on the
        # Distance-from-Source column, not on these helper notes).
        for txt in ["Enter up to 7 shallow",
                    "centerline wells",
                    "for key monitoring event"]:
            tk.Label(helper_fr, text=txt, font=FONT_LABEL_SMI,
                     bg=BG_MAIN, fg=FG_INPUT, anchor="e"
                     ).pack(anchor="e")
        sy_block._toggle_kind = "pack"
        self._simple_only_frames.append(sy_block)

        # ═══ DETAILED-ONLY: just a big button ═══════════════════════════
        detailed_block = tk.Frame(outer, bg=BG_MAIN)
        detailed_block.pack(side="left", anchor="n", padx=(60, 0))
        # Big rounded button — italic blue text, gray fill.  Matches the
        # Excel storyboard reference: visually dominant, with whitespace
        # around it.
        make_btn(detailed_block, "Enter Monitoring\nWell Data",
                 "CalibrationDataLoader",
                 font=FONT_BTN_LG, width=30, bg=BTN_FILL,
                 padx=40, pady=24
                 ).pack(anchor="center", padx=20, pady=30)
        detailed_block._toggle_kind = "pack"
        self._detailed_only_frames.append(detailed_block)

        # ═══ SIMPLE-ONLY: MW table + separator + Location Info ══════════
        simple_block = tk.Frame(outer, bg=BG_MAIN)
        simple_block.pack(side="left", anchor="nw")
        simple_block._toggle_kind = "pack"
        self._simple_only_frames.append(simple_block)

        mid = tk.Frame(simple_block, bg=BG_MAIN)
        mid.pack(side="left", anchor="nw")

        # Row 0 — italic group titles
        tk.Label(mid, text="Monitoring", font=FONT_LABEL_BI,
                 bg=BG_MAIN, fg=FG_GREY).grid(row=0, column=0, padx=4)
        tk.Label(mid, text="Monitoring Well", font=FONT_LABEL_BI,
                 bg=BG_MAIN, fg=FG_GREY
                 ).grid(row=0, column=1, columnspan=2, padx=4)

        # Row 1 — sub-titles
        tk.Label(mid, text="Well Name", font=FONT_LABEL_BI,
                 bg=BG_MAIN, fg=FG_GREY).grid(row=1, column=0, padx=4)
        sub_fr = tk.Frame(mid, bg=BG_MAIN)
        sub_fr.grid(row=1, column=1, columnspan=2, padx=4, sticky="ew")
        tk.Label(sub_fr, text="Concentration Data (ug/L)",
                 font=FONT_LABEL_BI, bg=BG_MAIN, fg=FG_GREY
                 ).pack(side="left")
        help_link(sub_fr, "OpenTable10_3").pack(side="left", padx=(4, 0))

        # Row 2 — column headers (Event | PFOS | None) + Event help.
        # v102: "Event" stays a static black label (per user spec —
        # uneditable, just labels the row of well-name entries below).
        # The well NAMES in rows 3-9 are editable Entry widgets so the
        # user can type per-well monitoring point names there.
        hdr_event = tk.Frame(mid, bg=BG_MAIN)
        hdr_event.grid(row=2, column=0, padx=4)
        tk.Label(hdr_event, text="Event", font=FONT_LABEL_B,
                 bg=BG_LOCKED, fg="#FFFFFF",
                 relief="solid", bd=1, padx=10
                 ).pack(side="left")
        help_link(hdr_event, "OpenTable10_2").pack(side="left", padx=(4, 0))
        # Headers bound to v_pfaa1 / v_pfaa2 — they live-update when
        # the user changes the §5 PFAA dropdowns.  Was hardcoded
        # "PFOS" / "None" so changing PFAA in §5 left §10 stale.
        tk.Label(mid, textvariable=self.v_pfaa1, font=FONT_LABEL_B,
                 bg=BG_LOCKED, fg="#FFFFFF",
                 relief="solid", bd=1, width=12
                 ).grid(row=2, column=1, padx=4, sticky="ew")
        tk.Label(mid, textvariable=self.v_pfaa2, font=FONT_LABEL_B,
                 bg=BG_LOCKED, fg="#FFFFFF",
                 relief="solid", bd=1, width=12
                 ).grid(row=2, column=2, padx=4, sticky="ew")

        # Rows 3-9 — 7 monitoring well rows.
        # Well names: WHITE bg + BLACK text (right-justified, read-only)
        #   — matches Excel; previous black-bg/white-text was wrong.
        # Concentration values: BLACK text on white (red is reserved for
        # the Distance column only).
        for i in range(7):
            # v102: well-name entries are editable now (were state="readonly")
            # so the user can type custom monitoring well names directly
            # into §10 — matches the Excel reference.
            tk.Entry(mid, textvariable=self.v_mw_names[i], width=10,
                     font=FONT_INPUT, bg=BG_INPUT_BLUE, fg=FG_INPUT,
                     relief="solid", bd=1, justify="right",
                     ).grid(row=3+i, column=0, padx=2, pady=1)
            tk.Entry(mid, textvariable=self.v_mw_conc[i], width=12,
                     font=FONT_INPUT, bg=BG_INPUT_BLUE, fg=FG_INPUT,
                     relief="solid", bd=1, justify="right"
                     ).grid(row=3+i, column=1, padx=2, pady=1)
            tk.Entry(mid, textvariable=self.v_mw_conc2[i], width=12,
                     font=FONT_INPUT, bg="#FFFFFF", fg=FG_INPUT,
                     relief="solid", bd=1, justify="right"
                     ).grid(row=3+i, column=2, padx=2, pady=1)
            tk.Label(mid, text="(ug/L)", font=FONT_LABEL_SMI,
                     bg=BG_MAIN, fg=FG_INPUT
                     ).grid(row=3+i, column=3, sticky="w", padx=2)

        # Vertical separator
        sep = tk.Frame(simple_block, bg="#000000", width=2)
        sep.pack(side="left", fill="y", padx=(16, 16))

        # Right — Location Info + Distance
        rt = tk.Frame(simple_block, bg=BG_MAIN)
        rt.pack(side="left", anchor="nw")
        loc_hdr = tk.Frame(rt, bg=BG_MAIN)
        loc_hdr.grid(row=0, column=0, columnspan=2, sticky="w",
                     padx=4, pady=(0, 2))
        tk.Label(loc_hdr, text="Monitoring Well Location Info:",
                 font=FONT_LABEL_BI, bg=BG_MAIN, fg=FG_GREY
                 ).pack(side="left")
        dist_hdr = tk.Frame(rt, bg=BG_MAIN)
        dist_hdr.grid(row=1, column=1, sticky="e", padx=4)
        # v102: distance from source is a length → flips m↔ft
        _dist10_lbl = tk.Label(dist_hdr,
                 text=f"Distance\nfrom\nSource ({self._unit_len()})",
                 font=FONT_LABEL_BI, bg=BG_MAIN, fg=FG_GREY,
                 justify="center")
        _dist10_lbl.pack(side="left")
        self._register_unit_label(_dist10_lbl, "Distance\nfrom\nSource ({u})")
        help_link(dist_hdr, "OpenTable10_4").pack(side="left", padx=(4, 0))

        descr = [
            "For simple model use only",
            "monitoring wells on or",
            "near the top of the  plume",
            "",
            "and on the",
            "centerline of the plume",
            "",
        ]
        # Right-side descriptive notes — BLACK italic per Excel.
        # Distances stay RED on white — that's the only red column in §10.
        for i in range(7):
            tk.Label(rt, text=descr[i], font=FONT_LABEL_SMI,
                     bg=BG_MAIN, fg=FG_INPUT, anchor="e"
                     ).grid(row=2+i, column=0, sticky="e",
                            padx=(0, 8), pady=1)
            tk.Entry(rt, textvariable=self.v_mw_dist[i], width=10,
                     font=FONT_INPUT, bg="#FFFFFF", fg="#C00000",
                     relief="solid", bd=1, justify="right"
                     ).grid(row=2+i, column=1, padx=2, pady=1)

        # Hide the wrong block on first paint based on current version.
        if self.active_sheet == "Detailed_2":
            self._toggle_widgets([sy_block, simple_block], show=False)
        else:
            self._toggle_widgets([detailed_block], show=False)

    # SECTION 11
    def _build_s11_output(self, parent):
        # Single horizontal row containing 3 columns:
        #   LEFT   – multi-line header + Change Numerical Parameters btn
        #   MIDDLE – "See Results Every: [val] (years) ?"
        #   RIGHT  – blue action panel (Run Model + 3x2 button grid)
        outer = tk.Frame(parent, bg=BG_MAIN)
        outer.pack(fill="x", anchor="w", pady=(4, 4))

        # ── LEFT — header + Change Numerical Parameters ──────────────
        leftcol = tk.Frame(outer, bg=BG_MAIN)
        leftcol.pack(side="left", anchor="nw", padx=(0, 16))
        tk.Label(leftcol,
                 text="11. MODEL OUTPUT AND\nNUMERICAL PARAMETERS",
                 font=FONT_SECTION, fg=FG_SECTION, bg=BG_MAIN,
                 justify="left", anchor="w"
                 ).pack(anchor="w", pady=(0, 8))
        make_btn(leftcol, "Change Numerical Parameters",
                 "ChangeNumericalParameters",
                 font=FONT_BTN_SM, width=26, bg=BTN_FILL
                 ).pack(anchor="w")

        # ── MIDDLE — See Results Every: [entry] (years) ? ────────────
        midcol = tk.Frame(outer, bg=BG_MAIN)
        midcol.pack(side="left", anchor="n", padx=(0, 16))
        tk.Label(midcol, text="See Results", font=FONT_LABEL_I,
                 bg=BG_MAIN, justify="center"
                 ).pack(anchor="center")
        tk.Label(midcol, text="Every:", font=FONT_LABEL_I,
                 bg=BG_MAIN, justify="center"
                 ).pack(anchor="center")
        sr_row = tk.Frame(midcol, bg=BG_MAIN); sr_row.pack(anchor="center")
        make_entry(sr_row, self.v_see_every, width=6
                   ).pack(side="left", padx=(0, 2))
        tk.Label(sr_row, text="(years)", font=FONT_LABEL_SM, bg=BG_MAIN
                 ).pack(side="left")
        help_link(sr_row, "OpenTable11_1").pack(side="left", padx=(2, 0))

        # ── RIGHT — blue action panel (looser internal spacing) ────
        # Bumped panel padx/pady 14 → 20 to give the buttons more breathing
        # room against the panel border, matching the Excel reference.
        bar = tk.Frame(outer, bg=BTN_FILL_BLUE, bd=3, relief="solid",
                       padx=20, pady=20)
        bar.pack(side="left", anchor="n", fill="x", expand=True)

        # Two big Run Model buttons — gray fill (Excel storyboard)
        # while the surrounding panel stays blue.  side-pad 10 → 16.
        make_btn(bar, "Run Model", "RunPythonScript",
                 fg=FG_BTN_NAVY, font=FONT_BTN_LG, padx=18, pady=14,
                 bg=BTN_FILL, width=14).pack(side="left", padx=16)
        # Routes to the standalone Python DDS optimizer (cali_1.run).
        # Was wired to Show_Visualization which just kicked off a
        # single forward run — same as plain "Run Model", which made
        # the Auto-Cal button effectively a no-op duplicate.
        make_btn(bar, "Run Model with\nAuto-Calibration",
                 "ScrollToCalibration",
                 fg=FG_BTN_NAVY, font=FONT_BTN_LG, padx=18, pady=14,
                 bg=BTN_FILL, width=18).pack(side="left", padx=16)

        # 3x2 action grid on the right (5 buttons; bottom-left empty).
        # Gap to grid bumped 20 → 32 so the action group is clearly
        # separated from the Run buttons.  Cell padding bumped 6 → 10
        # for visibly looser spacing between buttons in both rows.
        actfr = tk.Frame(bar, bg=BTN_FILL_BLUE)
        actfr.pack(side="left", padx=(32, 0))
        cells = [
            (0, 0, "Authors",          "Authors",                 BTN_FILL_GREEN, FG_BTN_GREEN),
            (0, 1, "Load Data",        "Load_Data",               "#FFFFFF",      FG_BTN_NAVY),
            (0, 2, "Save Data",        "Save_Data",               BTN_FILL_GREEN, FG_BTN_GREEN),
            # v106: new — visualize the output of a previous run.  Sits in
            # the previously-empty slot directly below "Authors".  Styled
            # like "Load Data" (white fill, navy text).
            (1, 0, "Visualize\nResults", "Visualize_Saved_Results", "#FFFFFF", FG_BTN_NAVY),
            (1, 1, "Clear All Data",   "Clear_Data",              BTN_FILL,       FG_BTN_NAVY),
            (1, 2, "Paste Example",    "Paste_Example",           BTN_FILL,       FG_BTN_NAVY),
        ]
        for r, c, txt, macro, bg, fg in cells:
            make_btn(actfr, txt, macro,
                     fg=fg, font=FONT_BTN, padx=14, pady=10,
                     bg=bg, width=14
                     ).grid(row=r, column=c, padx=10, pady=10, sticky="ew")

    # BOTTOM BAR
    # ─────────────────────────────────────────────────────────────────────
    # ACTION ROW — Run / Authors / Load / Save / Help / Clear / Paste.
    # Lives inside the right column of the top body (per PDF storyboard).
    # ─────────────────────────────────────────────────────────────────────
    def _build_action_row(self, parent):
        # Per Excel: the whole action cluster sits inside a solid blue
        # rectangle (BTN_FILL_BLUE) at the bottom-right of upper page.
        bar = tk.Frame(parent, bg=BTN_FILL_BLUE, bd=1, relief="solid",
                       padx=6, pady=6)
        bar.pack(fill="x", anchor="ne", pady=(8, 0))

        # Two big Run Model buttons (blue filled, navy italic)
        make_btn(bar, "Run Model", "RunPythonScript",
                 fg=FG_BTN_NAVY, font=FONT_BTN_LG, padx=10, pady=4,
                 bg=BTN_FILL_BLUE, width=12).pack(side="left", padx=4)
        make_btn(bar, "Run Model with\nAuto-Calibration",
                 "ScrollToCalibration",
                 fg=FG_BTN_NAVY, font=FONT_BTN_LG, padx=10, pady=4,
                 bg=BTN_FILL_BLUE, width=18).pack(side="left", padx=4)

        # 3x2 action grid on the right.  Authors + Save Data = green;
        # Load Data white-ish; Clear/Paste/Help neutral gray.
        actfr = tk.Frame(bar, bg=BTN_FILL_BLUE)
        actfr.pack(side="right", padx=4)
        cells = [
            (0, 0, "Authors",        "Authors",       BTN_FILL_GREEN, FG_BTN_GREEN),
            (0, 1, "Load Data",      "Load_Data",     BG_INPUT_BLUE,  FG_BTN_NAVY),
            (0, 2, "Save Data",      "Save_Data",     BTN_FILL_GREEN, FG_BTN_GREEN),
            (1, 0, "Help",           "Help",          BTN_FILL,       FG_BTN_NAVY),
            (1, 1, "Clear All Data", "Clear_Data",    BTN_FILL,       FG_BTN_NAVY),
            (1, 2, "Paste Example",  "Paste_Example", BTN_FILL,       FG_BTN_NAVY),
        ]
        for r, c, txt, macro, bg, fg in cells:
            make_btn(actfr, txt, macro,
                     fg=fg, font=FONT_BTN, padx=8, pady=3,
                     bg=bg, width=12).grid(
                row=r, column=c, padx=3, pady=2, sticky="ew")

    # ─────────────────────────────────────────────────────────────────────
    # CALIBRATION PANEL — bottom-right quadrant per the Excel storyboard.
    # Layout matches the Simple + Detailed reference screenshots exactly:
    #   • Header band: title + 5 step lines (left)  |  iterations + time
    #     remaining + caveat (right)
    #   • Three-box row: Step 2 (Calibrate using) | Step 3 (Weighting
    #     factor table) | LEGEND
    #   • Step 4: Enter Calibration Ranges — table with parameter
    #     checkboxes / Lowest / Mid-Range (black, locked) / Highest
    #     plus units and Section refs.  Detailed mode adds 6 Precursor
    #     rows after a small gap.
    #   • Action buttons centered in 3 rows.
    # ─────────────────────────────────────────────────────────────────────
    def _build_calibration_panel(self, parent):
        outer = tk.Frame(parent, bg=BG_MAIN)
        outer.pack(fill="both", expand=True, padx=8, pady=6)
        # Save the panel's outer frame so scroll_to_calibration() can
        # compute the exact yview fraction needed to bring it to the
        # top of the visible canvas — survives any layout reshuffle.
        self._calib_panel_frame = outer

        # ── Top header band: title + step lines (left) | iter/time (right)
        # Two-column GRID rather than left/right pack, so the right-side
        # iter/time block sits directly next to the right edge of the
        # step-line text instead of floating off in its own column.
        header_band = tk.Frame(outer, bg=BG_MAIN)
        header_band.pack(fill="x", pady=(0, 4))
        header_band.columnconfigure(0, weight=1)   # left text grows
        header_band.columnconfigure(1, weight=0)   # right block fixed

        head_l = tk.Frame(header_band, bg=BG_MAIN)
        head_l.grid(row=0, column=0, sticky="nw")
        _cal_title_row = tk.Frame(head_l, bg=BG_MAIN)
        _cal_title_row.pack(anchor="w")
        tk.Label(_cal_title_row,
                 text="REMFluor-MD MACHINE-CALIBRATION (Singh et al., 2025)",
                 font=FONT_LABEL_B, bg=BG_MAIN, fg=FG_INPUT, anchor="w"
                 ).pack(side="left")
        help_link(_cal_title_row, "OpenAppendix_11_1_Relative").pack(side="left", padx=(6, 0))
        tk.Label(head_l,
            text="This allows you to let the computer perform a simple "
                 "calibration of your REMFluor-MD model by:",
            font=FONT_LABEL, bg=BG_MAIN, fg=FG_INPUT, anchor="w",
            justify="left"
            ).pack(anchor="w", pady=(2, 0))
        # 5 Step lines — red "Step N)" tag, black body, single line each.
        # No wraplength = the label sizes to its full text on one line.
        steps = [
            ("Step 1)", "Entering monitoring well data in Section 9."),
            ("Step 2)", "Select which parameters to calibrate against."),
            ("Step 3)", "Enter calibration options:  decide if you want "
                        "to weight any monitoring data more or less "
                        "during calibration."),
            ("Step 4)", "Enter calibration data, the likely minimum and "
                        "maximum values for the selected parameters below."),
            ("Step 5)", 'Hitting the "Run Machine-Based Calibration" '
                        "button to the right."),
        ]
        for tag, body_txt in steps:
            srow = tk.Frame(head_l, bg=BG_MAIN)
            srow.pack(anchor="w", padx=14, fill="x")
            tk.Label(srow, text=tag, font=FONT_LABEL, bg=BG_MAIN,
                     fg=FG_HELP).pack(side="left")
            tk.Label(srow, text=" " + body_txt, font=FONT_LABEL,
                     bg=BG_MAIN, fg=FG_INPUT,
                     justify="left", anchor="w"
                     ).pack(side="left")

        # Right side: iterations + time remaining — sits flush against
        # the right edge of the step lines (small 24-px gap, not floating
        # to the far right of the panel).
        head_r = tk.Frame(header_band, bg=BG_MAIN)
        head_r.grid(row=0, column=1, sticky="nw", padx=(24, 4))
        nf = tk.Frame(head_r, bg=BG_MAIN); nf.pack(anchor="w")
        tk.Label(nf, text="Number of iteration", font=FONT_LABEL,
                 bg=BG_MAIN).pack(side="left")
        self.v_n_iter = tk.StringVar(value="50")
        tk.Entry(nf, textvariable=self.v_n_iter, width=6,
                 font=FONT_INPUT, bg=BG_INPUT_BLUE, fg=FG_INPUT,
                 relief="solid", bd=1, justify="right"
                 ).pack(side="left", padx=(6, 0))
        tf = tk.Frame(head_r, bg=BG_MAIN); tf.pack(anchor="w", pady=(2, 0))
        tk.Label(tf, text="Estim.Time Remaining", font=FONT_LABEL,
                 bg=BG_MAIN).pack(side="left")
        self.v_t_remain = tk.StringVar(value="0")
        tk.Entry(tf, textvariable=self.v_t_remain, width=6,
                 font=FONT_INPUT, bg=BG_LOCKED, fg=FG_LOCKED,
                 relief="solid", bd=1, justify="right",
                 readonlybackground=BG_LOCKED, state="readonly"
                 ).pack(side="left", padx=(6, 0))
        time_unit = "(hours)" if self.active_sheet == "Detailed_2" \
                    else "(minutes)"
        tk.Label(tf, text=time_unit, font=FONT_LABEL_SM,
                 bg=BG_MAIN).pack(side="left", padx=(2, 0))

        # ── Live update of Estim. Time Remaining ────────────────────────
        # Was a static black cell with hardcoded "2".  Now:
        #  • When idle: typing in "Number of iteration" rough-estimates
        #    total run time = N × per_iter_secs.  Detailed runs are
        #    typically 4× slower than Simple due to the larger model.
        #  • When the calibration is running, functions/cali_1.py's
        #    progress callback overwrites this cell with the actual
        #    remaining time computed from elapsed / iter pace
        #    (see _cali_progress_update below).
        self._cali_idle_per_iter_secs = (
            120 if self.active_sheet == "Detailed_2" else 30)
        self._cali_running = False    # set True by cali_1.py while DDS runs
        self._cali_t_unit  = time_unit  # "(minutes)" or "(hours)"

        def _format_t_remain(secs):
            if self._cali_t_unit.startswith("(hours)"):
                v = secs / 3600.0
                if v < 0.1: return f"{v:.2f}"
                if v < 10:  return f"{v:.1f}"
                return f"{int(round(v))}"
            else:
                v = secs / 60.0
                if v < 1:   return f"{v:.1f}"
                if v < 10:  return f"{v:.1f}"
                return f"{int(round(v))}"
        self._format_t_remain = _format_t_remain

        def _refresh_estim_time(*_):
            if self._cali_running:
                return    # actual progress drives the cell during a run
            try:
                n = int(float(str(self.v_n_iter.get()).strip()))
            except (ValueError, TypeError):
                self.v_t_remain.set("0")
                return
            secs = max(0, n) * self._cali_idle_per_iter_secs
            self.v_t_remain.set(_format_t_remain(secs))
        self._refresh_estim_time = _refresh_estim_time
        self.v_n_iter.trace_add("write", _refresh_estim_time)
        _refresh_estim_time()  # initial paint reflects default 50 iters

        def _cali_progress_update(i, total, elapsed_secs):
            """Called from functions/cali_1.py after each DDS eval.
            Writes the running ETA into v_t_remain."""
            if not self._cali_running:
                return
            if i <= 0:
                return
            per_iter = elapsed_secs / max(i, 1)
            remain = max(0.0, per_iter * (total - i))
            try:
                self.v_t_remain.set(_format_t_remain(remain))
            except Exception:
                pass
        self._cali_progress_update = _cali_progress_update
        tk.Label(head_r,
                 text="(add explainer text here about\npossible run times)",
                 font=FONT_LABEL_SMI, fg=FG_HELP, bg=BG_MAIN,
                 justify="left").pack(anchor="w", pady=(4, 0))

        # ── Three-box row: Step 2 | Step 3 | LEGEND ─────────────────────
        body = tk.Frame(outer, bg=BG_MAIN)
        body.pack(fill="x", expand=True, pady=(8, 0))

        # --- Step 2: Calibrate using ---
        s2 = tk.Frame(body, bg=BG_MAIN, bd=1, relief="solid",
                      padx=12, pady=10)
        s2.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        tk.Label(s2, text="Step 2:  Calibrate using:",
                 font=FONT_LABEL_B, bg=BG_MAIN, anchor="w"
                 ).grid(row=0, column=0, columnspan=2, sticky="w",
                        pady=(0, 6))
        # Step 2 checkboxes — bound to BooleanVars; the label text is
        # bound to the §5 PFAA dropdowns (v_pfaa1..4) so they live-
        # update when the user changes which compound to calibrate.
        # Each row's tk.Checkbutton uses `text=""` and the label is a
        # separate widget driven by `textvariable=` — keeps the bind
        # simple and lets the checkbox click target be tight.
        self.v_calib_pfoa     = tk.BooleanVar(value=True)
        self.v_calib_none     = tk.BooleanVar(value=False)
        self.v_calib_pre      = tk.BooleanVar(value=True)
        self.v_calib_pre_none = tk.BooleanVar(value=False)

        # Row 1 — PFAA-1 (textvariable on label so it tracks v_pfaa1).
        # Uses the display-mirror so a cleared dropdown still renders a
        # meaningful label next to the radio (falls back to "PFOS").
        tk.Label(s2, textvariable=self._v_pfaa1_disp, font=FONT_LABEL,
                 bg=BG_MAIN
                 ).grid(row=1, column=0, sticky="e", padx=(8, 4),
                        pady=2)
        big_check(s2, "", self.v_calib_pfoa, bg=BG_MAIN
                  ).grid(row=1, column=1, sticky="w", pady=2)
        # Row 2 — PFAA-2.  Was previously bound to v_pfaa2 directly,
        # which left the label blank whenever §5 PFAA-2 was empty.
        # The display-mirror falls back to "None" so the radio always
        # has a label.
        tk.Label(s2, textvariable=self._v_pfaa2_disp, font=FONT_LABEL,
                 bg=BG_MAIN
                 ).grid(row=2, column=0, sticky="e", padx=(8, 4),
                        pady=2)
        big_check(s2, "", self.v_calib_none, bg=BG_MAIN
                  ).grid(row=2, column=1, sticky="w", pady=2)
        # Detailed-only: rows 3-4 mirror Precursor 1/2 species
        # (display-mirrors so cleared dropdowns still render "None").
        det_lbl1 = tk.Label(s2, textvariable=self._v_pfaa3_disp,
                            font=FONT_LABEL, bg=BG_MAIN)
        det_chk1 = big_check(s2, "", self.v_calib_pre, bg=BG_MAIN)
        det_lbl2 = tk.Label(s2, textvariable=self._v_pfaa4_disp,
                            font=FONT_LABEL, bg=BG_MAIN)
        det_chk2 = big_check(s2, "", self.v_calib_pre_none, bg=BG_MAIN)
        det_lbl1.grid(row=3, column=0, sticky="e", padx=(8, 4), pady=2)
        det_chk1.grid(row=3, column=1, sticky="w", pady=2)
        det_lbl2.grid(row=4, column=0, sticky="e", padx=(8, 4), pady=2)
        det_chk2.grid(row=4, column=1, sticky="w", pady=2)
        for w in (det_lbl1, det_chk1, det_lbl2, det_chk2):
            w._toggle_kind = "grid"
            self._detailed_only_frames.append(w)

        # --- Step 3: Calibration Options (weighting) ---
        s3 = tk.Frame(body, bg=BG_MAIN, bd=1, relief="solid",
                      padx=12, pady=10)
        s3.grid(row=0, column=1, sticky="nsew", padx=(0, 6))
        tk.Label(s3, text="Step 3:  Calibration Options - "
                          "enter any weighting factors.",
                 font=FONT_LABEL_B, bg=BG_MAIN, anchor="w"
                 ).pack(anchor="w", pady=(0, 4))
        s3g = tk.Frame(s3, bg=BG_MAIN); s3g.pack()
        # v102: distance column flips m↔ft with §1 units selection
        _calib_hdrs = ["Monitoring\nPoint\nName",
                       f"Distance\nfrom\nSource ({self._unit_len()})",
                       "Weighting\nFactor\nfor Calb:",
                       "Source:"]
        for j, h in enumerate(_calib_hdrs):
            _hlbl = tk.Label(s3g, text=h, font=FONT_LABEL_BI, bg=BG_MAIN,
                             fg=FG_INPUT)
            _hlbl.grid(row=0, column=j, sticky="w", padx=4, pady=(0, 2))
            if j == 1:
                self._register_unit_label(_hlbl,
                    "Distance\nfrom\nSource ({u})")
        self.v_calib_w = []
        # Default weighting = 1.0 for every well.  Was "1,1,1,1,1,1,2"
        # — the trailing 2 came from the Excel example's specific
        # downgradient-emphasis preset, not a sensible default for a
        # blank app.  Paste Example loads per-row weights from
        # CalibrationTemplate_*.xlsx, which restores the example's
        # weights when the user actually wants the example.
        defaults = ["1.0"] * 7
        for i, w in enumerate(defaults):
            tk.Entry(s3g, textvariable=self.v_mw_names[i], width=10,
                     font=FONT_INPUT, bg=BG_LOCKED, fg=FG_LOCKED,
                     relief="solid", bd=1, justify="center",
                     state="readonly", readonlybackground=BG_LOCKED
                     ).grid(row=i+1, column=0, sticky="w", padx=2,
                            pady=1)
            tk.Entry(s3g, textvariable=self.v_mw_dist[i], width=8,
                     font=FONT_INPUT, bg=BG_LOCKED, fg=FG_LOCKED,
                     relief="solid", bd=1, justify="right",
                     state="readonly", readonlybackground=BG_LOCKED
                     ).grid(row=i+1, column=1, sticky="w", padx=2,
                            pady=1)
            wv = tk.StringVar(value=w); self.v_calib_w.append(wv)
            tk.Entry(s3g, textvariable=wv, width=8, font=FONT_INPUT,
                     bg=BG_FORMULA, fg=FG_INPUT, relief="solid", bd=1,
                     justify="right"
                     ).grid(row=i+1, column=2, sticky="w", padx=2,
                            pady=1)
            tk.Label(s3g, text="Section 10", font=FONT_LABEL_I,
                     bg=BG_MAIN
                     ).grid(row=i+1, column=3, sticky="w", padx=4,
                            pady=1)
        tk.Label(s3,
                 text="(Often downgradient monitoring wells are "
                      "weighted higher)",
                 font=FONT_LABEL_SMI, fg=FG_HELP, bg=BG_MAIN
                 ).pack(anchor="center", pady=(6, 0))

        # --- LEGEND ---
        legend = tk.Frame(body, bg=BG_MAIN, bd=1, relief="solid",
                          padx=12, pady=10)
        legend.grid(row=0, column=2, sticky="nsew")
        tk.Label(legend, text="LEGEND", font=FONT_LABEL_B, bg=BG_MAIN
                 ).grid(row=0, column=0, columnspan=2, sticky="w",
                        pady=(0, 4))
        legend_items = [
            (BG_INPUT_BLUE, "Enter value directly"),
            (BG_FORMULA,    "Formula or default, but ok to overwrite**"),
            (BG_PULLDOWN,   "Pull Down Menu"),
            (BG_LOCKED,     "Calculated value or taken from other cell."),
        ]
        for r, (color, txt) in enumerate(legend_items, start=1):
            tk.Frame(legend, bg=color, width=22, height=18, bd=1,
                     relief="solid"
                     ).grid(row=r, column=0, sticky="w", padx=(0, 8),
                            pady=3)
            tk.Label(legend, text=txt, font=FONT_LABEL_SMI,
                     bg=BG_MAIN, anchor="w", justify="left",
                     wraplength=220
                     ).grid(row=r, column=1, sticky="w", pady=3)

        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=2)
        body.columnconfigure(2, weight=1)

        # ── Step 4: Enter Calibration Ranges ─────────────────────────────
        s4_hdr = tk.Frame(outer, bg=BG_MAIN)
        s4_hdr.pack(fill="x", pady=(12, 0))
        tk.Label(s4_hdr, text="Step 4:  Enter Calibration Ranges:",
                 font=FONT_LABEL_B, bg=BG_MAIN, anchor="w"
                 ).pack(side="left")
        # "Mid-Range Values are taken from the input cells in these
        # sections:" — yellow-highlighted lozenge + italic body.
        rs_note = tk.Label(s4_hdr,
                           text="are taken from in the input cells in "
                                "these sections:",
                           font=FONT_LABEL_I, bg=BG_MAIN, fg=FG_INPUT)
        rs_y    = tk.Label(s4_hdr, text="Mid-Range Values",
                           font=FONT_LABEL_BI, bg=FG_YELLOW, fg=FG_INPUT,
                           padx=4)
        rs_note.pack(side="right", padx=(0, 4))
        rs_y.pack(side="right", padx=(0, 4))

        make_btn(outer,
                 "Enter Default Low-High Ranges from Literature, "
                 "Experience (default +/- x2)",
                 "DefaultRanges",
                 fg=FG_BTN_NAVY, font=FONT_BTN_SM, padx=10, pady=4,
                 bg=BTN_FILL).pack(anchor="w", pady=(2, 6))

        # Ranges table — column layout:
        #   0: Param name (italic, right-aligned)
        #   1: Use this Parameter? (checkbox)
        #   2: Lowest Likely Value      (sky blue)
        #   3: Mid-Range Value          (BLACK locked, white text)
        #   4: Highest Likely Value     (sky blue)
        #   5: Unit  e.g. (m/yr) (-)
        #   6: Section reference        (italic)
        #   7: Right-side red note      (multiplier rows only)
        tbl = tk.Frame(outer, bg=BG_MAIN); tbl.pack(anchor="w", pady=2,
                                                     fill="x")
        # Header row — only columns 1-4 have text; "Use this Parameter?"
        # spans the checkbox column.
        tk.Label(tbl, text="Use this\nParameter?", font=FONT_LABEL_BI,
                 bg=BG_MAIN, fg=FG_INPUT
                 ).grid(row=0, column=1, padx=2, pady=2)
        tk.Label(tbl, text="Lowest\nLikely Value", font=FONT_LABEL_BI,
                 bg=BG_MAIN, fg=FG_INPUT
                 ).grid(row=0, column=2, padx=2, pady=2)
        tk.Label(tbl, text="Mid-Range\nValue", font=FONT_LABEL_BI,
                 bg=FG_YELLOW, fg=FG_INPUT, padx=4
                 ).grid(row=0, column=3, padx=2, pady=2, sticky="ew")
        tk.Label(tbl, text="Highest\nLikely Value", font=FONT_LABEL_BI,
                 bg=BG_MAIN, fg=FG_INPUT
                 ).grid(row=0, column=4, padx=2, pady=2)

        # Each row: (label, lo, mid, hi, unit, section)
        simple_rows = [
            ("Source Start Year (nt)",                     "1965",   "1977",   "1980",  "(xxxx)",     "Section 2"),
            ("Hydraulic Conductivity (k)",                 "0.31536","3.15",   "31.536","(m/yr)",     "Section 3"),
            ("Hydraulic Gradient (i)",                     "0.0004", "0.0038", "0.0380","(-)",        "Section 3"),
            ("Effective Porosity (porf)",                  "0.16",   "0.20",   "0.24",  "(-)",        "Section 3"),
            ("Transmissive Fraction of Model (volfrac)",   "0.6",    "1.00",   "0.85",  "(-)",        "Section 4"),
            ("Average Diffusion Length (difflen)",         "1.5",    "0.00",   "6",     "(m)",        "Section 4"),
            ("Retardation Factor of PFAA-1 (ock(2))",      "1.6",    "2.9",    "6.4",   "(-)",        "Section 5"),
            ("Retardation Factor of PFAA-2 (ock(4))",      "1",      "0.0",    "4",     "(-)",        "Section 5"),
            ("Longitudinal Dispersivity (alphax (m))",     "1",      "3.2",    "1.5",   "(m)",        "Section 6"),
            ("Multiplier to PFAA-1 Source Concentration in #7 (czero(2,n))",
                                                            "0.5",   "1",      "2",     "x(Ct)",      "Section 7"),
            ("Multiplier to PFAA-2 Source Concentration in #7 (czero(4,n))",
                                                            "0.5",   "#DIV/0!","2",     "x(Ct)",      "Section 7"),
        ]
        precursor_rows = [
            ("First order decay rate coefficient for Precursors-1 (decayf(1))",
                                                            "0.5",   "0.4",    "4.5",   "(per year)", "Section 5"),
            ("First order decay rate coefficient for Precursors-2 (decayf(3))",
                                                            "0.5",   "0.0",    "4.5",   "(per year)", "Section 5"),
            ("Retardation Factor of Precursors-1 (ock(1))",
                                                            "2",     "4.7",    "4",     "(-)",        "Section 5"),
            ("Retardation Factor of Precursors-2 (ock(3))",
                                                            "2",     "0.0",    "4",     "(-)",        "Section 5"),
            ("Multiplier to Precursor-1 Source Concentration in #7 (czero(1,n))",
                                                            "0.5",   "1.0",    "2",     "x(Ct)",      "Section 7"),
            ("Multiplier to Precursor-2 Source Concentration in #7 (czero(3,n))",
                                                            "0.5",   "#DIV/0!","2",     "x(Ct)",      "Section 7"),
        ]
        # Pre-checked per the screenshot defaults
        # v88: Precursor-1 multiplier was previously hard-coded as
        # pre-checked, but precursors are only relevant in Detailed
        # mode AND only when K38/M38 has a name (ipre=1).  Pre-checking
        # it when no precursor is configured caused the multiplier to
        # leak into run_history.csv even on PFAA-only calibrations.
        # Now precursor rows start UNCHECKED — user must opt in.
        preselected = {
            "Hydraulic Conductivity (k)",
            "Hydraulic Gradient (i)",
            "Multiplier to PFAA-1 Source Concentration in #7 (czero(2,n))",
        }
        self.v_calib_chk  = []
        self.v_calib_low  = []
        self.v_calib_mid  = []
        self.v_calib_high = []

        # K and i are user-editable EVEN in the Mid column (sky-blue
        # input style) — the GW Velocity Calculator is the primary
        # source, but the user can override here without leaving the
        # calibration panel.  Every other row's Mid is BG_LOCKED black.
        EDITABLE_MID_LABELS = {
            "Hydraulic Conductivity (k)",
            "Hydraulic Gradient (i)",
        }

        def _add_row(i, lbl, lo, mid, hi, unit, section,
                     detailed_only=False):
            wname = tk.Label(tbl, text=lbl, font=FONT_LABEL_I,
                              bg=BG_MAIN, anchor="e", justify="right")
            wname.grid(row=i, column=0, sticky="e", padx=(2, 4),
                       pady=1)
            cv = tk.BooleanVar(value=(lbl in preselected))
            self.v_calib_chk.append(cv)
            wchk = big_check(tbl, "", cv, bg=BG_MAIN)
            wchk.grid(row=i, column=1, pady=1)
            cells = []
            mid_bg = (BG_FORMULA if lbl in EDITABLE_MID_LABELS
                      else BG_LOCKED)
            for col, vlist, val, bg in [
                (2, self.v_calib_low,  lo,  BG_FORMULA),
                (3, self.v_calib_mid,  mid, mid_bg),
                (4, self.v_calib_high, hi,  BG_FORMULA),
            ]:
                vv = tk.StringVar(value=val); vlist.append(vv)
                state_kw = {}
                fg = FG_INPUT
                if bg == BG_LOCKED:
                    fg = FG_LOCKED
                    state_kw = dict(state="readonly",
                                    readonlybackground=BG_LOCKED)
                e = tk.Entry(tbl, textvariable=vv, width=10,
                             font=FONT_INPUT, bg=bg, fg=fg,
                             relief="solid", bd=1, justify="right",
                             **state_kw)
                e.grid(row=i, column=col, padx=1, pady=1)
                cells.append(e)
            # v102: if unit string contains "(m)" or "(m/yr)" treat as
            # length-unit label that flips m↔ft with §1 toggle.
            _unit_fmt = None
            _disp_unit = unit
            if unit == "(m/yr)":
                _unit_fmt = "({u}/yr)"
                _disp_unit = _unit_fmt.format(u=self._unit_len())
            elif unit == "(m)":
                _unit_fmt = "({u})"
                _disp_unit = _unit_fmt.format(u=self._unit_len())
            wunit = tk.Label(tbl, text=_disp_unit, font=FONT_LABEL,
                             bg=BG_MAIN, anchor="w")
            wunit.grid(row=i, column=5, sticky="w", padx=(4, 2),
                       pady=1)
            if _unit_fmt is not None:
                self._register_unit_label(wunit, _unit_fmt)
            wsec = tk.Label(tbl, text=section, font=FONT_LABEL_I,
                             bg=BG_MAIN, anchor="w")
            wsec.grid(row=i, column=6, sticky="w", padx=(2, 4),
                      pady=1)
            if detailed_only:
                for w in (wname, wchk, *cells, wunit, wsec):
                    w._toggle_kind = "grid"
                    self._detailed_only_frames.append(w)

        ridx = 1
        for r in simple_rows:
            _add_row(ridx, *r)
            ridx += 1
        # Right-side red note next to the 2 PFAA multiplier rows
        note_pfaa = tk.Label(tbl,
            text="(All 11 source concentrations for PFAA-1\n"
                 "are changed by the same amount during\n"
                 "calibration runs).",
            font=FONT_LABEL_SMI, fg=FG_HELP, bg=BG_MAIN,
            justify="left")
        note_pfaa.grid(row=ridx - 2, column=7, rowspan=2, sticky="nw",
                       padx=(10, 0))

        # Detailed-only blank gap then precursor rows
        gap = tk.Frame(tbl, bg=BG_MAIN, height=10)
        gap.grid(row=ridx, column=0, columnspan=8)
        gap._toggle_kind = "grid"
        self._detailed_only_frames.append(gap)
        ridx += 1
        for r in precursor_rows:
            _add_row(ridx, *r, detailed_only=True)
            ridx += 1
        note_pre = tk.Label(tbl,
            text="(All 11 source concentrations for Precursor-1\n"
                 "are changed by the same amount during\n"
                 "calibration runs).",
            font=FONT_LABEL_SMI, fg=FG_HELP, bg=BG_MAIN,
            justify="left")
        note_pre.grid(row=ridx - 2, column=7, rowspan=2, sticky="nw",
                      padx=(10, 0))
        note_pre._toggle_kind = "grid"
        self._detailed_only_frames.append(note_pre)

        # ── Step 4 Mid-Range column = LIVE from input cells ──────────
        # User feedback: "the values aren't dynamic — should be pulled
        # from the same cells in the app whenever the user clicks the
        # checkbox option, then the black cell should change".
        #
        # SRC_MAP entries are either:
        #   * a tk.StringVar  → .get() value
        #   * a callable      → returns the string to display
        #   * None            → blank by default
        # Callables are used for sidecar-file readers (GW Velocity
        # Calculator → gwvelocity_inputs.txt; Heterogeneity Calculator
        # → heterogeneity_inputs.txt) so changes made in those popups
        # propagate to the calibration Mid column on the next refresh.
        UNIT_TO_MYR = {
            "m/s":     86400.0 * 365.25,
            "m/day":   365.25,
            "m/year":  1.0,
            "cm/s":    864.0   * 365.25,
            "cm/day":  0.01    * 365.25,
            "mm/day":  0.001   * 365.25,
            "ft/day":  0.3048  * 365.25,
            "ft/s":    26334.72 * 365.25,
            "in/day":  0.0254  * 365.25,
        }

        def _read_gw_inputs():
            """Return {'k_value': float, 'k_unit': str, 'gradient': float}
            from gwvelocity_inputs.txt (written by §3 GW Velocity
            Calculator).  Empty dict if the file's missing."""
            path = os.path.join(BASE_DIR, "gwvelocity_inputs.txt")
            out = {}
            if not os.path.exists(path):
                return out
            try:
                with open(path) as f:
                    for ln in f:
                        ln = ln.strip()
                        if "Bulk Hydraulic Conductivity Value" in ln and ":" in ln:
                            try: out["k_value"] = float(ln.split(":", 1)[1].strip())
                            except ValueError: pass
                        elif "Bulk Hydraulic Conductivity Unit" in ln and ":" in ln:
                            out["k_unit"] = ln.split(":", 1)[1].strip()
                        elif ln.startswith("Bulk Hydraulic Gradient") and ":" in ln:
                            try: out["gradient"] = float(ln.split(":", 1)[1].strip())
                            except ValueError: pass
            except Exception:
                pass
            return out

        def _read_hetero_inputs():
            """Return {'volfrac': float, 'difflen': float} from
            heterogeneity_inputs.txt.  Empty dict if missing."""
            path = os.path.join(BASE_DIR, "heterogeneity_inputs.txt")
            out = {}
            if not os.path.exists(path):
                return out
            try:
                with open(path) as f:
                    for ln in f:
                        ln = ln.strip()
                        if ln.startswith("Transmissive Fraction of Model") and ":" in ln:
                            try: out["volfrac"] = float(ln.split(":", 1)[1].strip())
                            except ValueError: pass
                        elif ln.startswith("Diffusion Length") and ":" in ln:
                            try: out["difflen"] = float(ln.split(":", 1)[1].strip())
                            except ValueError: pass
            except Exception:
                pass
            return out

        # Callable sources return "" when they have NO authoritative
        # value to publish.  _refresh_one_mid below treats empty as
        # "leave the cell alone" — that way the calibration optimizer's
        # output (which lives only in v_calib_mid for rows without a
        # writable §StringVar) doesn't get clobbered by a refresh
        # firing later.  Hardcoded fallbacks (i=1, multiplier=1) are
        # only seeded ONCE at build time + on Paste Example / Clear,
        # never via auto-refresh.
        def _src_K():
            d = _read_gw_inputs()
            if "k_value" in d:
                k_myr = d["k_value"] * UNIT_TO_MYR.get(
                    d.get("k_unit", "m/day"), 365.25)
                return f"{k_myr:g}"
            # Fall back to v_darcy IF it has a value — else blank.
            try:
                v = str(self.v_darcy.get()).strip()
                return v if v else ""
            except Exception:
                return ""

        def _src_i():
            d = _read_gw_inputs()
            if "gradient" in d:
                return f"{d['gradient']:g}"
            return ""   # blank → don't overwrite Mid

        def _src_volfrac():
            d = _read_hetero_inputs()
            if "volfrac" in d:
                return f"{d['volfrac']:g}"
            return ""

        def _src_difflen():
            d = _read_hetero_inputs()
            if "difflen" in d:
                return f"{d['difflen']:g}"
            return ""

        SRC_MAP = {
            "Source Start Year (nt)":                                              self.v_yr_start,
            "Hydraulic Conductivity (k)":                                          _src_K,
            "Hydraulic Gradient (i)":                                              _src_i,
            "Effective Porosity (porf)":                                           self.v_porf,
            "Transmissive Fraction of Model (volfrac)":                            _src_volfrac,
            "Average Diffusion Length (difflen)":                                  _src_difflen,
            "Retardation Factor of PFAA-1 (ock(2))":                               self.v_ret_trans1,
            "Retardation Factor of PFAA-2 (ock(4))":                               self.v_ret_trans2,
            "Longitudinal Dispersivity (alphax (m))":                              self.v_alpha_l,
            "Multiplier to PFAA-1 Source Concentration in #7 (czero(2,n))":        None,
            "Multiplier to PFAA-2 Source Concentration in #7 (czero(4,n))":        None,
            "First order decay rate coefficient for Precursors-1 (decayf(1))":     None,
            "First order decay rate coefficient for Precursors-2 (decayf(3))":     None,
            "Retardation Factor of Precursors-1 (ock(1))":                         self.v_ret_trans3,
            "Retardation Factor of Precursors-2 (ock(3))":                         self.v_ret_trans4,
            "Multiplier to Precursor-1 Source Concentration in #7 (czero(1,n))":   None,
            "Multiplier to Precursor-2 Source Concentration in #7 (czero(3,n))":   None,
        }

        def _refresh_one_mid(label, idx):
            """Pull the source value (if any) into v_calib_mid[idx].
            DEFERENTIAL:  only overwrite Mid if the source actually
            has something to publish.  This way calibration results
            written into the Mid cell (which has no writable §
            source for rows like multipliers, i, volfrac, difflen)
            don't get clobbered by a refresh firing later."""
            if idx >= len(self.v_calib_mid):
                return
            src = SRC_MAP.get(label)
            if src is None:
                # No source — never auto-overwrite.  Initial baseline
                # (e.g. "1" for multipliers) is seeded at build time
                # below.
                return
            try:
                val = src() if callable(src) else str(src.get()).strip()
            except Exception:
                val = ""
            if val == "":
                return   # blank source → leave Mid alone
            try:
                self.v_calib_mid[idx].set(val)
            except Exception:
                pass

        # Guard flag: when True, a programmatic push is in flight
        # (Save Optimal → push back to source cells, Load Optimal,
        # etc.).  Source-cell trace handlers must NOT fire during
        # that window because they'd read stale sidecar files
        # (e.g. gwvelocity_inputs.txt) and overwrite the freshly-
        # loaded Mid values.  The _push_* helpers + _load_optimal_
        # model both flip this around their writes.
        self._calib_pushing = False

        def _refresh_all_mids(*_):
            if getattr(self, "_calib_pushing", False):
                return
            for i, lbl in enumerate(_CALIB_PARAMS):
                _refresh_one_mid(lbl, i)
        self._refresh_calib_mids = _refresh_all_mids

        # ── Reverse direction: calib Mid → app source cells ──────────
        # Used by the calibration finish (pushes best DDS values into
        # the §1-§7 inputs) and by Load Optimal Data (pushes the
        # saved optimal parameters back).  Without this the §calib
        # panel would show the optimum but Run Optimal Model would
        # use the user's pre-calibration inputs.  Reverse mapping
        # mirrors SRC_MAP — only StringVar sources can be written
        # back; callable sources (sidecar files) are skipped because
        # they'd require re-writing the underlying txt file, which
        # is a deliberate user action via the calculator popups.
        # K is intentionally NOT mapped here — v_darcy is computed
        # JOINTLY from K × i in _push_calib_mids_to_inputs below
        # (matches the §3 GW Velocity Calculator's formula).
        # Mapping K → v_darcy directly would overwrite the i factor
        # and produce vd = K instead of vd = K × i.
        REVERSE_MAP = {
            "Source Start Year (nt)":                  self.v_yr_start,
            "Effective Porosity (porf)":               self.v_porf,
            "Retardation Factor of PFAA-1 (ock(2))":   self.v_ret_trans1,
            "Retardation Factor of PFAA-2 (ock(4))":   self.v_ret_trans2,
            "Longitudinal Dispersivity (alphax (m))":  self.v_alpha_l,
            "Retardation Factor of Precursors-1 (ock(1))": self.v_ret_trans3,
            "Retardation Factor of Precursors-2 (ock(3))": self.v_ret_trans4,
        }
        # i and the multiplier rows are deliberately NOT in
        # REVERSE_MAP — i is captured in gwvelocity_inputs.txt by
        # the GW Calc; multipliers are baseline 1.  volfrac /
        # difflen / decay rates would write into sidecar txts that
        # the heterogeneity / transformation popups own.

        # ── Bake calibration multipliers physically into §7 cells ──
        # czero(*,n) multiplier rows in Step 4 only used to be
        # "virtually" applied inside generate_input_file (input.inp
        # reflected the multiplier, but §7 stayed at baseline).
        # The user wants §7 to actually show the multiplied values so
        # the Optimal run is visibly the same as what gets written.
        # After multiplying, reset the multiplier Mid to 1 so a
        # subsequent run doesn't compound (v×m×m instead of v×m).
        def _apply_calib_multipliers_to_s7():
            mid = getattr(self, "v_calib_mid", [])
            chk = getattr(self, "v_calib_chk", [])
            label_to_idx = {lbl: i for i, lbl in enumerate(_CALIB_PARAMS)}

            ROWS = [
                ("Multiplier to PFAA-1 Source Concentration in #7 (czero(2,n))",
                 self.v_src_pfaa1),
                ("Multiplier to PFAA-2 Source Concentration in #7 (czero(4,n))",
                 self.v_src_pfaa2),
                ("Multiplier to Precursor-1 Source Concentration in #7 (czero(1,n))",
                 self.v_src_pre1),
                ("Multiplier to Precursor-2 Source Concentration in #7 (czero(3,n))",
                 self.v_src_pre2),
            ]
            for label, src_list in ROWS:
                idx = label_to_idx.get(label)
                if idx is None or idx >= len(mid) or idx >= len(chk):
                    continue
                if not bool(chk[idx].get()):
                    continue
                try:
                    s = str(mid[idx].get()).replace(",", "").strip()
                    factor = float(s)
                except Exception:
                    continue
                if factor <= 0 or abs(factor - 1.0) < 1e-9:
                    continue
                # Multiply each of the 11 source-row cells in this
                # column by `factor`.  Re-format with commas + 2dp.
                for v in src_list:
                    try:
                        cur = float(str(v.get()).replace(",", "").strip())
                    except Exception:
                        continue
                    new = cur * factor
                    try: v.set(f"{new:,.2f}")
                    except Exception: pass
                # Mid stays at the optimal multiplier so the user
                # can see WHAT was applied.  Compounding is prevented
                # by the design: this helper is only called from
                # cali_1._on_done (right after best DDS values land
                # in Mid).  Load Optimal Data does NOT call this
                # helper anymore — it restores §7 source values
                # directly from optimal_model.txt v3.
        self._apply_calib_multipliers_to_s7 = _apply_calib_multipliers_to_s7

        def _push_calib_mids_to_inputs():
            """Walk the §calibration Mid column and write each cell's
            value back to the corresponding source StringVar.

            Special case: K and i are JOINT — vd = K × i is computed
            and written to v_darcy, matching the §3 GW Velocity
            Calculator's formula.  Writing K alone (as we used to)
            silently dropped the i factor and produced vd = K.

            Sets self._calib_pushing during the writes so source-cell
            traces don't fight us."""
            self._calib_pushing = True
            try:
                mid = getattr(self, "v_calib_mid", [])
                for i, lbl in enumerate(_CALIB_PARAMS):
                    if i >= len(mid):
                        break
                    target = REVERSE_MAP.get(lbl)
                    if target is None:
                        continue
                    try:
                        val = str(mid[i].get()).strip()
                        if val and val.lower() != "none":
                            target.set(val)
                    except Exception:
                        pass

                # vd = K × i  (joint).  Read K from the K row's
                # Mid, i from the i row's Mid, multiply, write to
                # v_darcy.  Falls back: if i Mid is blank/missing,
                # use 1; if K Mid is blank, leave v_darcy alone.
                try:
                    label_to_idx = {lbl: idx for idx, lbl in
                                    enumerate(_CALIB_PARAMS)}
                    k_idx = label_to_idx.get("Hydraulic Conductivity (k)")
                    i_idx = label_to_idx.get("Hydraulic Gradient (i)")
                    K_str = mid[k_idx].get().strip() if (k_idx is not None
                            and k_idx < len(mid)) else ""
                    i_str = mid[i_idx].get().strip() if (i_idx is not None
                            and i_idx < len(mid)) else ""
                    if K_str and K_str.lower() != "none":
                        K = float(K_str.replace(",", ""))
                        try:
                            i_val = float(i_str.replace(",", "")) \
                                    if i_str else 1.0
                        except Exception:
                            i_val = 1.0
                        vd = K * i_val
                        # Format with 4 decimals max, trailing 0s
                        # stripped — the §3 cell looks cleaner.
                        s = f"{vd:.4f}".rstrip("0").rstrip(".")
                        self.v_darcy.set(s if s else "0")
                except Exception:
                    pass
            finally:
                self._calib_pushing = False
            try: _refresh_all_mids()
            except Exception: pass
        self._push_calib_mids_to_inputs = _push_calib_mids_to_inputs

        # Wire trace_add on every distinct source StringVar (callable
        # sources can't be traced; they get re-read whenever
        # _refresh_calib_mids is called from elsewhere — Paste
        # Example, GW/Heterogeneity popup close, etc.)
        seen = set()
        for src in SRC_MAP.values():
            if src is None or callable(src) or id(src) in seen:
                continue
            seen.add(id(src))
            try:
                src.trace_add("write", _refresh_all_mids)
            except Exception:
                pass
        # _src_K reads v_darcy via callable, so v_darcy isn't picked
        # up by the loop above — register it explicitly so K Mid
        # updates the moment v_darcy changes (e.g. after GW Velocity
        # Calculator Apply, which writes the new Darcy velocity).
        try:
            self.v_darcy.trace_add("write", _refresh_all_mids)
        except Exception:
            pass

        # ── Initial baseline seed for rows without a §StringVar ──────
        # _refresh_one_mid is deferential and won't auto-fill these.
        # Seed once at build time so the user sees sensible defaults.
        BASELINE_MID = {
            "Hydraulic Gradient (i)":                                              "1",
            "Multiplier to PFAA-1 Source Concentration in #7 (czero(2,n))":        "1",
            "Multiplier to PFAA-2 Source Concentration in #7 (czero(4,n))":        "1",
            "Multiplier to Precursor-1 Source Concentration in #7 (czero(1,n))":   "1",
            "Multiplier to Precursor-2 Source Concentration in #7 (czero(3,n))":   "1",
        }
        for i, lbl in enumerate(_CALIB_PARAMS):
            if i >= len(self.v_calib_mid): break
            if lbl in BASELINE_MID:
                try: self.v_calib_mid[i].set(BASELINE_MID[lbl])
                except Exception: pass
        # Now fire the deferential refresh so live-tracking rows
        # (year, K, porf, retardation, alphax) pull their values.
        _refresh_all_mids()  # initial paint

        # ── K ↔ i linked checkboxes ──────────────────────────────────
        # The two go hand-in-hand for groundwater calibration: vd =
        # K × i, so the optimizer needs both perturbed together (or
        # neither).  Toggle one → toggle the other.  Re-entrancy
        # guard prevents infinite loop.
        try:
            k_idx = _CALIB_PARAMS.index("Hydraulic Conductivity (k)")
            i_idx = _CALIB_PARAMS.index("Hydraulic Gradient (i)")
        except ValueError:
            k_idx = i_idx = None

        if (k_idx is not None and i_idx is not None
                and k_idx < len(self.v_calib_chk)
                and i_idx < len(self.v_calib_chk)):
            self._ki_link_busy = False
            def _link_ki(driver_idx, follower_idx):
                def _cb(*_):
                    if self._ki_link_busy:
                        return
                    self._ki_link_busy = True
                    try:
                        self.v_calib_chk[follower_idx].set(
                            bool(self.v_calib_chk[driver_idx].get()))
                    finally:
                        self._ki_link_busy = False
                return _cb
            self.v_calib_chk[k_idx].trace_add("write",
                                              _link_ki(k_idx, i_idx))
            self.v_calib_chk[i_idx].trace_add("write",
                                              _link_ki(i_idx, k_idx))

        # ── Bottom action buttons — uniform size, centered 3-row block ─
        # Every action button now uses the same width / padx / pady so the
        # rows visually balance regardless of how long the caption is.
        # _RoundButton sizes its canvas to max(text_width, width*"0").
        BTN_W   = 22
        BTN_PX  = 12
        BTN_PY  = 12      # taller so 2-line captions don't overflow
        BTN_GAP = 8

        # All button captions wrap to 2 lines (or stay on 1) so every
        # button gets the SAME height regardless of caption length.
        actwrap = tk.Frame(outer, bg=BG_MAIN)
        actwrap.pack(pady=(16, 4))

        actrow1 = tk.Frame(actwrap, bg=BG_MAIN); actrow1.pack(pady=3)
        make_btn(actrow1, "1. Save\nCalibration Data",
                 "Save_Data_Calibration",
                 fg=FG_BTN_NAVY, font=FONT_BTN_CALIB,
                 padx=BTN_PX, pady=BTN_PY,
                 bg=BTN_FILL, width=BTN_W).pack(side="left", padx=BTN_GAP)
        make_btn(actrow1, "2. Run Machine\nBased Calibration",
                 "Run_Machine_Based_Calibration",
                 fg=FG_BTN_RED, font=FONT_BTN_CALIB,
                 padx=BTN_PX, pady=BTN_PY,
                 bg=BTN_FILL, width=BTN_W).pack(side="left", padx=BTN_GAP)
        make_btn(actrow1, "Go Back to\nMain Interface",
                 "Show_MainInterface",
                 fg=FG_BTN_NAVY, font=FONT_BTN_CALIB,
                 padx=BTN_PX, pady=BTN_PY,
                 bg=BTN_FILL, width=BTN_W).pack(side="left", padx=BTN_GAP)

        actrow2 = tk.Frame(actwrap, bg=BG_MAIN); actrow2.pack(pady=3)
        make_btn(actrow2, "3. See All\nCalibration Data",
                 "See_Calibration_Data",
                 fg=FG_BTN_NAVY, font=FONT_BTN_CALIB,
                 padx=BTN_PX, pady=BTN_PY,
                 bg=BTN_FILL, width=BTN_W).pack(side="left", padx=BTN_GAP)
        make_btn(actrow2, "4. Load\nOptimal Data",
                 "Load_Optimal_Data",
                 fg=FG_BTN_NAVY, font=FONT_BTN_CALIB,
                 padx=BTN_PX, pady=BTN_PY,
                 bg=BTN_FILL, width=BTN_W).pack(side="left", padx=BTN_GAP)
        make_btn(actrow2, "5. Run Optimal\nModel",
                 "Run_Optimal_Model",
                 fg=FG_BTN_NAVY, font=FONT_BTN_CALIB,
                 padx=BTN_PX, pady=BTN_PY,
                 bg=BTN_FILL, width=BTN_W).pack(side="left", padx=BTN_GAP)
        make_btn(actrow2, "Help\n ",
                 "Help_Calibration",
                 fg=FG_BTN_NAVY, font=FONT_BTN_CALIB,
                 padx=BTN_PX, pady=BTN_PY,
                 bg=BTN_FILL, width=BTN_W).pack(side="left", padx=BTN_GAP)

        actrow3 = tk.Frame(actwrap, bg=BG_MAIN); actrow3.pack(pady=3)
        make_btn(actrow3, "6. Save\nOptimal Model",
                 "Save_Optimal_Model",
                 fg=FG_BTN_NAVY, font=FONT_BTN_CALIB,
                 padx=BTN_PX, pady=BTN_PY,
                 bg=BTN_FILL, width=BTN_W).pack(padx=BTN_GAP)

        # Initial state — hide the Detailed-only widgets I just registered
        # if we're starting in Simple mode (default).  Mirrors the pattern
        # used elsewhere in _build_s5_transport / _build_s9_psb / etc.
        if self.active_sheet != "Detailed_2":
            for w in (det_lbl1, det_chk1, det_lbl2, det_chk2,
                      gap, note_pre):
                try: w.grid_remove()
                except Exception: pass
            # Also hide every Detailed-only Step 4 row I tagged above —
            # they're all in self._detailed_only_frames after the
            # det_lbl* / gap / note_pre entries.  Safest: hide every
            # entry whose master is `tbl` and that has _toggle_kind set.
            for w in self._detailed_only_frames:
                try:
                    if w.master is tbl:
                        w.grid_remove()
                except Exception: pass


    def _build_bottom_bar(self, parent):
        return None


# ENTRY POINT
if __name__ == "__main__":
    app = REMFluorApp()
    app.mainloop()
