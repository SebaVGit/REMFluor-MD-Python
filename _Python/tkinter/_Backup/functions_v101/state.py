"""
state.py — central in-memory data store for REMFluor-MD standalone app.

Replaces xlsm / xlwings. All functions read/write through this module.
Cell addresses use A1 notation (uppercase), same as XLSM_CELL_MAP in main.py.
"""
import json
import os
import re

# ---------------------------------------------------------------------------
# Cell address <-> UI variable name map  (mirrors XLSM_CELL_MAP in main.py)
# ---------------------------------------------------------------------------
CELL_MAP = {
    "B4":  "v_site",
    "E4":  "v_date",
    "E11": "v_x_size",
    "E12": "v_y_size",
    "E13": "v_z_size",
    "E15": "v_sw_width",
    "E16": "v_sw_thick",
    "E18": "v_yr_start",
    "E19": "v_yr_end",
    "M16": "v_run_time",
    "C22": "v_darcy",
    "G22": "v_porf",
    "K26": "v_lowk_media",
    "K27": "v_lowk_por",
    "K28": "v_lowk_tort",
    # E44 / E39 / E40 / G39 / G40 intentionally omitted:
    # these are computed by _on_pfaa_change() trace and must not be
    # overwritten by push() — the trace handles them automatically.
    "E38": "v_pfaa1",
    "G38": "v_pfaa2",
    "K38": "v_pfaa3",        # Detailed-only Precursor 1 species
    "M38": "v_pfaa4",        # Detailed-only Precursor 2 species
    # §5 transformation rate / yield (Detailed-only)
    "K41": "v_trans_rate_3",
    "M41": "v_trans_rate_4",
    "K42": "v_yield_factor_3",
    "M42": "v_yield_factor_4",
    "V4":  "v_alpha_l",
    "X4":  "v_alpha_t",
    "Z4":  "v_alpha_v",
    "D27": "v_src_rem_yr",
    "D28": "v_src_conc_red",
    "R22": "v_model_psb",
    # § 9 PSB — Freundlich "a"
    "V23": "v_psb_a_1",
    "X23": "v_psb_a_2",
    "Z23": "v_psb_a_3",      # Detailed-only
    "AB23": "v_psb_a_4",     # Detailed-only
    # § 9 PSB — Freundlich Kf + unit dropdown
    "U24": "v_psb_kf_unit",
    "V24": "v_psb_kf_1",
    "X24": "v_psb_kf_2",
    "Z24": "v_psb_kf_3",     # Detailed-only
    "AB24": "v_psb_kf_4",    # Detailed-only
    # § 9 PSB — PFAS molecular weight (mol-based units)
    "V25": "v_psb_mw_1",
    "X25": "v_psb_mw_2",
    "Z25": "v_psb_mw_3",     # Detailed-only
    "AB25": "v_psb_mw_4",    # Detailed-only
    # § 9 PSB — Converted Kf (auto-computed in app, but read-back from
    # state lets paste-example re-populate the BLACK output cells)
    "V26": "v_psb_kf_conv",
    "X26": "v_psb_kf_conv2",
    "Z26": "v_psb_kf_conv3", # Detailed-only
    "AB26": "v_psb_kf_conv4",# Detailed-only — *also* used by some Simple
                              #   inp_to_state writes for psb_yr; v_psb_yr
                              #   is mapped to AB28 instead so this stays
                              #   exclusive to v_psb_kf_conv4
    # § 9 PSB — geometry / install
    "AB28": "v_psb_yr",      # was Y75 — fixed to match the workbook
    "X74": "v_psb_dist",
    "Y82": "v_psb_width",
    "AA82": "v_psb_load",
    "AC82": "v_psb_cells",
    "Y74": "v_sample_yr",
    "V47": "v_see_every",
}
for _i in range(11):
    CELL_MAP[f"U{8+_i}"]  = f"v_src_years_{_i}"
    CELL_MAP[f"V{8+_i}"]  = f"v_src_pfaa1_{_i}"
    CELL_MAP[f"X{8+_i}"]  = f"v_src_pfaa2_{_i}"
    CELL_MAP[f"Z{8+_i}"]  = f"v_src_pre1_{_i}"   # Detailed-only Precursor 1
    CELL_MAP[f"AB{8+_i}"] = f"v_src_pre2_{_i}"   # Detailed-only Precursor 2
del _i
for _j in range(7):
    CELL_MAP[f"U{34+_j}"] = f"v_mw_names_{_j}"
    CELL_MAP[f"V{34+_j}"] = f"v_mw_conc_{_j}"
    CELL_MAP[f"X{34+_j}"] = f"v_mw_conc2_{_j}"
    CELL_MAP[f"AF{34+_j}"] = f"v_mw_dist_{_j}"
del _j

# Reverse: var_name -> addr
VAR_MAP = {v: k for k, v in CELL_MAP.items()}

# Cells that restore/clear scripts zero out
CLEAR_CELLS = [
    # Section 7 source rows (years + PFAA concentrations)
    *[f"U{r}" for r in range(8, 19)],   # src_years  ← was missing
    *[f"V{r}" for r in range(8, 19)],   # src_pfaa1 conc
    *[f"X{r}" for r in range(8, 19)],   # src_pfaa2 conc
    *[f"Z{r}" for r in range(8, 19)],
    *[f"AB{r}" for r in range(8, 19)],
    # Dispersivity (Section 6)
    "V4", "X4", "Z4",
    # Transport / porosity (Sections 3 & 4)
    "C22",                              # Darcy velocity  ← was missing
    "G22",                              # trans. porosity
    "K27", "K28",                       # low-k porosity / tortuosity
    # Model domain (Section 2)
    "E11", "E12", "E13", "E15", "E16", "E18", "E19",
    "M16",                              # run time
    # Site info (Section 1)
    "B4", "E4",
    # Misc computed / flags
    "AE25", "AH27", "E22", "E44",
    "K42", "M42", "K41", "M41", "AB26",
    # Section 9 PSB — every UI cell must be cleared on "Clear All Data"
    # (previously AB28 / AC82 / U24 / V25..AB26 were missing, which is
    # why "Year PSB Installed" appeared static across clear cycles).
    # NOTE: U24 (Kf unit dropdown) is INTENTIONALLY NOT cleared —
    # peach/pink dropdown cells (legend "Pull Down Menu") preserve
    # their selection across Clear All Data, matching the Excel macro.
    "V23", "X23", "Z23", "AB23",        # Freundlich "a"
    "V24", "X24", "Z24", "AB24",        # Freundlich Kf
    "V25", "X25", "Z25", "AB25",        # PFAS molecular weight
    "V26", "X26", "Z26", "AB26",        # Converted Kf (auto-computed,
                                        # but must be reset visually)
    "AB28",                             # Year PSB Installed (Detailed)
    "AB26",                             # Year PSB Installed (Simple,
                                        # legacy — same cell re-used
                                        # by Detailed for kf_conv4;
                                        # version-specific code chooses)
    "Y75",                              # legacy psb_yr (older sheet)
    "X74",                              # PSB distance from source
    "Y82",                              # Total Width of PSB
    "AA82",                             # PSB Loading 'fcac' (% display)
    "AC82",                             # # of cells in PSB in x-dir
    "AE25", "AH27", "AH28", "AH29",     # legacy PSB addresses
    # Section 8 source remediation  ← was missing
    "D27", "D28",
    # Section 10 field data  ← was missing
    "Y74",                              # sample_yr
    # Section 11 / misc
    "V47",
    "AH7", "AH8", "AH9", "R36", "AC1", "AH28",
    *[f"AF{r}" for r in range(34, 41)],
    *[f"U{r}" for r in range(34, 41)],
    *[f"V{r}" for r in range(34, 41)],
    *[f"X{r}" for r in range(34, 41)],
    # PFAS names: E38/G38/K38/M38 are the §5 pink dropdowns — NOT
    # cleared (Excel macro preserves dropdown selections).  Retardation
    # factors (E39..M40) ARE cleared.
    "E39", "E40", "G39", "G40", "K39", "K40", "M39", "M40",
]

INPUT_TXT_FILES = [
    "retardation_inputs.txt",
    "heterogeneity_inputs.txt",
    "transformation_inputs.txt",
    "numerical_inputs.txt",
    "calibration_inputs.txt",
    "cellsize_input.txt",
    "gwvelocity_inputs.txt",   # written by §3 GW Velocity Calculator
    "longevity_inputs.txt",    # written by §9 Simple CAC Longevity Tool
    "optimal_model.txt",       # written by §calibration Save Optimal Model
    "best_calib.json",         # v90: written by cali_1 with best_x + RMSLE
    "run_history.csv",         # written by cali_1 with all DDS evals
    "cali_debug.log",          # v91: per-iteration calibration failure log
    "dashboard_state.json",    # v89: written by run_model.py
    "mw_observations.json",    # v94: written by §10 .xlsx importer
]


# ---------------------------------------------------------------------------
# Address helpers
# ---------------------------------------------------------------------------
def _col_to_num(col: str) -> int:
    n = 0
    for c in col.upper():
        n = n * 26 + (ord(c) - 64)
    return n

def _num_to_col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def rc_to_addr(row: int, col: int) -> str:
    return f"{_num_to_col(col)}{row}"

def addr_to_rc(addr: str):
    m = re.match(r'^([A-Za-z]+)(\d+)$', addr.strip())
    if not m:
        raise ValueError(f"Bad cell address: {addr!r}")
    return int(m.group(2)), _col_to_num(m.group(1))


# ---------------------------------------------------------------------------
# Value parser
# ---------------------------------------------------------------------------
def _parse(raw):
    if isinstance(raw, bool):
        return raw
    s = str(raw).strip()
    if s.lower() in ('true',):
        return True
    if s.lower() in ('false',):
        return False
    if not s or s.lower() == 'none':
        return None
    try:
        return int(s)
    except ValueError:
        pass
    try:
        return float(s)
    except ValueError:
        pass
    return s


# ---------------------------------------------------------------------------
# AppState
# ---------------------------------------------------------------------------
class AppState:
    """
    Single source of truth for all cell values.
    Use get_state() to access the module-level singleton.
    """

    def __init__(self):
        self._cells: dict = {}
        self.work_dir: str = ""   # project directory (where .txt/.inp files live)
        # v100: read-only assets dir.  In dev = work_dir.  In a frozen
        # --onefile build = sys._MEIPASS (the temp unpack location for
        # PyInstaller --add-data files like Example/, docs/, Figures/).
        self.bundle_dir: str = ""

    # ── primitive access ─────────────────────────────────────────────────

    def get(self, addr: str, default=None):
        return self._cells.get(addr.upper(), default)

    def set(self, addr: str, value):
        self._cells[addr.upper()] = value

    def get_rc(self, row: int, col: int, default=None):
        return self._cells.get(rc_to_addr(row, col), default)

    def set_rc(self, row: int, col: int, value):
        self._cells[rc_to_addr(row, col)] = value

    # ── sync with tkinter UI ─────────────────────────────────────────────

    def snapshot(self, app):
        """Pull current UI widget values into state. Call before any function.

        Var-name resolution order (matters because §9 has scalar
        StringVars whose names end with _1.._4 and the §7 source rows
        have list-of-StringVars whose elements are addressed the same
        way; the simpler scalar form must win when both exist):
          1. ``getattr(app, var_name)`` — if it's a tk.Variable, use it.
          2. else if var_name matches ``base + "_<int>"`` and
             ``getattr(app, base)`` is a list, treat as list element.
        """
        import tkinter as tk
        for addr, var_name in CELL_MAP.items():
            # 1) scalar StringVar (handles v_psb_a_1..4, v_pfaa3/4, etc.)
            var = getattr(app, var_name, None)
            if isinstance(var, tk.Variable):
                self.set(addr, _parse(var.get()))
                continue
            # 2) list-element form (v_src_pfaa1_0..10, v_mw_names_0..6, ...)
            if re.search(r'_\d+$', var_name):
                base, idx_s = var_name.rsplit('_', 1)
                lst = getattr(app, base, None)
                if isinstance(lst, list):
                    idx = int(idx_s)
                    if idx < len(lst):
                        self.set(addr, _parse(lst[idx].get()))

        # Extra flags not in CELL_MAP
        ver = getattr(app, 'v_model_version', None)
        if ver:
            self.set("A8", 1 if "Simple" in ver.get() else 2)
        unit = getattr(app, 'v_units', None)
        if unit:
            self.set("AD1", 1 if unit.get() == "feet" else 2)
        het = getattr(app, 'v_het', None)
        if het:
            self.set("A1", {"High": 1, "Medium": 2, "Weak": 3}.get(het.get(), 2))

    def push(self, app):
        """Write state back to UI widgets.  See snapshot() for the
        scalar-first / list-fallback resolution rules."""
        import tkinter as tk
        for addr, var_name in CELL_MAP.items():
            val = self.get(addr)
            # 1) scalar StringVar (handles v_psb_a_1..4, v_pfaa3/4, etc.)
            var = getattr(app, var_name, None)
            if isinstance(var, tk.Variable):
                var.set("" if val is None else str(val))
                continue
            # 2) list-element form (v_src_pfaa1_0..10, v_mw_names_0..6, ...)
            if re.search(r'_\d+$', var_name):
                base, idx_s = var_name.rsplit('_', 1)
                lst = getattr(app, base, None)
                if isinstance(lst, list):
                    idx = int(idx_s)
                    if idx < len(lst):
                        lst[idx].set("" if val is None else str(val))


        # Push version / units flags
        a8 = self.get("A8")
        if a8 is not None:
            ver = getattr(app, 'v_model_version', None)
            if ver:
                ver.set("Simple Version" if a8 == 1 else "Detailed Version")
        ad1 = self.get("AD1")
        if ad1 is not None:
            unit = getattr(app, 'v_units', None)
            if unit:
                unit.set("feet" if ad1 == 1 else "meters")

    # ── clear helpers ─────────────────────────────────────────────────────

    def clear_restore_cells(self):
        """Zero cells that clear_for_restore clears in the xlsm."""
        for addr in CLEAR_CELLS:
            self.set(addr, None)
        self.set("R22", False)

    # ── persistence ───────────────────────────────────────────────────────

    def save_json(self, path: str):
        with open(path, 'w') as f:
            json.dump(self._cells, f, indent=2, default=str)

    def load_json(self, path: str):
        with open(path) as f:
            raw = json.load(f)
        self._cells = {k.upper(): v for k, v in raw.items()}


# Module-level singleton
_state = AppState()

def get_state() -> AppState:
    return _state
