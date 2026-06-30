"""
sidecars.py — dedicated .txt sidecars for EXACT save/load of the
Section 6 dispersivity inputs and Section 9 PSB (Permeable Sorptive
Barrier) inputs.

Why these exist (instead of relying on the input.inp round-trip):
  * §6 dispersivity — when the user picks "Enter Your Own Value Below"
    the custom alphax/alphay/alphaz were getting clobbered on Load Data
    because the §6 heterogeneity trace re-derives them from a preset.
  * §9 PSB Freundlich — input.inp only stores the CONVERTED Kf folded
    into fcackf (= converted Kf × loading).  Recovering the user's
    ORIGINAL Kf + unit dropdown from that value is lossy and was showing
    the wrong Freundlich term / wrong units after Load Data.

Each sidecar stores the raw UI values verbatim so a Save → Load round
trip is exact.  The Converted-Kf cells are NOT stored; they are
auto-recomputed from the raw a / Kf / MW / unit after load.
"""
import os

DISPERSIVITY_FILE = "dispersivity_inputs.txt"
PSB_FILE = "psb_inputs.txt"

# Files this module owns — used by clean_sidecars().
OWNED_FILES = (DISPERSIVITY_FILE, PSB_FILE)


# ---------------------------------------------------------------------------
# small tk-var helpers
# ---------------------------------------------------------------------------
def _g(app, name):
    """Read a tk.Variable attribute as a string ('' if missing)."""
    v = getattr(app, name, None)
    if v is None:
        return ""
    try:
        return v.get()
    except Exception:
        return ""


def _s(app, name, val):
    """Set a tk.Variable attribute (no-op if missing)."""
    v = getattr(app, name, None)
    if v is None:
        return
    try:
        v.set("" if val is None else str(val))
    except Exception:
        pass


def _parse_kv(path):
    """Parse a 'Label:,value' sidecar into {label: value}."""
    d = {}
    if not os.path.exists(path):
        return d
    try:
        with open(path, encoding="utf-8") as fh:
            for ln in fh:
                ln = ln.strip()
                if not ln or ln.startswith("=") or ":" not in ln:
                    continue
                key, _, val = ln.partition(":")
                d[key.strip()] = val.strip().lstrip(",").strip()
    except Exception as exc:
        print(f"[sidecars] parse failed {os.path.basename(path)}: {exc}")
    return d


# ---------------------------------------------------------------------------
# Section 6 — dispersivity
# ---------------------------------------------------------------------------
def write_dispersivity(app, folder):
    lines = [
        "REMFluor Dispersivity Inputs (Section 6)",
        "=" * 50,
        f"Heterogeneity Mode:,{_g(app, 'v_het')}",
        f"Units:,{_g(app, 'v_units')}",
        f"Longitudinal alphax:,{_g(app, 'v_alpha_l')}",
        f"Transverse alphay:,{_g(app, 'v_alpha_t')}",
        f"Vertical alphaz:,{_g(app, 'v_alpha_v')}",
        # §5 General Molecular Diffusion Coefficient (m²/s) — stored here
        # for an EXACT round-trip.  The input.inp path rounds it to 3
        # decimals in m²/yr and the §5 PFAA trace re-derives it from the
        # species table, so a user override was being lost on Load.
        f"Molecular Diffusion m2s:,{_g(app, 'v_mol_diff')}",
    ]
    with open(os.path.join(folder, DISPERSIVITY_FILE), "w",
              encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")


def read_dispersivity(app, folder):
    """Apply dispersivity sidecar to the UI.  Returns True if applied.

    Caller MUST set app._disp_loading = True around this call so the §6
    heterogeneity trace does not re-derive / clear the alpha cells."""
    d = _parse_kv(os.path.join(folder, DISPERSIVITY_FILE))
    if not d:
        return False
    mode = d.get("Heterogeneity Mode")
    if mode:
        _s(app, "v_het", mode)
    _s(app, "v_alpha_l", d.get("Longitudinal alphax", ""))
    _s(app, "v_alpha_t", d.get("Transverse alphay", ""))
    _s(app, "v_alpha_v", d.get("Vertical alphaz", ""))
    # §5 General Molecular Diffusion Coefficient — restore the exact saved
    # value and flag it as a manual override so the §5 PFAA trace
    # (_on_pfaa_change → _update_mol_diff) keeps it instead of replacing it
    # with the species default.  last_species is pinned to the loaded
    # PFAA-1 so the value sticks until the user actually changes species.
    md = d.get("Molecular Diffusion m2s")
    if md not in (None, ""):
        _s(app, "v_mol_diff", md)
        try:
            app._mol_diff_user_edited = True
            pf = getattr(app, "v_pfaa1", None)
            app._mol_diff_last_species = pf.get() if pf is not None else None
        except Exception:
            pass
    return True


# ---------------------------------------------------------------------------
# Section 9 — PSB
# ---------------------------------------------------------------------------
# (label, attribute).  v_model_psb (BooleanVar) and v_psb_kf_unit are
# handled specially in read_psb; the rest are plain StringVars.
_PSB_FIELDS = [
    ("Kf Unit",              "v_psb_kf_unit"),
    ("Freundlich a 1",       "v_psb_a_1"),
    ("Freundlich a 2",       "v_psb_a_2"),
    ("Freundlich a 3",       "v_psb_a_3"),
    ("Freundlich a 4",       "v_psb_a_4"),
    ("Freundlich Kf 1",      "v_psb_kf_1"),
    ("Freundlich Kf 2",      "v_psb_kf_2"),
    ("Freundlich Kf 3",      "v_psb_kf_3"),
    ("Freundlich Kf 4",      "v_psb_kf_4"),
    ("PFAS MW 1",            "v_psb_mw_1"),
    ("PFAS MW 2",            "v_psb_mw_2"),
    ("PFAS MW 3",            "v_psb_mw_3"),
    ("PFAS MW 4",            "v_psb_mw_4"),
    ("Year PSB Installed",   "v_psb_yr"),
    ("Cells in X Direction", "v_psb_cells"),
    ("Distance from Source", "v_psb_dist"),
    ("Total Width of PSB",   "v_psb_width"),
    ("PSB Loading pct",      "v_psb_load"),
]


def write_psb(app, folder):
    lines = [
        "REMFluor PSB Inputs (Section 9)",
        "=" * 50,
        "Freundlich Kf 1-4 are RAW values in the unit named by 'Kf Unit'.",
        "Converted Kf is auto-recomputed on load as (ug/kg)(ug/L)^(-a).",
        f"PSB Enabled:,{_g(app, 'v_model_psb')}",
    ]
    for label, attr in _PSB_FIELDS:
        lines.append(f"{label}:,{_g(app, attr)}")
    with open(os.path.join(folder, PSB_FILE), "w", encoding="utf-8") as fh:
        fh.write("\n".join(lines) + "\n")


def read_psb(app, folder):
    """Apply PSB sidecar to the UI.  Returns True if applied.

    Sets the Kf unit FIRST so the Converted-Kf auto-formula uses the
    correct unit prefix, then the raw a / Kf / MW / geometry cells.  The
    converted-Kf trace recomputes V26/X26/Z26/AB26 from these."""
    d = _parse_kv(os.path.join(folder, PSB_FILE))
    if not d:
        return False
    # 1) unit dropdown first (drives the conversion factor)
    if d.get("Kf Unit"):
        _s(app, "v_psb_kf_unit", d["Kf Unit"])
    # 2) PSB enable checkbox (BooleanVar)
    en = d.get("PSB Enabled", "")
    pv = getattr(app, "v_model_psb", None)
    if pv is not None:
        try:
            pv.set(str(en).strip().lower() in ("1", "true", "yes", "on"))
        except Exception:
            pass
    # 3) remaining raw cells (skip the unit — already set)
    for label, attr in _PSB_FIELDS:
        if attr == "v_psb_kf_unit":
            continue
        if label in d:
            _s(app, attr, d[label])
    return True


# ---------------------------------------------------------------------------
# cleanup
# ---------------------------------------------------------------------------
def clean_sidecars(folder):
    """Delete this module's sidecars from a folder so stale values from a
    previous model don't leak into a freshly-loaded one."""
    for f in OWNED_FILES:
        p = os.path.join(folder, f)
        try:
            if os.path.exists(p):
                os.remove(p)
        except Exception as exc:
            print(f"[sidecars] could not remove {f}: {exc}")
