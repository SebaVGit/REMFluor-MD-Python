"""
cali_1.py — standalone Python port of Source_Py/cali_1.py.

Drives the "Run Machine Based Calibration" button.  Replaces the legacy
dist/cali_1.exe pipeline (which depended on xlwings + an Excel template
+ a pandas/sklearn stack) with a pure-Python implementation that:

  1.  Pulls the checked parameter ranges + monitoring-well weights
      directly from the app's StringVars (no Excel template required).
  2.  Runs the DDS (Dynamically Dimensioned Search) optimizer over
      those ranges.  DDS class is a verbatim port of the one in
      Source_Py/cali_1.py — it's pure numpy.
  3.  For each DDS evaluation:
        - Override the relevant state cells with the trial values
        - Rebuild input.inp via generate_input_file.build_inp_data
        - Run remfluor_v8a.exe
        - Parse obs_well*.out and compute RMSLE vs the observed
          monitoring-well concentrations
  4.  When done, pushes the best parameter set back into the app and
      writes run_history.csv next to input.inp.

Required: numpy (everything else is stdlib).  If numpy isn't present
the dispatcher catches the ImportError and tells the user to install
it.
"""
from __future__ import annotations

import csv
import glob
import math
import os
import re
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox
from typing import Callable, Optional, List, Dict, Any

from . import generate_input_file
from .state import get_state
from .inp_parser import parse_input_inp


# ─────────────────────────────────────────────────────────────────────
# Parameter map — calibration row label → state cell address (or
# a callable taking (state, value) and applying any custom mapping).
# Rows whose semantic mapping isn't trivial (e.g. retardation factors,
# decay rates that live in retardation_inputs.txt / transformation_
# inputs.txt) are listed but resolve to a no-op (None) for now — the
# DDS still perturbs them in-memory but the perturbation has no effect
# on input.inp until those sidecars get a Python writer.  This is a
# pragmatic V1; expand the map as new sidecar writers land.
# ─────────────────────────────────────────────────────────────────────
def _set_vd(state, val):
    state.set("C22", val)
    state.set("E22", val)


def _mul_source_col(state, addr_pattern, factor):
    """Multiply the 11 source-column cells (V8..V18 etc.) by `factor`."""
    for i in range(11):
        addr = addr_pattern.format(row=8 + i)
        cur = state.get(addr)
        try:
            cur_f = float(str(cur).replace(",", "").strip())
        except (TypeError, ValueError, AttributeError):
            cur_f = 0.0
        state.set(addr, cur_f * factor)


# K and i are calibrated JOINTLY: vd = K × i (matches the §3 GW
# Velocity Calculator's formula).  Their PARAM_MAP entries are
# sentinel "K" / "i" strings rather than direct setters — the eval
# function below picks them out of the DDS x-vector, multiplies, and
# writes vd to v_darcy in a single shot.  This avoids the bug where
# the K setter overwrote v_darcy with K-only and ignored i.
PARAM_MAP: Dict[str, Any] = {
    # label                                                  setter / sentinel
    "Source Start Year (nt)":
        lambda s, v: s.set("E18", int(round(v))),
    "Hydraulic Conductivity (k)":     "K",   # joint with i
    "Hydraulic Gradient (i)":         "i",   # joint with K
    "Effective Porosity (porf)":
        lambda s, v: s.set("G22", v),
    "Transmissive Fraction of Model (volfrac)":
        # volfrac is in heterogeneity_inputs.txt; sidecar writer not
        # yet implemented — DDS perturbs but no-op on input.inp.
        lambda s, v: None,
    "Average Diffusion Length (difflen)":
        lambda s, v: None,
    "Retardation Factor of PFAA-1 (ock(2))":
        lambda s, v: None,
    "Retardation Factor of PFAA-2 (ock(4))":
        lambda s, v: None,
    "Longitudinal Dispersivity (alphax (m))":
        lambda s, v: s.set("V4", v),
    "Multiplier to PFAA-1 Source Concentration in #7 (czero(2,n))":
        lambda s, v: _mul_source_col(s, "V{row}", v),
    "Multiplier to PFAA-2 Source Concentration in #7 (czero(4,n))":
        lambda s, v: _mul_source_col(s, "X{row}", v),
    "First order decay rate coefficient for Precursors-1 (decayf(1))":
        lambda s, v: None,
    "First order decay rate coefficient for Precursors-2 (decayf(3))":
        lambda s, v: None,
    "Retardation Factor of Precursors-1 (ock(1))":
        lambda s, v: None,
    "Retardation Factor of Precursors-2 (ock(3))":
        lambda s, v: None,
    "Multiplier to Precursor-1 Source Concentration in #7 (czero(1,n))":
        lambda s, v: _mul_source_col(s, "Z{row}", v),
    "Multiplier to Precursor-2 Source Concentration in #7 (czero(3,n))":
        lambda s, v: _mul_source_col(s, "AB{row}", v),
}


# ─────────────────────────────────────────────────────────────────────
# DDS optimizer — verbatim port of Source_Py/cali_1.py:DDS, simplified
# to drop matplotlib / verbose plotting.  Maximizes the objective; we
# pass in -RMSLE so that maximizing the objective minimizes the error.
# ─────────────────────────────────────────────────────────────────────
class DDS:
    def __init__(self, function, x_min, x_max, max_evals,
                 r=0.2, x_initial=None, seed=None,
                 progress: Optional[Callable[[int, int, float, float], None]] = None,
                 stop_flag: Optional[Dict[str, bool]] = None):
        import numpy as np
        if seed is not None:
            np.random.seed(seed)
        self._np = np
        self.f = function
        self.x_min = np.asarray(x_min, dtype=float)
        self.x_max = np.asarray(x_max, dtype=float)
        if self.x_min.shape != self.x_max.shape:
            raise ValueError("x_min and x_max must have the same shape")
        self.d = self.x_min.size
        if max_evals < 1:
            raise ValueError("max_evals must be >= 1")
        self.max_evals = int(max_evals)
        self.r = float(r)
        self.progress = progress
        self.stop_flag = stop_flag or {"stop": False}

        if x_initial is not None:
            x0 = np.asarray(x_initial, dtype=float)
            self.best = np.minimum(np.maximum(x0, self.x_min), self.x_max)
        else:
            self.best = (np.random.rand(self.d) * (self.x_max - self.x_min)
                         + self.x_min)
        self.best_cost = float(self.f(self.best))
        self.x_history = [self.best.copy()]
        self.f_history = [self.best_cost]
        self.best_costs = [self.best_cost]
        self.x_range = self.x_max - self.x_min
        self._zero = self.x_range == 0.0
        self.x_range[self._zero] = 1.0

    def _reflect(self, x):
        np = self._np
        below = x < self.x_min
        x[below] = 2 * self.x_min[below] - x[below]
        over = below & (x > self.x_max)
        x[over] = self.x_min[over]
        above = x > self.x_max
        x[above] = 2 * self.x_max[above] - x[above]
        under = above & (x < self.x_min)
        x[under] = self.x_max[under]
        np.clip(x, self.x_min, self.x_max, out=x)
        return x

    def run(self):
        np = self._np
        if self.progress:
            self.progress(1, self.max_evals, self.best_cost, self.best_cost)
        for i in range(2, self.max_evals + 1):
            if self.stop_flag.get("stop"):
                break
            p = 1.0 - (math.log(i) / math.log(max(self.max_evals, 2)))
            p = max(0.0, min(1.0, p))
            choose = (np.random.rand(self.d) < p)
            if not np.any(choose):
                choose[np.random.randint(0, self.d)] = True
            k = int(np.count_nonzero(choose))
            curr = self.best.copy()
            curr[choose] = (curr[choose]
                            + np.random.normal(size=k)
                              * self.r * self.x_range[choose])
            curr = self._reflect(curr)
            curr_cost = float(self.f(curr))
            self.x_history.append(curr.copy())
            self.f_history.append(curr_cost)
            if curr_cost > self.best_cost:
                self.best = curr
                self.best_cost = curr_cost
            self.best_costs.append(self.best_cost)
            if self.progress:
                self.progress(i, self.max_evals, curr_cost, self.best_cost)
        return {
            "best_x":     self.best.copy(),
            "best_cost":  self.best_cost,
            "x_history":  np.array(self.x_history),
            "f_history":  np.array(self.f_history),
            "best_costs": np.array(self.best_costs),
        }


# ─────────────────────────────────────────────────────────────────────
# RMSLE + observation parsing
# ─────────────────────────────────────────────────────────────────────
def _rmsle(obs, pred):
    import numpy as np
    obs = np.asarray(obs, dtype=float)
    pred = np.asarray(pred, dtype=float)
    if obs.size == 0:
        return float("inf")
    obs = np.clip(obs, 0, None)
    pred = np.clip(pred, 0, None)
    diff = np.log1p(pred) - np.log1p(obs)
    return float(np.sqrt(np.nanmean(diff ** 2)))


def _read_obs_wells(work_dir: str) -> Dict[str, List[List[float]]]:
    """Parse obs_well1.out..obs_wellN.out.  Each file is a CSV with
    columns Time(yr), C1well, C2well, C3well, C4well.  Returns
    {filename: [[time, c1, c2, c3, c4], ...]}."""
    out = {}
    for path in sorted(glob.glob(os.path.join(work_dir, "obs_well*.out"))):
        rows: List[List[float]] = []
        try:
            with open(path, "r") as fp:
                for ln in fp:
                    parts = [p.strip() for p in ln.split(",") if p.strip()]
                    if not parts: continue
                    try:
                        rows.append([float(x) for x in parts[:5]])
                    except ValueError:
                        continue   # header line
        except Exception:
            continue
        if rows:
            out[os.path.basename(path)] = rows
    return out


def _objective_for_app(app, work_dir, exe_path) -> Callable:
    """Build a closure that maps a DDS x-vector to RMSLE."""
    state = get_state()
    # Snapshot UI once so we have a known-good baseline; we'll reset
    # to this between evaluations so each trial only sees DDS overrides.
    state.snapshot(app)
    baseline = dict(state._cells)

    # Observed concentrations: pull from §10 monitoring well grid.
    # User typed values into v_mw_conc / v_mw_conc2 with a sample year
    # (Y74).  We compare each well's loaded concentration to the
    # solver's last available time-step concentration at that well's
    # well-row obs_well<N>.out.
    obs_pfaa1 = []
    obs_pfaa2 = []
    for v in getattr(app, "v_mw_conc", [])[:7]:
        try:
            obs_pfaa1.append(float(str(v.get()).replace(",", "").strip()))
        except (ValueError, TypeError, AttributeError):
            obs_pfaa1.append(float("nan"))
    for v in getattr(app, "v_mw_conc2", [])[:7]:
        try:
            obs_pfaa2.append(float(str(v.get()).replace(",", "").strip()))
        except (ValueError, TypeError, AttributeError):
            obs_pfaa2.append(float("nan"))

    # Sample year drives which row of obs_well*.out to compare against.
    try:
        sample_year = int(float(str(app.v_sample_yr.get()).strip()))
    except (ValueError, TypeError, AttributeError):
        sample_year = 0
    try:
        start_year = int(float(state.get("E18", 1977)))
    except (ValueError, TypeError):
        start_year = 1977
    target_t = max(0, sample_year - start_year)

    # Parameter labels (kept in DDS x-vector order)
    labels = _objective_for_app.labels  # set by run()

    def _eval(x):
        import numpy as np
        # Reset state to baseline then overlay DDS values
        state._cells.clear()
        state._cells.update(baseline)

        # First pass: split DDS values into "joint" sentinels (K, i)
        # vs regular setters.  K and i need to be combined into vd =
        # K × i in a single atomic step so neither overwrites the
        # other.
        K_val = i_val = None
        for lbl, val in zip(labels, x):
            entry = PARAM_MAP.get(lbl)
            if entry is None:
                continue
            if entry == "K":
                K_val = float(val); continue
            if entry == "i":
                i_val = float(val); continue
            if callable(entry):
                try: entry(state, float(val))
                except Exception: pass

        # If K and/or i is being calibrated, recompute vd.  When only
        # one is checked, fall back to the baseline value of the
        # other so the math still makes physical sense (vd = K × i).
        if K_val is not None or i_val is not None:
            # Baseline K and i — read from current state (fall back
            # to v_darcy if no separate K cell exists).
            base_vd = 0.0
            try:
                base_vd = float(str(baseline.get("C22", "0"))
                                .replace(",", "").strip())
            except Exception:
                pass
            K = K_val if K_val is not None else (base_vd or 0.0)
            i = i_val if i_val is not None else 1.0
            vd = K * i
            _set_vd(state, vd)

        # Build + write input.inp
        try:
            data = generate_input_file.build_inp_data(state)
        except Exception as exc:
            print(f"[cali] build_inp_data failed: {exc}")
            return -1e9

        proj = os.path.dirname(exe_path)
        template_path = os.path.join(proj, "template.inp")
        out_path      = os.path.join(proj, "input.inp")
        if not os.path.exists(template_path):
            return -1e9
        try:
            with open(template_path) as fp:
                tmpl = fp.read()
            src_lines = []
            for t, c1, c2, c3, c4 in zip(
                    data['times'], data['concs1'], data['concs2'],
                    data['concs3'], data['concs4']):
                src_lines.append(f"{t or 0}, {c1 or 0}, {c2 or 0}, "
                                 f"{c3 or 0}, {c4 or 0}")
            filled = tmpl.format(
                source_zone_data="\n".join(src_lines),
                observation_wells="\n".join(data['wells']),
                **{k: v for k, v in data.items()
                   if k not in ('times','concs1','concs2',
                                'concs3','concs4','wells')},
            )
            with open(out_path, "w") as fp:
                fp.write(filled)
        except Exception as exc:
            print(f"[cali] template fill failed: {exc}")
            return -1e9

        # Run solver
        try:
            cmd = '"{}" < "input.inp" > "output.out"'.format(exe_path)
            subprocess.run(cmd, shell=True, cwd=proj, timeout=600,
                           capture_output=True)
        except Exception as exc:
            print(f"[cali] solver failed: {exc}")
            return -1e9

        # Compare obs_well*.out @ target_t vs typed observations
        wells = _read_obs_wells(proj)
        if not wells:
            return -1e9
        sims_pfaa1, obss_pfaa1 = [], []
        sims_pfaa2, obss_pfaa2 = [], []
        well_names = sorted(wells.keys())
        for i, name in enumerate(well_names[:7]):
            rows = wells[name]
            # nearest-time row to target_t
            best_row = min(rows, key=lambda r: abs(r[0] - target_t))
            sim_c2 = best_row[2]   # PFAA 1 → C2well
            sim_c4 = best_row[4]   # PFAA 2 → C4well
            if i < len(obs_pfaa1) and not np.isnan(obs_pfaa1[i]):
                sims_pfaa1.append(sim_c2)
                obss_pfaa1.append(obs_pfaa1[i])
            if i < len(obs_pfaa2) and not np.isnan(obs_pfaa2[i]):
                sims_pfaa2.append(sim_c4)
                obss_pfaa2.append(obs_pfaa2[i])

        all_sim = sims_pfaa1 + sims_pfaa2
        all_obs = obss_pfaa1 + obss_pfaa2
        if not all_obs:
            return -1e9
        err = _rmsle(all_obs, all_sim)
        return -err   # DDS maximizes; we want to minimize RMSLE

    return _eval


# ─────────────────────────────────────────────────────────────────────
# Public entry point
# ─────────────────────────────────────────────────────────────────────
def run(app, parent=None) -> bool:
    """Run the DDS calibration loop.  Returns True on success."""
    try:
        import numpy as np   # noqa
    except ImportError:
        messagebox.showerror(
            "Run Machine Based Calibration",
            "numpy is required for the calibration optimizer.\n\n"
            "Install with:\n    pip install numpy\n\n"
            "Then click Run Machine Based Calibration again.")
        return False

    # Gather checked parameter ranges -----------------------------------
    chk  = getattr(app, "v_calib_chk",  [])
    lo   = getattr(app, "v_calib_low",  [])
    hi   = getattr(app, "v_calib_high", [])

    # v88: introspect the model configuration to filter out checked
    # rows that don't apply.  Precursor parameters only matter when
    # the corresponding Precursor name is non-empty (ipre=1 logic in
    # generate_input_file); PFAA-2 / Precursor-2 parameters only
    # matter when PFAA-2 / Precursor-2 names are non-empty (ncomp=2).
    # Without this filter, the user could leave the legacy default
    # "Precursor-1 multiplier" checkbox ticked and end up with a
    # bogus parameter in run_history.csv even on PFAA-only configs.
    def _has(addr_var):
        v = getattr(app, addr_var, None)
        if v is None:
            return False
        try:
            s = str(v.get()).strip()
        except Exception:
            return False
        return s != "" and s.lower() != "none"

    has_pfaa1   = _has("v_pfaa1") or _has("v_pfaa2")  # PFAA-1 always present
    has_pfaa2   = _has("v_pfaa2") and str(app.v_pfaa2.get()).strip().lower() != "none"
    has_pre1    = _has("v_pfaa3")   # K38 — Precursor 1 name (StringVar v_pfaa3)
    has_pre2    = _has("v_pfaa4")   # M38 — Precursor 2 name (StringVar v_pfaa4)
    is_simple   = (getattr(app, "active_sheet", "Detailed_2") == "Simple")
    if is_simple:
        # Simple mode has no precursors at all
        has_pre1 = False
        has_pre2 = False

    def _row_applies(label):
        """Return False when a checked row references a disabled component."""
        L = label.lower()
        if "precursor-2" in L or "precursors-2" in L or "(czero(3" in L \
           or "(decayf(3))" in L or "(ock(3))" in L:
            return has_pre2
        if "precursor-1" in L or "precursors-1" in L or "(czero(1" in L \
           or "(decayf(1))" in L or "(ock(1))" in L:
            return has_pre1
        if "pfaa-2" in L or "(czero(4" in L or "(ock(4))" in L:
            return has_pfaa2
        return True

    from main import _CALIB_PARAMS  # type: ignore  (set by main.py)
    labels: List[str] = []
    x_min, x_max = [], []
    skipped_rows = []
    for i, label in enumerate(_CALIB_PARAMS):
        if i >= len(chk): break
        if not bool(chk[i].get()):
            continue
        if not _row_applies(label):
            skipped_rows.append(label)
            continue
        try:
            lo_v = float(str(lo[i].get()).replace(",", "").strip())
            hi_v = float(str(hi[i].get()).replace(",", "").strip())
        except (ValueError, TypeError):
            continue
        if hi_v <= lo_v:
            continue
        labels.append(label)
        x_min.append(lo_v)
        x_max.append(hi_v)

    if skipped_rows:
        msg = ("These checked rows were skipped because the "
               "corresponding component isn't configured:\n\n"
               + "\n".join(f"  • {r}" for r in skipped_rows)
               + "\n\nFill in the component name in §5 (precursors) "
                 "or add a 2nd PFAA in §1 to include them.")
        try:
            messagebox.showwarning("Calibration — rows skipped", msg)
        except Exception:
            print(f"[cali] skipped rows: {skipped_rows}")

    if not labels:
        messagebox.showinfo(
            "Run Machine Based Calibration",
            "No applicable calibration parameters checked.\n\n"
            "Tick at least one row in §Step 4 and set a Lowest /\n"
            "Highest Likely Value, then try again.")
        return False

    try:
        max_evals = max(1, int(float(app.v_n_iter.get())))
    except Exception:
        max_evals = 50

    # Resolve solver path
    here = os.path.dirname(os.path.abspath(__file__))
    project = os.path.abspath(os.path.join(here, "..", "..", ".."))
    exe_path = os.path.join(project, "remfluor_v8a.exe")
    if not os.path.exists(exe_path):
        messagebox.showerror(
            "Run Machine Based Calibration",
            f"Model executable not found:\n{exe_path}")
        return False

    # Progress popup ----------------------------------------------------
    # Big + centered so the user can read iteration / RMSLE progress
    # at a glance.  Title bar + 26-pt label + 14-pt detail; minsize
    # 900x520 — sized to look like the Run Model runtime clock.
    root = tk.Toplevel(parent or app)
    root.title("REMFluor Calibration in Progress")
    root.configure(bg="#F0F0F0")
    try:
        root.transient(parent or app)
    except Exception:
        pass

    title = tk.Label(root, text="Calibration in Progress",
                     font=("Arial", 20, "bold"),
                     bg="#F0F0F0", fg="#222",
                     padx=24, pady=14, justify="center",
                     wraplength=1100)
    title.pack(padx=36, pady=(28, 6))

    label = tk.Label(root, text="Starting...",
                     font=("Arial", 22, "bold"),
                     bg="#F0F0F0", fg="#222",
                     padx=24, pady=8, justify="center",
                     wraplength=1100)
    label.pack(padx=36, pady=(0, 8))

    # Wider wraplength on the detail label so the multi-line summary
    # printed at the end ("3 parameter(s) optimized over 1 evaluations.
    # Best values pushed into the Mid-Range column. Full history →
    # run_history.csv") fits without breaking words mid-line.
    detail = tk.Label(root, text="",
                      font=("Arial", 13),
                      bg="#F0F0F0", fg="#444",
                      justify="center",
                      wraplength=1100)
    detail.pack(padx=36, pady=(0, 22))

    stop_flag = {"stop": False}
    btn_row = tk.Frame(root, bg="#F0F0F0"); btn_row.pack(pady=(0, 32))
    def _cancel():
        stop_flag["stop"] = True
        label.config(text="Cancelling...")
    tk.Button(btn_row, text="Cancel", width=16,
              font=("Arial", 14),
              command=_cancel).pack(side="left", padx=14)

    # Bigger popup so the "Best RMSLE: …" line + 3-line summary +
    # Close button all fit visibly without resizing.  Was 900x520.
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 64, 1200)
    h = max(root.winfo_reqheight() + 80, 640)
    try:
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        x = max(0, (sw - w) // 2); y = max(0, (sh - h) // 2 - 40)
        root.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        root.geometry(f"{w}x{h}")
    root.minsize(w, h)
    root.resizable(True, True)
    try:
        root.lift(); root.focus_force()
        root.attributes("-topmost", True)
        root.after(800, lambda: root.attributes("-topmost", False))
    except Exception:
        pass

    # Track elapsed time so we can drive the §calibration "Estim. Time
    # Remaining" cell with a real ETA.
    run_start = time.time()

    def _progress(i, total, curr, best):
        rmse_curr = -curr if curr > -1e8 else float("inf")
        rmse_best = -best if best > -1e8 else float("inf")
        elapsed = time.time() - run_start
        try:
            root.after(0, lambda: label.config(
                text=f"Iteration {i}/{total}"))
            root.after(0, lambda: detail.config(
                text=f"Current RMSLE: {rmse_curr:.4g}\n"
                     f"Best RMSLE so far: {rmse_best:.4g}"))
        except Exception:
            pass
        # Push the ETA into the calibration panel's Estim. Time
        # Remaining cell (the live tracker added in main.py).  Wrapped
        # in try/except so a missing helper doesn't break the run.
        cb = getattr(app, "_cali_progress_update", None)
        if callable(cb):
            try:
                root.after(0, lambda: cb(i, total, elapsed))
            except Exception:
                pass

    # Worker thread runs the DDS so the popup stays responsive
    result_holder = {"res": None, "err": None}
    _objective_for_app.labels = labels  # smuggle into closure

    def _worker():
        # Flip the app-level "calibration in progress" flag so the
        # §calibration Estim. Time Remaining cell switches from the
        # idle estimate to the live ETA driven by _progress() above.
        try: setattr(app, "_cali_running", True)
        except Exception: pass
        try:
            obj = _objective_for_app(app, project, exe_path)
            dds = DDS(obj, x_min, x_max, max_evals,
                      progress=_progress, stop_flag=stop_flag)
            res = dds.run()
            result_holder["res"] = res
        except Exception as exc:
            import traceback
            traceback.print_exc()
            result_holder["err"] = str(exc)
        try: setattr(app, "_cali_running", False)
        except Exception: pass
        # Reset the ETA cell to the static estimate now that we're
        # done — keeps the panel useful for the next planned run.
        cb = getattr(app, "_refresh_estim_time", None)
        if callable(cb):
            try:
                root.after(0, cb)
            except Exception:
                pass
        try:
            root.after(0, _on_done)
        except Exception:
            pass

    def _on_done():
        if result_holder["err"]:
            label.config(text="Calibration FAILED")
            detail.config(text=result_holder["err"])
            return
        res = result_holder["res"]
        if res is None:
            label.config(text="Cancelled.")
            return
        # Step 1 — write best DDS values into v_calib_mid.  This is
        # the canonical record of the optimum for each checked row;
        # everything below uses these Mid values as the source of
        # truth.
        mid = getattr(app, "v_calib_mid", [])
        try:
            label_to_idx = {lbl: i for i, lbl in enumerate(_CALIB_PARAMS)}
            for lbl, val in zip(labels, res["best_x"]):
                idx = label_to_idx.get(lbl)
                if idx is not None and idx < len(mid):
                    mid[idx].set(f"{val:g}")
        except Exception:
            pass

        # Step 2 — write gwvelocity_inputs.txt FIRST, BEFORE _push.
        # Reason: _push computes vd = K × i and sets v_darcy to the
        # product (a small number).  Its trailing _refresh would then
        # call _src_K — which falls back to v_darcy if no sidecar
        # exists — and overwrite mid[K] with the velocity instead of
        # K.  By writing the sidecar here first, _src_K reads K from
        # the sidecar and the refresh is consistent.
        try:
            k_idx = label_to_idx.get("Hydraulic Conductivity (k)")
            i_idx = label_to_idx.get("Hydraulic Gradient (i)")
            k_val = i_val_str = None
            if k_idx is not None and k_idx < len(mid):
                try: k_val = float(mid[k_idx].get())
                except Exception: pass
            if i_idx is not None and i_idx < len(mid):
                try: i_val_str = mid[i_idx].get().strip()
                except Exception: pass
            if k_val is not None or (i_val_str and i_val_str != ""):
                if not i_val_str:
                    i_val_str = "1"
                try: i_val = float(i_val_str)
                except Exception: i_val = 1.0
                if k_val is None:
                    try: k_val = float(app.v_darcy.get()) / max(i_val, 1e-12)
                    except Exception: k_val = 0.0
                vd_my = k_val * i_val
                vd_fy = vd_my / 0.3048
                gw_path = os.path.join(project, "gwvelocity_inputs.txt")
                with open(gw_path, "w", encoding="utf-8") as fp:
                    fp.write("Groundwater BulkDarcy Velocity Calculator Results\n")
                    fp.write(f"Bulk Hydraulic Conductivity Value: {k_val:g}\n")
                    fp.write(f"Bulk Hydraulic Conductivity Unit: m/year\n")
                    fp.write(f"Bulk Hydraulic Gradient: {i_val:g}\n")
                    fp.write(f"Bulk Darcy Velocity (m/year): {vd_my:.6f}\n")
                    fp.write(f"Bulk Darcy Velocity (ft/year): {vd_fy:.6f}\n")
        except Exception as exc:
            print(f"[cali] gwvelocity_inputs.txt write failed: {exc}")

        # Step 3 — delete heterogeneity sidecar (volfrac/difflen
        # optima live only in v_calib_mid for now; no Python writer
        # exists for heterogeneity_inputs.txt yet).
        try:
            sp = os.path.join(project, "heterogeneity_inputs.txt")
            if os.path.exists(sp):
                try: os.remove(sp)
                except Exception: pass
        except Exception:
            pass

        # Step 4 — push Mid → app source cells (vd = K × i, porf,
        # alphax, retardation).  With the sidecar already in place
        # (step 2), _push's trailing refresh reads K from sidecar
        # and Mid stays correct.
        cb = getattr(app, "_push_calib_mids_to_inputs", None)
        if callable(cb):
            try: cb()
            except Exception: pass

        # Step 5 — bake calibration multipliers physically into §7
        # source cells, then reset multiplier Mid to 1 so a future
        # run doesn't compound.  This makes the §7 view match what
        # the solver will see in input.inp at Run Optimal Model time.
        cb = getattr(app, "_apply_calib_multipliers_to_s7", None)
        if callable(cb):
            try: cb()
            except Exception: pass

        # Step 6 — recompute retardation from new porf + PFAA species.
        cb = getattr(app, "_on_pfaa_change", None)
        if callable(cb):
            try: cb()
            except Exception: pass

        # Save run history
        try:
            csv_path = os.path.join(project, "run_history.csv")
            with open(csv_path, "w", newline="", encoding="utf-8") as fp:
                w = csv.writer(fp)
                w.writerow(list(labels) + ["RMSLE"])
                for x_row, f_row in zip(res["x_history"],
                                        res["f_history"]):
                    w.writerow(list(x_row) + [-f_row])
        except Exception as exc:
            print(f"[cali] run_history.csv write failed: {exc}")

        best_rmsle = -res["best_cost"]
        label.config(text=f"Best RMSLE: {best_rmsle:.4g}")
        detail.config(
            text=f"{len(labels)} parameter(s) optimized over "
                 f"{len(res['x_history'])} evaluations.\n"
                 f"Best values pushed into the Mid-Range column.\n"
                 f"Full history → run_history.csv")
        # Replace Cancel with Close
        for child in btn_row.winfo_children():
            child.destroy()
        tk.Button(btn_row, text="Close", width=16, font=("Arial", 14),
                  command=root.destroy).pack(side="left", padx=14)

    threading.Thread(target=_worker, daemon=False).start()
    return True
