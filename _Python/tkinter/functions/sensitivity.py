# -*- coding: utf-8 -*-
"""
sensitivity.py — Monte Carlo sensitivity / uncertainty analysis for
REMFluor-MD.

Drives the "Run Sensitivity Analysis" button.  Based on the proof of
concept in calib2_Aman.py, ported onto the current cali_1.py (v108)
code base.  V1 is intentionally SERIAL — one solver run at a time —
so it is simple and correct; parallel execution can be added later.

What it does:
  1.  Asks the user for: number of runs (default 1000), plume boundary
      target concentration in ng/L (default 4), an optional year at
      which plume length is evaluated (blank = end of simulation), and
      a random seed.
  2.  Asks for a parent folder; all outputs go to <parent>/sensitivity/
      (input.inp + solver outputs are re-written there on every run,
      like the calibration work folder).
  3.  Draws N Latin-Hypercube samples over the CHECKED Step-4 parameter
      rows.  Each parameter gets a triangular distribution over
      (Lowest, Mid-Range, Highest); if Mid is blank / out of range the
      midpoint is used as the mode.  Pure numpy — no scipy required.
  4.  For each sample: overlay the values on the state snapshot (same
      PARAM_MAP + joint K×i logic as the calibrator), rebuild
      input.inp, run remfluor_v9a.exe, then:
        - RMSE and RMSLE vs the §10 observed well concentrations
        - plume length: furthest distance where the plume-centerline
          (max over Y,Z) concentration >= the target, linearly
          interpolated between grid nodes, from REMFluor-MD.out
  5.  Writes sensitivity_runs.csv (one row per run: every parameter,
      RMSE, RMSLE, per-well obs/sim, plume lengths, status) and
      plume_length_PFAA1.png / _PFAA2.png histograms with a smoothed
      density curve and P5 / P50 / mean / P95 statistics.

Requires: numpy + pandas (already bundled).  The PNG figure needs
matplotlib — REMEMBER to remove 'matplotlib' from the PyInstaller
--exclude-module list in build_exe.bat, otherwise the frozen build
will show the "matplotlib is required" error.
"""
from __future__ import annotations

import csv
import glob
import math
import os
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import messagebox, filedialog
from typing import Dict, List, Optional

import numpy as np

from . import generate_input_file
from .state import get_state, INPUT_TXT_FILES
from .cali_1 import PARAM_MAP, _set_vd, _read_obs_wells, _rmsle

# State override keys written by PARAM_MAP setters during evaluation.
# MUST be cleared when the run ends (success, cancel or crash) —
# leftover overrides make later manual Run Models silently ignore the
# dashboard values (the v107 calibration bug).
_CALIB_OVERRIDE_KEYS = (
    "_calib_volfrac", "_calib_difflen",
    "_calib_ock1", "_calib_ock2", "_calib_ock3", "_calib_ock4",
    "_calib_decayf1", "_calib_decayf3",
)


# ─────────────────────────────────────────────────────────────────────
# Sampling — pure numpy (scipy is excluded from the frozen build)
# ─────────────────────────────────────────────────────────────────────
def _lhs(ns: int, dim: int, rng: np.random.Generator) -> np.ndarray:
    """Latin Hypercube: one stratified uniform sample per row/column."""
    u = np.empty((ns, dim))
    for j in range(dim):
        # one point in each of ns equal-probability strata, shuffled
        strata = (rng.permutation(ns) + rng.random(ns)) / ns
        u[:, j] = strata
    return u


def _triangular_ppf(u: np.ndarray, lo: float, mode: float,
                    hi: float) -> np.ndarray:
    """Inverse CDF of the triangular distribution on [lo, hi] with the
    given mode.  Equivalent to scipy.stats.triang.ppf."""
    if hi <= lo:
        return np.full_like(u, lo)
    mode = min(max(mode, lo), hi)
    fc = (mode - lo) / (hi - lo)
    out = np.empty_like(u)
    left = u <= fc
    if fc > 0:
        out[left] = lo + np.sqrt(u[left] * (hi - lo) * (mode - lo))
    else:
        out[left] = lo
    if fc < 1:
        out[~left] = hi - np.sqrt((1.0 - u[~left]) * (hi - lo) * (hi - mode))
    else:
        out[~left] = hi
    return out


def _norm_ppf(u: np.ndarray) -> np.ndarray:
    """Inverse standard-normal CDF (Acklam's rational approximation,
    |error| < 1.15e-9) — avoids the scipy dependency."""
    a = [-3.969683028665376e+01, 2.209460984245205e+02,
         -2.759285104469687e+02, 1.383577518672690e+02,
         -3.066479806614716e+01, 2.506628277459239e+00]
    b = [-5.447609879822406e+01, 1.615858368580409e+02,
         -1.556989798598866e+02, 6.680131188771972e+01,
         -1.328068155288572e+01]
    c = [-7.784894002430293e-03, -3.223964580411365e-01,
         -2.400758277161838e+00, -2.549732539343734e+00,
         4.374664141464968e+00, 2.938163982698783e+00]
    d = [7.784695709041462e-03, 3.224671290700398e-01,
         2.445134137142996e+00, 3.754408661907416e+00]
    u = np.clip(np.asarray(u, dtype=float), 1e-15, 1 - 1e-15)
    out = np.empty_like(u)
    p_lo, p_hi = 0.02425, 1 - 0.02425
    lo = u < p_lo
    hi = u > p_hi
    mid = ~(lo | hi)
    if lo.any():
        q = np.sqrt(-2 * np.log(u[lo]))
        out[lo] = ((((((c[0]*q + c[1])*q + c[2])*q + c[3])*q + c[4])*q + c[5])
                   / ((((d[0]*q + d[1])*q + d[2])*q + d[3])*q + 1))
    if hi.any():
        q = np.sqrt(-2 * np.log(1 - u[hi]))
        out[hi] = -((((((c[0]*q + c[1])*q + c[2])*q + c[3])*q + c[4])*q + c[5])
                    / ((((d[0]*q + d[1])*q + d[2])*q + d[3])*q + 1))
    if mid.any():
        q = u[mid] - 0.5
        r = q * q
        out[mid] = ((((((a[0]*r + a[1])*r + a[2])*r + a[3])*r + a[4])*r + a[5])*q
                    / (((((b[0]*r + b[1])*r + b[2])*r + b[3])*r + b[4])*r + 1))
    return out


def _norm_cdf(x: float) -> float:
    return 0.5 * (1.0 + math.erf(x / math.sqrt(2.0)))


def _lognormal_trunc_ppf(u: np.ndarray, lo: float, mode: float,
                         hi: float) -> Optional[np.ndarray]:
    """Truncated log-normal mapped from the Step-4 values:
      * Mid-Range  = median (mu = ln(Mid))
      * Lowest / Highest = 5th / 95th percentiles (set sigma)
      * samples truncated to [Lowest, Highest]
    Returns None when the mapping is invalid (needs 0 < lo < hi and a
    positive Mid) — caller falls back to triangular for that row."""
    if not (lo > 0 and hi > lo):
        return None
    med = mode if (lo < mode < hi) else math.sqrt(lo * hi)
    mu = math.log(med)
    sigma = math.log(hi / lo) / (2 * 1.6448536269514722)   # P5..P95 span
    if sigma <= 0:
        return None
    # Truncate by remapping u into [F(lo), F(hi)] before the ppf.
    f_lo = _norm_cdf((math.log(lo) - mu) / sigma)
    f_hi = _norm_cdf((math.log(hi) - mu) / sigma)
    uu = f_lo + np.asarray(u, dtype=float) * (f_hi - f_lo)
    return np.exp(mu + sigma * _norm_ppf(uu))


def make_samples(x_min, x_mid, x_max, ns: int,
                 seed: Optional[int] = 9000,
                 dist: str = "triangular"):
    """ns × d matrix of parameter samples via Latin Hypercube
    stratification.  dist = "triangular" (Lowest/Mid/Highest as the
    triangle) or "lognormal" (Mid = median, Lowest/Highest = P5/P95,
    truncated to the bounds).  Log-normal rows whose values can't
    support the mapping (need 0 < Low < High) automatically fall back
    to triangular.  Returns (samples, fallback_indices)."""
    x_min = np.asarray(x_min, dtype=float)
    x_mid = np.asarray(x_mid, dtype=float)
    x_max = np.asarray(x_max, dtype=float)
    dim = x_min.size
    rng = np.random.default_rng(seed)
    u = _lhs(ns, dim, rng)
    out = np.empty_like(u)
    fallback = []
    for j in range(dim):
        col = None
        if dist == "lognormal":
            col = _lognormal_trunc_ppf(u[:, j], x_min[j], x_mid[j], x_max[j])
            if col is None:
                fallback.append(j)
        if col is None:
            col = _triangular_ppf(u[:, j], x_min[j], x_mid[j], x_max[j])
        out[:, j] = col
    return out, fallback


# ─────────────────────────────────────────────────────────────────────
# Plume length from REMFluor-MD.out
# ─────────────────────────────────────────────────────────────────────
def plume_length_from_md(md_path: str, target_ugl: float,
                         conc_col: str = "Conc2",
                         at_time: Optional[float] = None) -> float:
    """Plume length (m): furthest down-gradient distance where the
    plume-core concentration >= target_ugl, linearly interpolated
    between grid nodes.

    REMFluor-MD.out is a CSV: Time,X,Y,Z,Conc1..Conc4 written every
    npt timesteps.  We take the snapshot nearest `at_time` (or the
    LAST snapshot when at_time is None), reduce to the plume core by
    taking max concentration over Y,Z at each X, then find the last
    crossing of the target.

    Returns nan if the file/column is missing, 0.0 if the target is
    never reached (no plume above the boundary concentration)."""
    import pandas as pd
    try:
        df = pd.read_csv(md_path, skipinitialspace=True)
        df.columns = [str(c).strip() for c in df.columns]
    except Exception:
        return float("nan")
    if conc_col not in df.columns or "X" not in df.columns:
        return float("nan")

    times = df["Time"].unique()
    if len(times) == 0:
        return float("nan")
    t_sel = (times[-1] if at_time is None
             else times[np.argmin(np.abs(times - at_time))])
    snap = df[df["Time"] == t_sel]

    # Plume core: max over Y,Z at each X, sorted by distance.
    prof = snap.groupby("X")[conc_col].max().sort_index()
    x = prof.index.to_numpy(dtype=float)
    c = prof.to_numpy(dtype=float)
    if x.size == 0:
        return float("nan")

    above = c >= target_ugl
    if not above.any():
        return 0.0
    i_last = int(np.max(np.nonzero(above)))
    if i_last == len(x) - 1:
        # Plume extends to (at least) the end of the model domain.
        return float(x[-1])
    # Interpolate between the last node above target and the next one.
    x0, x1 = x[i_last], x[i_last + 1]
    c0, c1 = c[i_last], c[i_last + 1]
    if c0 == c1:
        return float(x0)
    return float(x0 + (c0 - target_ugl) * (x1 - x0) / (c0 - c1))


# ─────────────────────────────────────────────────────────────────────
# Histogram + density PNG
# ─────────────────────────────────────────────────────────────────────
def _gaussian_kde(vals: np.ndarray, grid: np.ndarray) -> np.ndarray:
    """Minimal Gaussian KDE (Silverman bandwidth) — avoids scipy."""
    n = vals.size
    sd = np.std(vals)
    if n < 2 or sd == 0:
        return np.zeros_like(grid)
    bw = 1.06 * sd * n ** (-1 / 5)
    diff = (grid[:, None] - vals[None, :]) / bw
    return np.exp(-0.5 * diff ** 2).sum(axis=1) / (n * bw * math.sqrt(2 * math.pi))


def _compute_stats(vals: List[float]) -> Optional[Dict]:
    v = np.asarray([x for x in vals if x is not None and np.isfinite(x)],
                   dtype=float)
    if v.size == 0:
        return None
    return {
        "_v": v,
        "n": int(v.size),
        "mean": float(np.mean(v)), "std": float(np.std(v)),
        "min": float(np.min(v)), "max": float(np.max(v)),
        "p05": float(np.percentile(v, 5)),
        "p50": float(np.percentile(v, 50)),
        "p95": float(np.percentile(v, 95)),
    }


def _draw_hist_panel(ax1, stats: Dict, title: str, xlabel: str,
                     unit: str = "", fmt: str = ".1f",
                     legend: bool = True, title_size: Optional[int] = None):
    """Draw one dual-axis panel into `ax1`: histogram in COUNTS on the
    left axis, probability density (KDE) on a twinned right axis kept
    proportional (density = count / (n × bin width)) so the curve
    envelopes the bars, plus P5-P95 band + median/mean lines."""
    v = stats["_v"]
    u = f" {unit}" if unit else ""
    nbins = max(10, min(40, int(math.sqrt(v.size) * 2)))
    counts, edges, _ = ax1.hist(
        v, bins=nbins, color="#1F639E", edgecolor="white",
        alpha=0.85, zorder=2, label=f"Count ({v.size} runs)")
    binw = float(edges[1] - edges[0]) if len(edges) > 1 else 1.0
    ax1.set_xlabel(xlabel)
    ax1.set_ylabel("Count of runs", color="#1F639E")
    ax1.tick_params(axis="y", labelcolor="#1F639E")

    ax2 = ax1.twinx()
    if v.size >= 2 and np.std(v) > 0:
        grid = np.linspace(v.min() - 2 * np.std(v),
                           v.max() + 2 * np.std(v), 400)
        ax2.plot(grid, _gaussian_kde(v, grid), color="#C43B3B", lw=2.5,
                 zorder=3, label="Probability density")
    ax2.set_ylabel("Probability density", color="#C43B3B")
    ax2.tick_params(axis="y", labelcolor="#C43B3B")
    ax1.set_ylim(bottom=0)
    ax2.set_ylim(0, ax1.get_ylim()[1] / (v.size * binw))

    ax1.axvspan(stats["p05"], stats["p95"], color="#1F639E", alpha=0.10,
                zorder=1, label="90% interval (P5-P95)")
    ax1.axvline(stats["p50"], color="#333333", lw=1.6, ls="--",
                label=f"Median = {stats['p50']:{fmt}}{u}")
    ax1.axvline(stats["mean"], color="#C47F1B", lw=1.6, ls="-.",
                label=f"Mean = {stats['mean']:{fmt}}{u}")
    ax1.set_title(f"{title}\n"
                  f"P5 = {stats['p05']:{fmt}}{u}   "
                  f"P50 = {stats['p50']:{fmt}}{u}   "
                  f"P95 = {stats['p95']:{fmt}}{u}   (n = {stats['n']})",
                  fontsize=title_size)
    if legend:
        h1, l1 = ax1.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()
        ax1.legend(h1 + h2, l1 + l2, fontsize=8)
    ax1.grid(alpha=0.25, zorder=0)


def _get_pyplot():
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        return plt
    except ImportError:
        print("[sensitivity] matplotlib not available — PNG skipped "
              "(remove 'matplotlib' from the PyInstaller excludes).")
        return None


def write_histogram_png(vals: List[float], out_png: str, title: str,
                        xlabel: str, unit: str = "",
                        fmt: str = ".1f") -> Optional[Dict]:
    """Single dual-axis histogram figure (used for RMSE)."""
    stats = _compute_stats(vals)
    if stats is None:
        return None
    plt = _get_pyplot()
    if plt is None:
        return stats
    fig, ax1 = plt.subplots(figsize=(9, 5.5), dpi=150)
    _draw_hist_panel(ax1, stats, title, xlabel, unit, fmt)
    fig.tight_layout()
    fig.savefig(out_png)
    plt.close(fig)
    return stats


def write_plume_grid_png(series: List, out_png: str,
                         target_ngl: float,
                         time_label: str) -> Dict[str, Dict]:
    """ONE combined plume-length figure — a 2-column grid with one
    dual-axis panel per configured species (1 species → single panel,
    2 → side by side, 3-4 → 2×2).  `series` is a list of
    (label, values) tuples.  Returns {label: stats}."""
    panels = []
    all_stats: Dict[str, Dict] = {}
    for label, vals in series:
        st = _compute_stats(vals)
        if st is not None:
            panels.append((label, st))
            all_stats[label] = st
    if not panels:
        return all_stats
    plt = _get_pyplot()
    if plt is None:
        return all_stats

    n = len(panels)
    ncols = 1 if n == 1 else 2
    nrows = (n + ncols - 1) // ncols
    fig, axes = plt.subplots(nrows, ncols,
                             figsize=(9 * ncols * 0.85, 4.6 * nrows),
                             dpi=150, squeeze=False)
    for k, (label, st) in enumerate(panels):
        ax = axes[k // ncols][k % ncols]
        _draw_hist_panel(ax, st, label, "Plume length (m)", unit="m",
                         legend=(k == 0), title_size=10)
    # Hide any unused grid cell (e.g. 3 species on a 2x2 grid)
    for k in range(n, nrows * ncols):
        axes[k // ncols][k % ncols].axis("off")
    fig.suptitle(f"Plume length @ {target_ngl:g} ng/L — {time_label}",
                 fontsize=13)
    fig.tight_layout(rect=(0, 0, 1, 0.95))
    fig.savefig(out_png)
    plt.close(fig)
    return all_stats


# ─────────────────────────────────────────────────────────────────────
# Public entry point
# ─────────────────────────────────────────────────────────────────────
def _center_over(win, anchor, width=None, height=None):
    """Place Toplevel `win` centered over `anchor`'s toplevel window.
    Using the ANCHOR's geometry (not winfo_screenwidth) means the popup
    opens on the SAME monitor the app is on — screenwidth-based
    centering put popups on the primary monitor even when the app was
    running on a secondary one."""
    try:
        win.update_idletasks()
        w = width or win.winfo_reqwidth()
        h = height or win.winfo_reqheight()
        top = anchor.winfo_toplevel()
        ax, ay = top.winfo_rootx(), top.winfo_rooty()
        aw, ah = top.winfo_width(), top.winfo_height()
        if aw <= 1 or ah <= 1:      # app not mapped yet — fall back
            raise ValueError
        x = ax + max(0, (aw - w) // 2)
        y = ay + max(0, (ah - h) // 3)
        if width or height:
            win.geometry(f"{w}x{h}+{x}+{y}")
        else:
            win.geometry(f"+{x}+{y}")
    except Exception:
        try:
            sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
            win.geometry(f"+{max(0, (sw - (width or 400)) // 2)}"
                         f"+{max(0, (sh - (height or 300)) // 2)}")
        except Exception:
            pass


def run(app, parent=None) -> bool:
    """Run the Monte Carlo sensitivity analysis.  True if started."""
    try:
        import numpy as _np  # noqa
        import pandas as _pd  # noqa
    except ImportError as exc:
        messagebox.showerror("Run Sensitivity Analysis",
                             f"numpy + pandas are required:\n{exc}")
        return False

    from main import _CALIB_PARAMS  # type: ignore

    # ── Gather checked parameter rows (same rules as cali_1.run) ────
    chk = getattr(app, "v_calib_chk", [])
    lo_l = getattr(app, "v_calib_low", [])
    hi_l = getattr(app, "v_calib_high", [])
    mid_l = getattr(app, "v_calib_mid", [])

    def _fnum(var_list, i):
        try:
            return float(str(var_list[i].get()).replace(",", "").strip())
        except (ValueError, TypeError, AttributeError, IndexError):
            return None

    labels: List[str] = []
    x_min: List[float] = []
    x_mid: List[float] = []
    x_max: List[float] = []
    for i, label in enumerate(_CALIB_PARAMS):
        if i >= len(chk):
            break
        try:
            if not bool(chk[i].get()):
                continue
        except Exception:
            continue
        lo_v, hi_v = _fnum(lo_l, i), _fnum(hi_l, i)
        if lo_v is None or hi_v is None or hi_v <= lo_v:
            continue
        mid_v = _fnum(mid_l, i)
        if mid_v is None or not (lo_v <= mid_v <= hi_v):
            mid_v = lo_v + 0.5 * (hi_v - lo_v)
        labels.append(label)
        x_min.append(lo_v)
        x_mid.append(mid_v)
        x_max.append(hi_v)

    if not labels:
        messagebox.showinfo(
            "Run Sensitivity Analysis",
            "No calibration parameters checked.\n\n"
            "Tick at least one row in §Step 4 and set Lowest / Highest\n"
            "Likely Values, then try again.")
        return False

    # ── Observed concentrations (§10) for RMSE / RMSLE ──────────────
    def _obs_list(attr):
        out = []
        for v in getattr(app, attr, [])[:7]:
            try:
                out.append(float(str(v.get()).replace(",", "").strip()))
            except (ValueError, TypeError, AttributeError):
                out.append(float("nan"))
        return out
    obs_pfaa1 = _obs_list("v_mw_conc")
    obs_pfaa2 = _obs_list("v_mw_conc2")
    has_any_obs = any(not math.isnan(v) for v in obs_pfaa1 + obs_pfaa2)
    if not has_any_obs:
        if not messagebox.askyesno(
                "Run Sensitivity Analysis",
                "No observed concentrations in §10 — RMSE/RMSLE cannot\n"
                "be computed and those columns will be blank.\n\n"
                "Continue anyway (plume-length uncertainty only)?"):
            return False

    # ── Configured species → plume-length targets ────────────────────
    # Generic: one plume length per species the model actually
    # simulates.  REMFluor-MD.out columns: Conc1 = Precursor-1,
    # Conc2 = PFAA-1, Conc3 = Precursor-2, Conc4 = PFAA-2.
    # Precursors exist only in Detailed mode.
    def _species_name(attr):
        try:
            s = str(getattr(app, attr).get()).strip()
            return s if s and s.lower() != "none" else None
        except Exception:
            return None

    is_simple = (getattr(app, "active_sheet", "Detailed_2") == "Simple")
    pfaa1_name = _species_name("v_pfaa1") or "PFAA-1"
    # (key, panel label, REMFluor-MD.out column)
    species = [("PFAA1", f"PFAA-1 ({pfaa1_name})", "Conc2")]
    _n = _species_name("v_pfaa2")
    has_pfaa2 = _n is not None
    if has_pfaa2:
        species.append(("PFAA2", f"PFAA-2 ({_n})", "Conc4"))
    if not is_simple:
        _n = _species_name("v_pfaa3")   # Precursor 1 (K38)
        if _n:
            species.append(("PRE1", f"Precursor-1 ({_n})", "Conc1"))
        _n = _species_name("v_pfaa4")   # Precursor 2 (M38)
        if _n:
            species.append(("PRE2", f"Precursor-2 ({_n})", "Conc3"))

    try:
        sample_year = int(float(str(app.v_sample_yr.get()).strip()))
    except (ValueError, TypeError, AttributeError):
        sample_year = 0

    # ── Settings dialog ──────────────────────────────────────────────
    dlg = tk.Toplevel(parent or app)
    dlg.title("Sensitivity Analysis — Settings")
    dlg.configure(bg="#F0F0F0")
    dlg.transient(parent or app)
    dlg.grab_set()

    # v109: defaults come from the calibration panel — "Number of runs
    # for sensitivity" box (v_sens_runs) and the Distribution selector
    # (v_sens_dist).  The user can still adjust runs here.
    try:
        default_runs = str(int(float(str(app.v_sens_runs.get()).strip())))
    except (ValueError, TypeError, AttributeError):
        default_runs = "1000"
    try:
        dist = str(app.v_sens_dist.get()).strip().lower()
    except Exception:
        dist = "triangular"
    if dist not in ("triangular", "lognormal"):
        dist = "triangular"

    # v109: target concentration default comes from the calibration
    # panel's "Concentration for plume (ng/L)" box (v_sens_target).
    try:
        default_target = str(float(str(app.v_sens_target.get()).strip()))
    except Exception:
        default_target = "4"

    fields = [
        ("Number of runs:", default_runs),
        ("Plume boundary target concentration (ng/L):", default_target),
        ("Year for plume length (blank = end of simulation):", ""),
        ("Random seed:", "9000"),
    ]
    entries = []
    for r, (lbl, default) in enumerate(fields):
        tk.Label(dlg, text=lbl, bg="#F0F0F0", font=("Arial", 11),
                 anchor="w").grid(row=r, column=0, sticky="w",
                                  padx=(16, 8), pady=6)
        e = tk.Entry(dlg, width=12, font=("Arial", 11))
        e.insert(0, default)
        e.grid(row=r, column=1, sticky="w", padx=(0, 16), pady=6)
        entries.append(e)
    _dist_desc = ("Triangular (peak at Mid-Range)" if dist == "triangular"
                  else "Log-normal (Mid = median, Low/High = P5/P95)")
    tk.Label(dlg, text=f"{len(labels)} checked parameter(s) will be varied.\n"
                       f"Distribution: {_dist_desc}\n"
                       f"(change it in the calibration panel's "
                       f"'Distribution for Sensitivity Analysis' box)",
             bg="#F0F0F0", fg="#555", font=("Arial", 10), justify="left"
             ).grid(row=len(fields), column=0, columnspan=2,
                    padx=16, pady=(4, 8), sticky="w")

    settings = {}

    def _ok():
        try:
            settings["ns"] = max(1, int(float(entries[0].get())))
            settings["target_ngl"] = float(entries[1].get())
            yr = entries[2].get().strip()
            settings["plume_year"] = int(float(yr)) if yr else None
            sd = entries[3].get().strip()
            settings["seed"] = int(float(sd)) if sd else None
        except ValueError:
            messagebox.showerror("Sensitivity Analysis",
                                 "Please enter valid numbers.", parent=dlg)
            return
        if settings["target_ngl"] <= 0:
            messagebox.showerror("Sensitivity Analysis",
                                 "Target concentration must be > 0.",
                                 parent=dlg)
            return
        dlg.destroy()

    btns = tk.Frame(dlg, bg="#F0F0F0")
    btns.grid(row=len(fields) + 1, column=0, columnspan=2, pady=(4, 14))
    tk.Button(btns, text="Start", width=12, command=_ok).pack(
        side="left", padx=8)
    tk.Button(btns, text="Cancel", width=12,
              command=dlg.destroy).pack(side="left", padx=8)
    _center_over(dlg, parent or app)
    dlg.wait_window()
    if "ns" not in settings:
        return False    # cancelled

    ns = settings["ns"]
    target_ugl = settings["target_ngl"] / 1000.0    # solver works in ug/L
    plume_year = settings["plume_year"]

    # ── Resolve solver + template (same search as cali_1 v108) ──────
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(os.path.abspath(sys.executable))
        bundle = getattr(sys, "_MEIPASS", exe_dir)
    else:
        here = os.path.dirname(os.path.abspath(__file__))
        exe_dir = os.path.abspath(os.path.join(here, "..", "..", ".."))
        bundle = exe_dir
    exe_path = None
    for _name in ("remfluor_v9a.exe", "remfluor_v8a.exe"):
        for _base in (exe_dir, bundle):
            cand = os.path.join(_base, _name)
            if os.path.exists(cand):
                exe_path = cand
                break
        if exe_path:
            break
    if exe_path is None or not os.path.exists(exe_path):
        messagebox.showerror("Run Sensitivity Analysis",
                             "Model executable not found "
                             "(remfluor_v9a.exe / remfluor_v8a.exe).")
        return False

    # ── Output folder: <parent>/sensitivity ─────────────────────────
    state = get_state()
    old_wd = getattr(state, "work_dir", "") or os.getcwd()
    parent_dir = filedialog.askdirectory(
        title="Choose the parent folder for the sensitivity results",
        initialdir=old_wd, mustexist=False, parent=parent or app)
    if not parent_dir:
        return False
    work = os.path.join(parent_dir, "sensitivity")
    # Same pattern as Run Model / Run Calibration: if the sensitivity
    # folder already holds results from a previous run, confirm before
    # overwriting them.
    _outputs = ["sensitivity_runs.csv", "plume_length_histograms.png",
                "rmse_histogram.png", "sensitivity_log.txt"]
    _existing = [f for f in _outputs
                 if os.path.exists(os.path.join(work, f))]
    if _existing:
        _lst = "\n".join(f"  {f}" for f in _existing)
        if not messagebox.askyesno(
                "Overwrite Existing Files?",
                "This folder already contains sensitivity results that "
                f"will be overwritten:\n{work}\n\n{_lst}\n\n"
                "Overwrite them?",
                parent=parent or app):
            return False
    try:
        os.makedirs(work, exist_ok=True)
    except Exception as exc:
        messagebox.showerror("Run Sensitivity Analysis",
                             f"Cannot create output folder:\n{work}\n{exc}")
        return False

    # Copy model sidecars + template.inp so build_inp_data reads the
    # right inputs from the sensitivity folder (mirrors cali_1 v108).
    for _fn in INPUT_TXT_FILES:
        _src = os.path.join(old_wd, _fn)
        if os.path.exists(_src):
            try:
                shutil.copy2(_src, os.path.join(work, _fn))
            except Exception:
                pass
    if not os.path.exists(os.path.join(work, "template.inp")):
        for _b in (old_wd, exe_dir, bundle):
            _tpl = os.path.join(_b, "template.inp")
            if os.path.exists(_tpl):
                try:
                    shutil.copy2(_tpl, os.path.join(work, "template.inp"))
                except Exception:
                    pass
                break

    # ── Snapshot UI state ON THE MAIN THREAD ─────────────────────────
    state.snapshot(app)
    baseline = dict(state._cells)
    state.work_dir = work

    try:
        start_year = int(float(state.get("E18", 1977)))
    except (ValueError, TypeError):
        start_year = 1977
    target_t_obs = max(0, sample_year - start_year)     # for RMSE rows
    plume_t = (None if plume_year is None
               else max(0, plume_year - start_year))    # for plume length

    samples, dist_fallback = make_samples(x_min, x_mid, x_max, ns,
                                          settings["seed"], dist=dist)
    if dist_fallback:
        fb_names = [labels[j] for j in dist_fallback]
        messagebox.showwarning(
            "Sensitivity Analysis",
            "Log-normal needs 0 < Lowest < Highest.  These rows fell "
            "back to the Triangular distribution:\n\n"
            + "\n".join(f"  • {n}" for n in fb_names),
            parent=parent or app)

    # ── Progress popup ───────────────────────────────────────────────
    root = tk.Toplevel(parent or app)
    root.title("REMFluor Sensitivity Analysis in Progress")
    root.configure(bg="#F0F0F0")
    try:
        root.transient(parent or app)
    except Exception:
        pass
    title = tk.Label(root, text="Sensitivity Analysis in Progress",
                     font=("Arial", 20, "bold"), bg="#F0F0F0", fg="#222",
                     padx=24, pady=14)
    title.pack(padx=36, pady=(28, 6))
    label = tk.Label(root, text="Starting...", font=("Arial", 22, "bold"),
                     bg="#F0F0F0", fg="#222", padx=24, pady=8)
    label.pack(padx=36, pady=(0, 8))
    detail = tk.Label(root, text="", font=("Arial", 13), bg="#F0F0F0",
                      fg="#444", justify="center", wraplength=1000)
    detail.pack(padx=36, pady=(0, 22))
    stop_flag = {"stop": False}
    btn_row = tk.Frame(root, bg="#F0F0F0")
    btn_row.pack(pady=(0, 32))

    def _cancel():
        stop_flag["stop"] = True
        label.config(text="Cancelling after current run...")
    tk.Button(btn_row, text="Cancel", width=16, font=("Arial", 14),
              command=_cancel).pack(side="left", padx=14)
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 64, 900)
    h = max(root.winfo_reqheight() + 60, 420)
    _center_over(root, parent or app, width=w, height=h)
    root.minsize(w, h)

    log_path = os.path.join(work, "sensitivity_log.txt")

    def _log(msg):
        try:
            with open(log_path, "a", encoding="utf-8") as fp:
                fp.write(msg + "\n")
        except Exception:
            pass
    try:
        with open(log_path, "w", encoding="utf-8") as fp:
            fp.write("REMFluor-MD sensitivity analysis log\n"
                     f"runs={ns}  target={settings['target_ngl']} ng/L "
                     f"({target_ugl} ug/L)  plume_year={plume_year}  "
                     f"seed={settings['seed']}  distribution={dist}\n"
                     f"parameters={labels}\n"
                     + (f"triangular fallback rows={[labels[j] for j in dist_fallback]}\n\n"
                        if dist_fallback else "\n"))
    except Exception:
        pass

    # ── One model evaluation ─────────────────────────────────────────
    def _apply_sample(x):
        state._cells.clear()
        state._cells.update(baseline)
        K_val = i_val = None
        for lbl, val in zip(labels, x):
            entry = PARAM_MAP.get(lbl)
            if entry is None:
                continue
            if entry == "K":
                K_val = float(val)
                continue
            if entry == "i":
                i_val = float(val)
                continue
            if callable(entry):
                try:
                    entry(state, float(val))
                except Exception:
                    pass
        if K_val is not None or i_val is not None:
            try:
                base_vd = float(str(baseline.get("C22", "0"))
                                .replace(",", "").strip())
            except Exception:
                base_vd = 0.0
            K = K_val if K_val is not None else (base_vd or 0.0)
            i = i_val if i_val is not None else 1.0
            _set_vd(state, K * i)

    def _run_one(run_id, x):
        """Returns a result-row dict; 'status' is 'ok' or an error tag."""
        row = {"run_id": run_id, "status": "ok",
               "rmse": None, "rmsle": None,
               "sim1": [None] * 7, "sim2": [None] * 7,
               "plume": {}}
        _apply_sample(x)
        try:
            data = generate_input_file.build_inp_data(state)
        except Exception as exc:
            _log(f"run {run_id}: build_inp_data failed: {exc}")
            row["status"] = "build_failed"
            return row
        try:
            with open(os.path.join(work, "template.inp")) as fp:
                tmpl = fp.read()
            src_lines = [f"{t or 0}, {c1 or 0}, {c2 or 0}, {c3 or 0}, {c4 or 0}"
                         for t, c1, c2, c3, c4 in zip(
                             data['times'], data['concs1'], data['concs2'],
                             data['concs3'], data['concs4'])]
            filled = tmpl.format(
                source_zone_data="\n".join(src_lines),
                observation_wells="\n".join(data['wells']),
                **{k: v for k, v in data.items()
                   if k not in ('times', 'concs1', 'concs2',
                                'concs3', 'concs4', 'wells')})
            with open(os.path.join(work, "input.inp"), "w") as fp:
                fp.write(filled)
        except Exception as exc:
            _log(f"run {run_id}: template fill failed: {exc}")
            row["status"] = "template_failed"
            return row
        try:
            cmd = '"{}" < "input.inp" > "output.out"'.format(exe_path)
            subprocess.run(cmd, shell=True, cwd=work, timeout=600,
                           capture_output=True)
        except Exception as exc:
            _log(f"run {run_id}: solver failed: {exc}")
            row["status"] = "solver_failed"
            return row

        # RMSE / RMSLE vs §10 observations
        wells = _read_obs_wells(work)
        if wells:
            names = sorted(wells.keys())
            sims, obss = [], []
            for i, name in enumerate(names[:7]):
                rows_w = wells[name]
                best = min(rows_w, key=lambda r: abs(r[0] - target_t_obs))
                s1 = best[2] if len(best) > 2 else 0.0
                s2 = best[4] if len(best) > 4 else 0.0
                row["sim1"][i] = s1
                row["sim2"][i] = s2
                if i < len(obs_pfaa1) and not math.isnan(obs_pfaa1[i]):
                    sims.append(s1); obss.append(obs_pfaa1[i])
                if has_pfaa2 and i < len(obs_pfaa2) \
                        and not math.isnan(obs_pfaa2[i]):
                    sims.append(s2); obss.append(obs_pfaa2[i])
            if obss:
                a_o = np.asarray(obss, dtype=float)
                a_s = np.asarray(sims, dtype=float)
                row["rmse"] = float(np.sqrt(np.mean((a_s - a_o) ** 2)))
                row["rmsle"] = _rmsle(a_o, a_s)
        else:
            _log(f"run {run_id}: no obs_well*.out produced")
            row["status"] = "no_output"
            return row

        # Plume length per configured species from REMFluor-MD.out
        md = os.path.join(work, "REMFluor-MD.out")
        for key, _lbl, col in species:
            row["plume"][key] = plume_length_from_md(md, target_ugl,
                                                     col, plume_t)
        return row

    # ── Worker thread: serial loop over samples ──────────────────────
    results: List[Dict] = []
    holder = {"err": None, "t0": time.time()}

    def _fmt_eta(done, total, t0):
        el = time.time() - t0
        if done == 0:
            return ""
        rem = el / done * (total - done)
        return (f"Elapsed {el/60:.1f} min — "
                f"~{rem/60:.1f} min remaining")

    def _worker():
        try:
            for k in range(ns):
                if stop_flag.get("stop"):
                    break
                res = _run_one(k, samples[k, :])
                results.append(res)
                done = k + 1
                try:
                    root.after(0, lambda d=done: label.config(
                        text=f"Run {d}/{ns}"))
                    n_ok = sum(1 for r in results if r["status"] == "ok")
                    pl = [r["plume"].get("PFAA1") for r in results
                          if r["plume"].get("PFAA1") is not None
                          and np.isfinite(r["plume"].get("PFAA1"))]
                    msg = (f"{n_ok} ok / {done - n_ok} failed    "
                           + _fmt_eta(done, ns, holder["t0"]))
                    if pl:
                        msg += (f"\nPFAA-1 plume length so far: "
                                f"median {np.median(pl):.1f} m "
                                f"(min {min(pl):.1f}, max {max(pl):.1f})")
                    root.after(0, lambda m=msg: detail.config(text=m))
                except Exception:
                    pass
        except Exception as exc:
            import traceback
            traceback.print_exc()
            holder["err"] = str(exc)
        finally:
            # ALWAYS drop the _calib_* overrides + restore work_dir.
            try:
                state._cells.clear()
                state._cells.update(baseline)
                for kk in _CALIB_OVERRIDE_KEYS:
                    try:
                        state.set(kk, None)
                    except Exception:
                        pass
                state.work_dir = old_wd
            except Exception:
                pass
        try:
            root.after(0, _on_done)
        except Exception:
            pass

    def _on_done():
        if holder["err"]:
            label.config(text="Sensitivity analysis FAILED")
            detail.config(text=holder["err"])
            return
        n_done = len(results)
        n_ok = sum(1 for r in results if r["status"] == "ok")

        # ── sensitivity_runs.csv ─────────────────────────────────────
        csv_path = os.path.join(work, "sensitivity_runs.csv")
        try:
            hdr = (["run_id"] + list(labels) + ["RMSE", "RMSLE"]
                   + [f"obs_PFAA1_w{i+1}" for i in range(7)]
                   + [f"sim_PFAA1_w{i+1}" for i in range(7)])
            if has_pfaa2:
                hdr += ([f"obs_PFAA2_w{i+1}" for i in range(7)]
                        + [f"sim_PFAA2_w{i+1}" for i in range(7)])
            hdr += [f"plume_length_{key}_m" for key, _l, _c in species]
            hdr += ["status"]

            def _c(v):
                if v is None:
                    return ""
                try:
                    if isinstance(v, float) and math.isnan(v):
                        return ""
                except Exception:
                    pass
                return v
            with open(csv_path, "w", newline="", encoding="utf-8") as fp:
                wtr = csv.writer(fp)
                wtr.writerow(hdr)
                for r in results:
                    line = ([r["run_id"]]
                            + [f"{v:g}" for v in samples[r["run_id"], :]]
                            + [_c(r["rmse"]), _c(r["rmsle"])]
                            + [_c(o if not math.isnan(o) else None)
                               for o in obs_pfaa1]
                            + [_c(s) for s in r["sim1"]])
                    if has_pfaa2:
                        line += ([_c(o if not math.isnan(o) else None)
                                  for o in obs_pfaa2]
                                 + [_c(s) for s in r["sim2"]])
                    line += [_c(r["plume"].get(key))
                             for key, _l, _cc in species]
                    line += [r["status"]]
                    wtr.writerow(line)
        except Exception as exc:
            _log(f"CSV write failed: {exc}")

        # ── PNG figures — only OK runs.  EXACTLY TWO FILES: ──────────
        #  1. plume_length_histograms.png — one combined figure, a
        #     2-column grid with one dual-axis panel per configured
        #     species (PFAAs + Precursors in Detailed mode).
        #  2. rmse_histogram.png — single panel (one fit metric for
        #     the whole model, all wells/species together).
        time_label = (f"year {plume_year}" if plume_year is not None
                      else "end of simulation")
        png_msgs = []
        try:
            series = []
            for key, lbl, _cc in species:
                vals = [r["plume"].get(key) for r in results
                        if r["status"] == "ok"]
                series.append((lbl, vals))
            grid_stats = write_plume_grid_png(
                series, os.path.join(work, "plume_length_histograms.png"),
                settings["target_ngl"], time_label)
            for lbl, st in grid_stats.items():
                png_msgs.append(
                    f"{lbl}: P5={st['p05']:.1f}  P50={st['p50']:.1f}  "
                    f"P95={st['p95']:.1f} m")
            rmse_vals = [r["rmse"] for r in results if r["status"] == "ok"]
            st3 = write_histogram_png(
                rmse_vals, os.path.join(work, "rmse_histogram.png"),
                "RMSE of simulated vs observed well concentrations "
                "(all sensitivity runs)",
                "RMSE (ug/L)", unit="", fmt=".3g")
            if st3:
                png_msgs.append(
                    f"RMSE: P5={st3['p05']:.3g}  P50={st3['p50']:.3g}  "
                    f"P95={st3['p95']:.3g}")
        except Exception as exc:
            _log(f"PNG write failed: {exc}")

        cancelled = stop_flag.get("stop") and n_done < ns
        label.config(text=("Cancelled — partial results saved"
                           if cancelled else "Sensitivity analysis complete"))
        detail.config(text=(
            f"{n_ok} of {n_done} runs succeeded.\n"
            + "\n".join(png_msgs)
            + f"\n\nResults saved in:\n{work}\n"
              "(sensitivity_runs.csv + plume_length_histograms.png + "
              "rmse_histogram.png + sensitivity_log.txt)"))
        for child in btn_row.winfo_children():
            child.destroy()
        tk.Button(btn_row, text="Close", width=16, font=("Arial", 14),
                  command=root.destroy).pack(side="left", padx=14)

    threading.Thread(target=_worker, daemon=False).start()
    return True
