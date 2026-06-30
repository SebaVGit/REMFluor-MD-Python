"""
visualize_saved.py — §11 "Visualize Results" button.

Lets the user pick a model OUTPUT file (*.out) — including runs produced
OUTSIDE this GUI (e.g. the Fortran solver run directly) — and opens the
Plotly-Dash dashboard for every result file sitting next to it, WITHOUT
re-running the solver.

The dashboard reads these from the project dir:
    REMFluor-MD.out   — main plume output            (REQUIRED)
    discharge.out     — mass discharge (kg/yr)        (REQUIRED)
    plume_mass.out    — plume mass (kg)               (REQUIRED)
    obs_well1..N.out  — modeled MW concentrations     (recommended)
    dashboard_state.json — run context (years, names) (auto-generated
                        from the current form if absent)

If dashboard_state.json is missing it is regenerated from the values
currently in the form, so result sets created outside this GUI can still
be visualized.  Clear notifications are shown for anything missing.
"""
import os
import glob
import shutil
from tkinter import messagebox, filedialog

from . import run_model
from .state import get_state


# .out files the dashboard reads UNCONDITIONALLY — all three must exist
# or the dashboard subprocess crashes on read.
REQUIRED_OUT = ["REMFluor-MD.out", "discharge.out", "plume_mass.out"]


def _gather_result_files(folder):
    """Full paths of every dashboard result file present in *folder*."""
    found = []
    for name in run_model.DASHBOARD_RESULT_FILES:
        p = os.path.join(folder, name)
        if os.path.exists(p):
            found.append(p)
    for pat in run_model.DASHBOARD_RESULT_GLOBS:
        found.extend(glob.glob(os.path.join(folder, pat)))
    return found


def run(app, parent=None):
    """Prompt for a results folder, stage the files (generating
    dashboard_state.json if needed), then launch the dashboard."""
    state = get_state()
    work_dir = state.work_dir or os.getcwd()
    par = parent or app

    # Show a FILE picker filtered to *.out (intuitive — the user sees the
    # result files, not just folders).  They pick the main output file and
    # we load every result file sitting next to it.
    chosen = filedialog.askopenfilename(
        title="Select the model output file (REMFluor-MD.out)",
        initialdir=work_dir,
        filetypes=[("REMFluor result files", "*.out"),
                   ("All files", "*.*")],
    )
    if not chosen:
        return  # user cancelled
    folder = os.path.dirname(chosen)

    # ── Validate the REQUIRED .out files ────────────────────────────────
    have = {n: os.path.exists(os.path.join(folder, n)) for n in REQUIRED_OUT}
    if not have["REMFluor-MD.out"]:
        messagebox.showerror(
            "Visualize Results",
            "No model output was found in the selected folder.\n\n"
            "Required: REMFluor-MD.out\n\n"
            "Pick a folder that contains the *.out files produced by a "
            "model run (from this app or the solver run directly).",
            parent=par,
        )
        return
    missing_req = [n for n in REQUIRED_OUT if not have[n]]
    if missing_req:
        messagebox.showerror(
            "Visualize Results",
            "The dashboard needs the full set of output files, but these "
            "are missing from the selected folder:\n  "
            + "\n  ".join(missing_req)
            + "\n\nThey are written together by a model run — please pick a "
              "folder that has the complete set of *.out files.",
            parent=par,
        )
        return

    obs_files = glob.glob(os.path.join(folder, "obs_well*.out"))

    # ── Stage every result file into the active run dir (dashboard cwd) ──
    proj = run_model._run_dir()
    try:
        same = (os.path.normcase(os.path.realpath(folder)) ==
                os.path.normcase(os.path.realpath(proj)))
    except Exception:
        same = (folder == proj)

    if not same:
        for src in _gather_result_files(folder):
            dst = os.path.join(proj, os.path.basename(src))
            try:
                if os.path.exists(dst):
                    try: os.chmod(dst, 0o666)
                    except Exception: pass
                shutil.copy2(src, dst)
            except Exception as exc:
                print(f"[visualize] copy failed {os.path.basename(src)}: {exc}")

    # ── dashboard_state.json: use the folder's copy, else GENERATE it ────
    proj_state = os.path.join(proj, "dashboard_state.json")
    generated = False
    if not os.path.exists(proj_state):
        # Snapshot the current form into state, then dump the context the
        # dashboard needs (years, geometry, PFAA names, well names …).
        try:
            state.snapshot(app)
        except Exception:
            pass
        try:
            run_model._dump_dashboard_state(app, proj)
            generated = os.path.exists(proj_state)
        except Exception as exc:
            messagebox.showerror(
                "Visualize Results",
                "Could not build the run context (dashboard_state.json) "
                f"needed for the plots:\n{exc}",
                parent=par,
            )
            return

    # ── Launch + notify ─────────────────────────────────────────────────
    sheet_name = getattr(app, "active_sheet", "Simple")
    ok, msg = run_model.launch_dashboard(parent=par, sheet_name=sheet_name)

    notes = []
    if not obs_files:
        notes.append("No obs_well*.out found — the monitoring-well "
                     "comparison curves will be empty.")
    if generated:
        notes.append("dashboard_state.json was not in the folder, so it was "
                     "generated from the inputs currently in the form. For "
                     "matching years/labels, Load Data for this model first, "
                     "then Visualize Results again.")

    if ok:
        full = msg + (("\n\nNotes:\n- " + "\n- ".join(notes)) if notes else "")
        messagebox.showinfo("Visualize Results", full, parent=par)
    else:
        messagebox.showerror("Visualize Results", msg, parent=par)
