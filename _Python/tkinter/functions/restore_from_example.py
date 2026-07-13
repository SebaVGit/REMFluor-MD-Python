"""
restore_from_example.py — standalone replacement for xlwings version.

1. Determines example subfolder from version + unit flags in state.
2. Copies .txt files to work_dir.
3. Parses input.inp + additional info.
4. Writes values into state and pushes to UI.
"""
import os
import shutil
from tkinter import messagebox

from .state import get_state, INPUT_TXT_FILES
from .inp_parser import parse_input_inp, parse_additional_info, parse_retardation_pfas_names
from .inp_to_state import write_inp_to_state

_EXAMPLE_SUBFOLDERS = {
    (1, 2): "1_Simple",
    (2, 2): "3_Detailed_2",
    (1, 1): "Simple_ft",
    (2, 1): "Detailed_ft",
}


def _copy_files(src_dir: str, dst_dir: str) -> list:
    copied = []
    for fname in INPUT_TXT_FILES:
        src = os.path.join(src_dir, fname)
        dst = os.path.join(dst_dir, fname)
        if os.path.exists(src):
            try:
                # v107: Windows leaves result/sidecar files read-only at
                # times; copy2 onto a read-only dst raises PermissionError
                # and the stale file silently survived.  chmod first.
                if os.path.exists(dst):
                    try: os.chmod(dst, 0o666)
                    except Exception: pass
                shutil.copy2(src, dst)
                copied.append(fname)
                print(f"Copied: {fname}")
            except Exception as e:
                print(f"Warning: could not copy {fname}: {e}")
    return copied


def run(app):
    """Restore from Example/ subfolder matching current version + unit flags."""
    state = get_state()
    state.snapshot(app)

    version_flag = state.get("A8", 2)
    unit_flag    = state.get("AD1", 2)
    work_dir     = state.work_dir or os.getcwd()

    sub = _EXAMPLE_SUBFOLDERS.get((version_flag, unit_flag))
    if sub is None:
        messagebox.showerror("Error",
            "Cannot determine example folder.\n"
            "Check version and units selection.")
        return

    # v107: the Example/ folder ships NEXT TO the .exe (base_dir) —
    # it is NOT bundled inside the .exe.  When the user saves/loads a
    # model, work_dir switches to that model folder, which has no
    # Example/, so keep the reference to the .exe folder.  Search
    # order: work_dir (dev tree / user override) → base_dir (exe
    # folder) → bundle_dir (_MEIPASS, legacy fallback).
    base_dir   = getattr(state, "base_dir",   "") or ""
    bundle_dir = getattr(state, "bundle_dir", "") or ""
    example_dir = os.path.join(work_dir, "Example", sub)
    for cand_base in (base_dir, bundle_dir):
        if os.path.isdir(example_dir):
            break
        if cand_base:
            alt = os.path.join(cand_base, "Example", sub)
            if os.path.isdir(alt):
                example_dir = alt
    if not os.path.isdir(example_dir):
        messagebox.showerror("Error",
            f"Example folder not found:\n{example_dir}")
        return

    inp_file = os.path.join(example_dir, "input.inp")
    if not os.path.exists(inp_file):
        messagebox.showerror("Error",
            f"input.inp not found in:\n{example_dir}")
        return

    # v107: Paste Example ALWAYS runs in the app's BASE folder with a
    # clean slate, regardless of which model folder was loaded before.
    # Two guarantees: (1) example results are reproducible — identical
    # every time, because no leftover sidecar from a previously-loaded
    # model can skew the run; (2) a user's model folder is never
    # contaminated with example files or example results.
    _base = getattr(state, "base_dir", "") or ""
    if _base and os.path.isdir(_base):
        state.work_dir = _base
        work_dir = _base

    # Remove stale sidecars the example does NOT ship (dispersivity,
    # PSB, transformation, calibration leftovers, …) so the run reads
    # only example-provided or default values.
    for fname in list(INPUT_TXT_FILES) + ["mw_observations.json"]:
        wp = os.path.join(work_dir, fname)
        sp = os.path.join(example_dir, fname)
        try:
            if os.path.exists(wp) and not os.path.exists(sp):
                os.remove(wp)
        except Exception as e:
            print(f"[restore_from_example] could not clean {fname}: {e}")

    # Copy auxiliary .txt files to work dir
    _copy_files(example_dir, work_dir)

    # Parse input.inp
    try:
        data = parse_input_inp(inp_file)
    except Exception as e:
        messagebox.showerror("Error", f"Could not parse input.inp:\n{e}")
        return

    # Parse additional info
    add_file = os.path.join(example_dir, "store_info_additional_input.txt")
    additional = parse_additional_info(add_file) if os.path.exists(add_file) else {}

    # Parse and set PFAS names from retardation file
    ret_file = os.path.join(work_dir, "retardation_inputs.txt")
    pfas_names = parse_retardation_pfas_names(ret_file)
    for addr, name in pfas_names.items():
        state.set(addr, name)

    # v107: input.inp stores vd × transmissive volume fraction; read the
    # fraction from the example's heterogeneity sidecar so the loader can
    # recover the original pre-scaling velocity.
    try:
        from .generate_input_file import _read_hetero
        _, volfrac, _ = _read_hetero(
            os.path.join(example_dir, "heterogeneity_inputs.txt"))
    except Exception:
        volfrac = 1.0

    # Write parsed data into state
    write_inp_to_state(state, data, additional, unit_flag, volfrac)

    # v102: temporarily disable §7 auto-year-interpolator so the
    # values loaded from input.inp / store_info don't get overwritten
    # by the trace that fires when v_yr_start / v_yr_end change.
    # v107: ALSO hold the §6 heterogeneity trace guard across the push
    # (same as restore_from_saved).  push() sets v_het last, and the
    # _on_het_change trace re-derives the three alpha cells from the
    # PRESET table — overwriting the alphax/y/z values just loaded from
    # the example's input.inp (e.g. UI on "High" → 7.5 replaced the
    # example's 3.2).
    setattr(app, "_s7_years_filling", True)
    setattr(app, "_disp_loading", True)
    try:
        # Push to UI
        state.push(app)
    finally:
        setattr(app, "_s7_years_filling", False)
        setattr(app, "_disp_loading", False)

    # v107: §5 mol-diff — examples ship no dispersivity sidecar, so
    # restore the coefficient recovered from the example's input.inp
    # (state E44) before the §5 trace fills the table default.
    try:
        from . import sidecars as _sc
        _sc.mol_diff_from_state_fallback(app, state, example_dir)
    except Exception as exc:
        print(f"[restore_from_example] mol-diff fallback failed: {exc}")

    # Recompute section 5 (retardation factors + mol-diff) from PFAS names
    # that were just pushed. _on_pfaa_change uses PFAA_KOC lookup table.
    if hasattr(app, '_on_pfaa_change'):
        app._on_pfaa_change()

    # Re-run §9 Converted-Kf auto-formula AFTER state.push.  state.push
    # iterates CELL_MAP and ends up overwriting V26..AB26 (the Converted
    # Kf cells) with "" because state has no value for them — they're
    # populated by the trace_add on V24/U24/V23.  Calling the recompute
    # explicitly here re-fills them from the now-final V24/U24/V23.
    if hasattr(app, '_recompute_psb_conv_kf'):
        try:
            app._recompute_psb_conv_kf()
        except Exception:
            pass

    # Re-seed §7 baseline from the freshly-loaded values so the §8 →
    # §7 auto-apply trace works against the example's pristine
    # concentrations (and not whatever was in §7 before paste).
    if hasattr(app, '_refresh_s7_baseline'):
        try:
            app._refresh_s7_baseline()
        except Exception:
            pass

    # Load the calibration template (Simple / Detailed) into the
    # §calibration quadrant — Step 3 weights, Step 4 ranges, Step 2
    # checkboxes.  Without this Paste Example left the calibration
    # panel untouched, which was confusing for users who expected
    # the example's full setup.
    try:
        from . import restore_calibration
        restore_calibration.run(app)
    except Exception as exc:
        print(f"[restore_from_example] calibration load failed: {exc}")

    # v102: Also auto-import per-well screen depths from the
    # calibration template .xlsx (CalibrationTemplate_Detailed.xlsx /
    # _Simple.xlsx) so the next Run Model picks up real per-well
    # zwelltop/zwellbot from Model Location columns 4/5 — same flow
    # as if the user had clicked the §calibration data loader button.
    # Skips quietly if no calibration .xlsx is reachable.
    try:
        # Look for the calibration .xlsx alongside the example folder.
        # restore_calibration writes its path into calibration_inputs.txt.
        cal_txt = os.path.join(work_dir, "calibration_inputs.txt")
        cal_xlsx = None
        if os.path.exists(cal_txt):
            with open(cal_txt, encoding="utf-8", errors="replace") as fh:
                for ln in fh:
                    if ln.lstrip().startswith("Excel File Path:"):
                        cand = ln.split(":", 1)[1].strip()
                        if cand and os.path.isfile(cand):
                            cal_xlsx = cand
                        break
        if cal_xlsx is None:
            is_detailed = (state.get("A8", 2) == 2)
            tmpl = ("CalibrationTemplate_Detailed.xlsx" if is_detailed
                    else "CalibrationTemplate_Simple.xlsx")
            for base in (work_dir,
                         getattr(state, "base_dir", "") or "",
                         getattr(state, "bundle_dir", "") or ""):
                if not base:
                    continue
                cand = os.path.join(base, tmpl)
                if os.path.isfile(cand):
                    cal_xlsx = cand
                    break
        if cal_xlsx:
            from . import popups_calibration
            try:
                status = popups_calibration._import_xlsx_into_app(app, cal_xlsx)
                print(f"[restore_from_example] xlsx import: {status}")
            except Exception as exc:
                print(f"[restore_from_example] xlsx import failed: {exc}")
    except Exception as exc:
        print(f"[restore_from_example] xlsx auto-import failed: {exc}")

    # Re-pull the Step 4 Mid-Range column from the freshly-loaded
    # source cells (§3 vd / porf, §5 retardation, §2 years, etc.).
    if hasattr(app, '_refresh_calib_mids'):
        try:
            app._refresh_calib_mids()
        except Exception:
            pass

    messagebox.showinfo("Success",
        f"Example data loaded.\nFolder: {os.path.basename(example_dir)}")
