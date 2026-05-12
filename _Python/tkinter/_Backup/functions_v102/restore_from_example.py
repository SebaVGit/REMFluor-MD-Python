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

    # v100: search both work_dir (writable, dev tree) and bundle_dir
    # (read-only PyInstaller _MEIPASS) for the Example/ folder.  In a
    # frozen .exe the bundled Example/ lives in _MEIPASS, not next to
    # the .exe itself, so the work_dir-only lookup used to fail.
    example_dir = os.path.join(work_dir, "Example", sub)
    if not os.path.isdir(example_dir):
        bundle_dir = getattr(state, "bundle_dir", "") or work_dir
        alt = os.path.join(bundle_dir, "Example", sub)
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

    # Write parsed data into state
    write_inp_to_state(state, data, additional, unit_flag)

    # v102: temporarily disable §7 auto-year-interpolator so the
    # values loaded from input.inp / store_info don't get overwritten
    # by the trace that fires when v_yr_start / v_yr_end change.
    setattr(app, "_s7_years_filling", True)
    try:
        # Push to UI
        state.push(app)
    finally:
        setattr(app, "_s7_years_filling", False)

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
            for base in (work_dir, getattr(state, "bundle_dir", "") or ""):
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
