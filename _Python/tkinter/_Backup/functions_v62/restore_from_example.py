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

    example_dir = os.path.join(work_dir, "Example", sub)
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

    # Push to UI
    state.push(app)

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
    # §calibration quadrant — Step 3 weig