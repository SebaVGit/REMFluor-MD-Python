"""
restore_from_saved.py — standalone replacement for xlwings version.

Asks user to pick a saved folder, copies .txt files, parses input.inp,
writes values into state and pushes to UI.
"""
import os
import shutil
from tkinter import messagebox, filedialog

from .state import get_state, INPUT_TXT_FILES
from .inp_parser import parse_input_inp, parse_additional_info, parse_retardation_pfas_names
from .inp_to_state import write_inp_to_state


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
    """Restore from a user-selected saved folder."""
    state = get_state()
    state.snapshot(app)

    work_dir = state.work_dir or os.getcwd()
    unit_flag = state.get("AD1", 2)

    # Ask user to pick folder
    saved_dir = filedialog.askdirectory(
        title="Select saved folder to restore from",
        initialdir=work_dir,
    )
    if not saved_dir:
        return   # user cancelled

    inp_file = os.path.join(saved_dir, "input.inp")
    if not os.path.exists(inp_file):
        messagebox.showerror("Error",
            f"input.inp not found in selected folder:\n{saved_dir}")
        return

    # Copy .txt files to work dir
    _copy_files(saved_dir, work_dir)

    # Parse input.inp
    try:
        data = parse_input_inp(inp_file)
    except Exception as e:
        messagebox.showerror("Error", f"Could not parse input.inp:\n{e}")
        return

    # Parse additional info
    add_file = os.path.join(saved_dir, "store_info_additional_input.txt")
    additional = parse_additional_info(add_file) if os.path.exists(add_file) else {}

    # Parse and set PFAS names
    ret_file = os.path.join(work_dir, "retardation_inputs.txt")
    pfas_names = parse_retardation_pfas_names(ret_file)
    for addr, name in pfas_names.items():
        state.set(addr, name)

    # Write parsed data into state
    write_inp_to_state(state, data, additional, unit_flag)

    # Push to UI
    state.push(app)

    # Re-run §9 Converted-Kf auto-formula (state.push wipes V26..AB26)
    if hasattr(app, '_recompute_psb_conv_kf'):
        try:
            app._recompute_psb_conv_kf()
        except Exception:
            pass

    # Recompute section 5 (retardation factors + mol-diff) from PFAS names
    if hasattr(app, '_on_pfaa_change'):
        app._on_pfaa_change()

    # Re-seed §7 baseline from the freshly-loaded values so the §8 →
    # §7 auto-apply trace works against the loaded concentrations.
    if hasattr(app, '_refresh_s7_baseline'):
        try:
            app._refresh_s7_baseline()
        except Exception:
            pass

    messagebox.showinfo("Success",
        f"Data restored from:\n{os.path.basename(saved_dir)}")
