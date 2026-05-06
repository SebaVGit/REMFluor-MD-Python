"""
clear_for_restore.py — standalone replacement for the xlwings-based version.

Clears state cells and deletes input .txt files from the working directory.
Called by: Clear_Data, and as step 1 of Load_Data / Paste_Example.
"""
import os
from .state import get_state, INPUT_TXT_FILES


def run(app) -> list:
    """
    Clear all model cells in state and delete .txt input files.
    Returns list of deleted filenames.
    """
    state = get_state()
    state.snapshot(app)

    # Zero all restore-related cells
    state.clear_restore_cells()

    # Delete .txt files from working directory
    work_dir = state.work_dir or os.getcwd()
    deleted = []
    for fname in INPUT_TXT_FILES:
        fpath = os.path.join(work_dir, fname)
        if os.path.exists(fpath):
            try:
                os.chmod(fpath, 0o666)
                os.remove(fpath)
                deleted.append(fname)
                print(f"Deleted: {fname}")
            except Exception as e:
                print(f"Warning: could not delete {fname}: {e}")

    state.push(app)

    # Recompute section 5 retardation / mol-diff via the existing trace
    # (v_pfaa1 was set to "" by push → _on_pfaa_change fires, but we call
    # it explicitly to guarantee a clean clear in all edge cases)
    if hasattr(app, '_on_pfaa_change'):
        app._on_pfaa_change()

    # Same idea for the §9 Converted-Kf cells.  state.push leaves them
    # blank after a clear; the recompute here propagates the cleared
    # V24/V23 inputs into V26..AB26 (also blank).
    if hasattr(app, '_recompute_psb_conv_kf'):
        try:
            app._recompute_psb_conv_kf()
        except Exception:
            pass

    # Re-seed §7 baseline from the cleared values so the §8 → §7
    # auto-apply trace doesn't keep stale baselines from before
    # the clear.
    if hasattr(app, '_refresh_s7_baseline'):
        try:
            app._refresh_s7_baseline()
        except Exception:
            pass

    # Clear the §c