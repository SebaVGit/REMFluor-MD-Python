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

    # v102: blank the molecular-diffusion StringVar AFTER the trace runs.
    # v_pfaa1 (E38) is intentionally preserved on Clear All Data (per the
    # Excel macro behaviour — dropdown selections survive a clear), so
    # _on_pfaa_change above sees "PFOS" still in the dropdown and
    # re-populates v_mol_diff with PFOS's diffusion value.  Override
    # that by clearing v_mol_diff as the FINAL step.  This gives the
    # user the visible "clean slate" they expect from Clear All Data —
    # the cell stays empty until they explicitly change a §5 dropdown
    # (which fires _on_pfaa_change again with a real species).
    try:
        if hasattr(app, 'v_mol_diff'):
            app.v_mol_diff.set("")
    except Exception:
        pass

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

    # Clear the §calibration quadrant — all weights → 1.0, all param
    # ranges blank, all "Use this Parameter?" checkboxes off, Step 2
    # checkboxes off.  Without this, Clear All Data left the
    # calibration panel still showing the previously-loaded example
    # weights / ranges, which was confusing.
    try:
        for v in getattr(app, "v_calib_w", []):
            try: v.set("1.0")
            except Exception: pass
        for v in getattr(app, "v_calib_chk", []):
            try: v.set(False)
            except Exception: pass
        for v in (getattr(app, "v_calib_low", [])
                  + getattr(app, "v_calib_high", [])):
            try: v.set("")
            except Exception: pass
        # Mid-Range column auto-populates from the source cells via
        # the deferential _refresh_calib_mids, but rows without a §
        # StringVar (i, multipliers) need the "1" baseline seeded
        # explicitly because the refresh leaves them alone.
        BASELINE_MID = {
            "Hydraulic Gradient (i)":                                            "1",
            "Multiplier to PFAA-1 Source Concentration in #7 (czero(2,n))":      "1",
            "Multiplier to PFAA-2 Source Concentration in #7 (czero(4,n))":      "1",
            "Multiplier to Precursor-1 Source Concentration in #7 (czero(1,n))": "1",
            "Multiplier to Precursor-2 Source Concentration in #7 (czero(3,n))": "1",
        }
        try:
            from main import _CALIB_PARAMS  # type: ignore
        except Exception:
            _CALIB_PARAMS = []
        mids = getattr(app, "v_calib_mid", [])
        for i, lbl in enumerate(_CALIB_PARAMS):
            if i >= len(mids):
                break
            try:
                mids[i].set(BASELINE_MID.get(lbl, ""))
            except Exception:
                pass
        for name in ("v_calib_pfoa", "v_calib_none",
                     "v_calib_pre", "v_calib_pre_none"):
            v = getattr(app, name, None)
            if v is not None:
                try: v.set(False)
                except Exception: pass
    except Exception as exc:
        print(f"[clear_for_restore] calibration clear failed: {exc}")

    return deleted
