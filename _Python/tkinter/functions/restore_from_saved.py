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
from . import sidecars


def _clean_work_dir(work_dir: str, saved_dir: str):
    """v106: remove *stale* sidecars from work_dir before copying the
    saved ones in, so values from a previously-loaded model can't leak
    into the freshly-loaded one.

    CRITICAL SAFETY (v106 fix): if the folder being restored IS the work
    dir, do NOTHING — otherwise we'd delete the very files we're about to
    read.  And we only ever remove a work_dir sidecar that is ABSENT from
    the saved folder; any sidecar present in the saved folder is left
    alone (it gets overwritten by _copy_files anyway)."""
    try:
        same = (os.path.normcase(os.path.realpath(work_dir)) ==
                os.path.normcase(os.path.realpath(saved_dir)))
    except Exception:
        same = (work_dir == saved_dir)
    if same:
        return
    for fname in INPUT_TXT_FILES:
        wp = os.path.join(work_dir, fname)
        sp = os.path.join(saved_dir, fname)
        try:
            if os.path.exists(wp) and not os.path.exists(sp):
                os.remove(wp)
        except Exception as e:
            print(f"Warning: could not clean {fname}: {e}")


def _copy_files(src_dir: str, dst_dir: str) -> list:
    copied = []
    # Copy input.inp too so subsequent runs in work_dir use the restored
    # model, not a stale input.inp left from a prior session.
    for fname in ["input.inp", *INPUT_TXT_FILES]:
        src = os.path.join(src_dir, fname)
        dst = os.path.join(dst_dir, fname)
        if os.path.exists(src):
            try:
                if os.path.exists(dst) and os.path.samefile(src, dst):
                    copied.append(fname)
                    continue
                # v107: chmod read-only destinations before overwrite —
                # a PermissionError here used to leave the stale file in
                # place silently (root cause of "cell sizes not loading").
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

    # v107: ADOPT the loaded folder as the ACTIVE work_dir — mirrors
    # Save Data's "one folder per model" behaviour.  Previously Load
    # COPIED the saved sidecars into the old work_dir; one read-only
    # destination file made the copy fail silently and the §-button
    # popups (cell sizes, GW velocity, heterogeneity, …) kept showing
    # the PREVIOUS model's values.  Adopting the folder means every
    # sidecar read below, every popup, and Run Model all use the loaded
    # model's own files — no copying, nothing stale, nothing clobbered.
    state.work_dir = saved_dir
    work_dir = saved_dir

    # (kept for the same-folder edge case: no-ops when work_dir IS the
    # saved folder, and still protects any legacy flow that calls run()
    # with a divergent work_dir)
    _clean_work_dir(work_dir, saved_dir)
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

    # v107: input.inp stores vd × transmissive volume fraction; read the
    # fraction from the SAVED folder's heterogeneity sidecar so the loader
    # can recover the original pre-scaling velocity.
    try:
        from .generate_input_file import _read_hetero
        _, volfrac, _ = _read_hetero(
            os.path.join(saved_dir, "heterogeneity_inputs.txt"))
    except Exception:
        volfrac = 1.0

    # v107: convert lengths using the SAVED model's unit flag when the
    # store_info records one — the pre-load UI unit may differ.
    _uf = additional.get('unit_flag')
    if _uf in (1, 2):
        unit_flag = _uf

    # Write parsed data into state — sets E38/G38/K38/M38 from
    # store_info_additional_input.txt's "PFAA 1:" / "PFAA 2:" lines.
    try:
        write_inp_to_state(state, data, additional, unit_flag, volfrac)
    except Exception as e:
        messagebox.showerror("Error",
            f"Could not apply loaded values:\n{e}")
        return

    # v103: parse retardation_inputs.txt LAST so its "PFAS Names from
    # Excel" block wins when the two sidecars disagree.  This file is
    # the authoritative §5 source (it's what _on_pfaa_change reads to
    # compute R); store_info is a secondary mirror written by Save
    # Data.  Previously the order was reversed and store_info won,
    # which meant any prior calibration glitch that wrote a wrong
    # species into store_info would override the correct
    # retardation_inputs.txt value on Load Data.
    ret_file = os.path.join(work_dir, "retardation_inputs.txt")
    pfas_names = parse_retardation_pfas_names(ret_file)
    # v108: spurious-precursor guard (Ron review item 13).  The
    # retardation_inputs.txt "PFAS Names" block can carry a STALE
    # precursor species (e.g. "PFAA 1-able" inherited from an Example
    # the user started from) even after the user cleared the precursor
    # in the UI.  store_info is rewritten from the live UI on every
    # Save, so when it records a precursor slot (K38 = Precursor 1,
    # M38 = Precursor 2) as None/blank, honour that and do NOT let the
    # stale retardation name re-add the component -- otherwise K38
    # flips ipre=1 and Run Model demands Precursor-1 data that the
    # model does not have.  Only the None case is guarded; when both
    # sources name a species the retardation file still wins (v103).
    _precursor_cleared = set()
    for _addr, _key in (("K38", "precursor1"), ("M38", "precursor2")):
        _pv = additional.get(_key)
        if _pv is not None and str(_pv).strip().lower() in ("", "none"):
            _precursor_cleared.add(_addr)
    for addr, name in pfas_names.items():
        if addr in _precursor_cleared:
            state.set(addr, "None")   # UI cleared it -> keep it cleared
            continue
        if name and str(name).strip():
            state.set(addr, name)

    # v102: temporarily disable §7 auto-year-interpolator so the
    # values loaded from input.inp / store_info don't get overwritten
    # by the trace that fires when v_yr_start / v_yr_end change.
    # v106: ALSO hold the §6 heterogeneity trace guard (_disp_loading)
    # across the whole push.  push() sets v_het, and for "Enter Your Own
    # Value Below" the het trace would CLEAR the alpha cells that push
    # just recovered from input.inp.  Guarding here means those recovered
    # alphas survive as a fallback even when the dispersivity sidecar is
    # missing (e.g. an older save); when the sidecar IS present,
    # read_dispersivity below overwrites them with the exact values.
    setattr(app, "_s7_years_filling", True)
    setattr(app, "_disp_loading", True)
    try:
        # Push to UI
        state.push(app)
    finally:
        setattr(app, "_s7_years_filling", False)

    # v106: apply the dedicated §6 dispersivity + §9 PSB sidecars AFTER
    # push so their exact values win over the lossy input.inp recovery.
    # _disp_loading is still True (held from before push) so the het trace
    # can't clear the alpha cells during read_dispersivity either.
    try:
        sidecars.read_dispersivity(app, work_dir)
    except Exception as exc:
        print(f"[restore] dispersivity sidecar apply failed: {exc}")
    finally:
        setattr(app, "_disp_loading", False)
    try:
        sidecars.read_psb(app, work_dir)
    except Exception as exc:
        print(f"[restore] PSB sidecar apply failed: {exc}")

    # v107: §5 mol-diff fallback — when the dispersivity sidecar is
    # missing (pre-v106 save) or lacks the line, restore the value that
    # inp_to_state recovered from input.inp (state E44) instead of
    # letting the §5 trace revert it to the species-table default.
    try:
        sidecars.mol_diff_from_state_fallback(app, state, work_dir)
    except Exception as exc:
        print(f"[restore] mol-diff fallback failed: {exc}")

    # Re-run §9 Converted-Kf auto-formula (state.push wipes V26..AB26)
    if hasattr(app, '_recompute_psb_conv_kf'):
        try:
            app._recompute_psb_conv_kf()
        except Exception:
            pass

    # Recompute section 5 (retardation factors + mol-diff) from PFAS names
    if hasattr(app, '_on_pfaa_change'):
        try:
            app._on_pfaa_change()
        except Exception as exc:
            print(f"[restore] _on_pfaa_change failed: {exc}")

    # Re-seed §7 baseline from the freshly-loaded values so the §8 →
    # §7 auto-apply trace works against the loaded concentrations.
    if hasattr(app, '_refresh_s7_baseline'):
        try:
            app._refresh_s7_baseline()
        except Exception:
            pass

    messagebox.showinfo("Success",
        f"Data restored from:\n{os.path.basename(saved_dir)}\n\n"
        "This is now your active model folder — Run Model will write "
        "results here, and the section buttons (cell sizes, GW velocity, "
        "heterogeneity, …) read this folder's inputs.")
