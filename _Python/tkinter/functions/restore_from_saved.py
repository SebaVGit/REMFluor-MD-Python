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

    # v106: clean stale sidecars in work_dir, then copy the saved set
    # (input.inp + every sidecar) so nothing from a prior model lingers.
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

    # Write parsed data into state — sets E38/G38/K38/M38 from
    # store_info_additional_input.txt's "PFAA 1:" / "PFAA 2:" lines.
    try:
        write_inp_to_state(state, data, additional, unit_flag)
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
    for addr, name in pfas_names.items():
        if name and str(name).strip():
            state.set(addr, name)

    # v102: temporarily disable §7 auto-year-interpolator so the
    # values loaded from input.inp / store_info don't get overwritten
    # by the trace that fires when v_yr_start / v_yr_end change.
    setattr(app, "_s7_years_filling", True)
    try:
        # Push to UI
        state.push(app)
    finally:
        setattr(app, "_s7_years_filling", False)

    # v106: apply the dedicated §6 dispersivity + §9 PSB sidecars AFTER
    # push so their exact values win over the lossy input.inp recovery.
    #   - Dispersivity: guard the §6 het trace so custom "Enter Your Own
    #     Value" alpha cells aren't re-derived / cleared.
    #   - PSB: sets raw Kf + unit dropdown; the converted-Kf recompute
    #     below then rebuilds V26..AB26 with the correct units.
    setattr(app, "_disp_loading", True)
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
        f"Data restored from:\n{os.path.basename(saved_dir)}")
