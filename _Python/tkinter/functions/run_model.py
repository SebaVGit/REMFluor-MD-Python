"""
run_model.py — Run Model pipeline (replaces the broken inline launcher).

Mirrors the Source_Py/input_variables.py __main__ block:

    1.  Build/refresh input.inp via generate_input_file.run(app)
    2.  Run remfluor_v8a.exe in a shell with stdin/stdout redirection:
            remfluor_v8a.exe < input.inp > output.out
    3.  Show a tkinter Toplevel with a live "Runtime Clock" while the
        Fortran solver runs.
    4.  When the model finishes, launch the Plotly-Dash dashboard
        (functions/generate_dashboard.py) — the dashboard auto-opens
        the user's default browser (Edge/Chrome) at http://localhost:8050.

Called from main.run_script() when RunPythonScript is dispatched.
"""
from __future__ import annotations
import os
import sys
import json
import time
import atexit
import threading
import subprocess
import tkinter as tk
from tkinter import messagebox

from . import generate_input_file


# Track the most-recently-spawned dashboard subprocess so we can
# terminate it before launching a new one — Plotly-Dash binds to a
# fixed port (8050) and a stale server would silently hold the port,
# which lets the new subprocess crash and leaves the user looking at
# the OLD dashboard data.  See _kill_previous_dashboard().
_LAST_DASHBOARD_PROC = None
# Uncommon base port (8050 is the well-known Dash default that other
# tools grab).  The dashboard scans _DASHBOARD_PORT .. +_DASHBOARD_PORT_SPAN-1
# for the first free one and records the actual choice in
# dashboard_port.txt so cleanup can target it exactly even if the base
# was busy.  Keep these in sync with find_available_port() in
# generate_dashboard.py.
_DASHBOARD_PORT = 8753
_DASHBOARD_PORT_SPAN = 10
_PORT_FILE = "dashboard_port.txt"


def _port_file_path() -> str:
    return os.path.join(_run_dir(), _PORT_FILE)


def _recorded_dashboard_port():
    """Return the port the running dashboard wrote to dashboard_port.txt,
    or None if unavailable/unparseable."""
    try:
        with open(_port_file_path(), "r", encoding="utf-8") as fh:
            return int(fh.read().strip())
    except Exception:
        return None


def _project_dir() -> str:
    """Writable runtime location.  In dev = walked-up __file__.
    In a frozen .exe = the folder containing the .exe."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(here, "..", "..", ".."))


def _run_dir() -> str:
    """The ACTIVE model folder for this run's inputs + outputs.

    v106: defaults to the app folder (_project_dir()), but once the user
    saves a model with Save Data the state work_dir points at that saved
    folder — so each model's input.inp + .out results + dashboard read/
    write in ITS OWN folder, keeping things organized and preventing one
    model's run from overwriting another's.  Bundled read-only assets
    (remfluor_v8a.exe, template.inp) are still resolved via the app
    folder / PyInstaller bundle, so moving the run dir is safe."""
    try:
        from .state import get_state
        wd = getattr(get_state(), "work_dir", "") or ""
        if wd and os.path.isdir(wd):
            return wd
    except Exception:
        pass
    return _project_dir()


def _bundle_dir() -> str:
    """Read-only assets location (template.inp, remfluor_v8a.exe,
    Example/, docs/, Figures/).  In dev = same as project_dir.
    In a frozen .exe = sys._MEIPASS (where PyInstaller --add-data
    files are unpacked)."""
    bundled = getattr(sys, "_MEIPASS", None)
    if bundled and os.path.isdir(bundled):
        return bundled
    return _project_dir()


def _resolve_asset(*parts) -> str:
    """Resolve a bundled read-only file.  Searches _project_dir() first
    (so a user can drop an updated copy next to the .exe to override the
    bundled one), falls back to _bundle_dir().  Returns the
    project-side path even if neither exists, so the caller can give a
    useful 'missing file' error message."""
    rel = os.path.join(*parts)
    proj_cand = os.path.join(_project_dir(), rel)
    if os.path.exists(proj_cand):
        return proj_cand
    bundle_cand = os.path.join(_bundle_dir(), rel)
    if os.path.exists(bundle_cand):
        return bundle_cand
    return proj_cand


def _kill_previous_dashboard():
    """Terminate the previously-spawned dashboard subprocess if it's
    still alive, then free port 8050.  Best-effort — we do NOT raise
    on failure since the new dashboard's own bind will raise a clearer
    error if the port really is stuck."""
    global _LAST_DASHBOARD_PROC
    proc = _LAST_DASHBOARD_PROC
    if proc is not None and proc.poll() is None:
        try:
            proc.terminate()
        except Exception:
            pass
        try:
            proc.wait(timeout=3)
        except Exception:
            try:
                proc.kill()
            except Exception:
                pass
    _LAST_DASHBOARD_PROC = None

    # Belt + suspenders: kill anything else listening on the dashboard
    # port(s).  This catches dashboards left over from a prior app
    # session that crashed (no _LAST_DASHBOARD_PROC handle), or zombies
    # bash didn't reap.  We target the whole scan range AND the exact
    # port the last dashboard recorded in dashboard_port.txt — so even
    # if the base port was busy and it landed on +1/+2, cleanup still
    # finds it.  Cross-platform via psutil if available, else OS fallback.
    target_ports = set(range(_DASHBOARD_PORT,
                             _DASHBOARD_PORT + _DASHBOARD_PORT_SPAN))
    rec = _recorded_dashboard_port()
    if rec:
        target_ports.add(rec)

    try:
        import psutil   # type: ignore
        for c in psutil.net_connections(kind="inet"):
            if c.status != psutil.CONN_LISTEN:
                continue
            laddr = getattr(c, "laddr", None)
            if laddr and laddr.port in target_ports and c.pid:
                try:
                    psutil.Process(c.pid).terminate()
                except Exception:
                    pass
        return
    except Exception:
        pass

    # OS fallback: Windows = netstat | findstr + taskkill
    try:
        if os.name == "nt":
            import re
            ns = subprocess.run(
                ["netstat", "-ano"],
                capture_output=True, text=True, timeout=5,
            )
            ports_alt = "|".join(str(p) for p in sorted(target_ports))
            pattern = re.compile(rf":(?:{ports_alt})\s+.*LISTENING\s+(\d+)")
            seen = set()
            for line in (ns.stdout or "").splitlines():
                m = pattern.search(line)
                if m:
                    pid = m.group(1)
                    if pid != "0" and pid not in seen:
                        seen.add(pid)
                        subprocess.run(
                            ["taskkill", "/F", "/PID", pid],
                            capture_output=True, text=True, timeout=5,
                        )
        else:
            # POSIX: lsof or fuser.  fuser is available on most Linux/macOS.
            for p in sorted(target_ports):
                subprocess.run(
                    ["fuser", "-k", f"{p}/tcp"],
                    capture_output=True, text=True, timeout=5,
            )
    except Exception:
        pass


@atexit.register
def _cleanup_dashboard_on_exit():
    """Terminate the dashboard server when the GUI process exits
    normally (window closed -> mainloop returns -> interpreter exit).

    NOTE: atexit does NOT run on a hard crash or force-quit — that case
    is covered by _kill_previous_dashboard()'s port-8050 sweep on the
    next Run Model, and by the timestamped log names which prevent a
    leaked handle from locking a reused filename.  Together these give
    layered cleanup; no single mechanism is relied on alone."""
    try:
        _kill_previous_dashboard()
    except Exception:
        pass


def _xlsm_path() -> str:
    # Only used to hand the dashboard a DIRECTORY to chdir into (it reads
    # the .out files + dashboard_state.json from there).  Point it at the
    # active run dir so the dashboard reads the current model's results.
    return os.path.join(_run_dir(),
                        "REMFluor-MD Interface Storyboard v2.6.xlsm")


def _dashboard_script() -> str:
    # Lives next to this file (tkinter/functions/) — kept in functions/
    # for consistency with every other ported module.  Source_Py copy
    # remains in the repo as historical reference only.
    return os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "generate_dashboard.py")


def _prune_dashboard_logs(proj: str, keep: int = 10):
    """Keep only the newest `keep` dashboard_*.log files.  Best-effort:
    a file still locked by a running/zombie dashboard simply won't
    delete, and that's fine — the next run writes a fresh timestamped
    name regardless.  Also removes a legacy fixed-name dashboard.log if
    present and unlocked."""
    import glob as _glob
    # Legacy single-name log from older builds — try to clear it so it
    # stops being the file users couldn't delete.
    legacy = os.path.join(proj, "dashboard.log")
    if os.path.exists(legacy):
        try:
            os.remove(legacy)
        except Exception:
            pass
    try:
        logs = _glob.glob(os.path.join(proj, "dashboard_*.log"))
        logs.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        for old in logs[keep:]:
            try:
                os.remove(old)
            except Exception:
                pass
    except Exception:
        pass


def _dump_dashboard_state(app, project: str) -> str:
    """Write a JSON snapshot of every state cell the dashboard needs,
    with HARD-CODED FALLBACK DEFAULTS so the dashboard NEVER sees a
    None value in any field it touches.

    v86: full standalone — replaces openpyxl/win32com .xlsm reads in
    the dashboard.  The previous v82 attempt just dumped raw state
    (which is None for unfilled fields), then the dashboard tried to
    cope with None at read time — that whack-a-mole'd through three
    versions.  Now defaults live HERE, at the dump.  The dashboard
    just trusts the JSON and never has to handle None.
    """
    # v89: state is a singleton (state.get_state()), NOT an attribute
    # on the app object.  Previous versions read getattr(app, "state",
    # None) which always returned None — so every cell fell through to
    # the hard-coded defaults below, which is why the JSON dump (and
    # therefore the whole dashboard) ignored the user's form values.
    from .state import get_state
    state = get_state()
    # generate_input_file.run(app) above already snapshotted the form
    # into state, so state has the user's actual values now.

    def _raw(addr):
        try:
            return state.get(addr) if state is not None else None
        except Exception:
            return None

    def _is_blank(v):
        return v is None or (isinstance(v, str) and not str(v).strip()) \
            or (isinstance(v, str) and str(v).strip().lower() == "none")

    def _num(addr, default):
        """Numeric cell with default.  Strips commas, tolerates strings."""
        v = _raw(addr)
        if _is_blank(v):
            return default
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return v
        try:
            return float(str(v).replace(",", "").strip())
        except (TypeError, ValueError):
            return default

    def _int(addr, default):
        v = _num(addr, default)
        try:
            return int(round(float(v)))
        except Exception:
            return default

    def _str(addr, default):
        v = _raw(addr)
        if _is_blank(v):
            return default
        return str(v).strip()

    def _list(start_row, end_row, col, default_factory):
        """Read a column range, replacing blanks with default_factory(i)."""
        out = []
        for i, r in enumerate(range(start_row, end_row + 1)):
            v = _raw(f"{col}{r}")
            if _is_blank(v):
                out.append(default_factory(i))
            else:
                # Try to coerce to number, else keep as string
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    out.append(v)
                else:
                    s = str(v).strip()
                    try:
                        out.append(float(s.replace(",", "")))
                    except (TypeError, ValueError):
                        out.append(s)
        return out

    # v87: defaults below MUST match the form's StringVar initial
    # values in main.py (v_yr_start="1977", v_yr_end="2077",
    # v_sw_width="60", v_sw_thick="5", v_z_size="10", v_see_every="10").
    # When the form is pre-filled (which it is on app startup), state
    # has those values and these defaults aren't used; but if anything
    # races / clears state, the dashboard still matches what the user
    # would see in the form.
    start_year = _int("E18", 1977)
    end_year   = _int("E19", 2077)

    payload = {
        # ── Scalars used for plot context ──────────────────────────
        "A8":   _int("A8", 1),       # version flag (1=Simple, 2=Detailed)
        "AD1":  _int("AD1", 1),      # unit flag (1=feet, 2=meters)
        "K38":  _str("K38", ""),     # Precursor 1 name (blank = ipre=0)
        "E38":  _str("E38", "PFAA-1"),
        "M38":  _str("M38", ""),     # Precursor 2 name
        "G38":  _str("G38", ""),     # PFAA 2 name (blank = ncomp=1)

        # See Results Every (yr) — int avoids modulo issues
        "V47":  _int("V47", 10),

        # Time
        "E18":  start_year,
        "E19":  end_year,

        # Geometry / source — match form defaults
        "E13":  _num("E13", 10.0),   # Total depth   (v_z_size)
        "E15":  _num("E15", 60.0),   # Source width  (v_sw_width)
        "E16":  _num("E16", 5.0),    # Source thick  (v_sw_thick)
        "E22":  _num("E22", 1.0),    # Velocity (Darcy)

        # Sample years (only one is used depending on version_flag).
        # v101: v_sample_yr only maps to Y74 in CELL_MAP, so R36 is
        # always blank in state for Simple users.  Default R36 to Y74
        # (the user's actual §10 Sample Year input) so the dashboard's
        # Simple-mode read of R36 doesn't fall through to start_year
        # — which made obs markers plot at the simulation start year
        # instead of the user-specified sample year.
        "R36":  _int("R36", _int("Y74", start_year)),
        "Y74":  _int("Y74", start_year),

        # ── Ranges ──────────────────────────────────────────────────
        # Source loading times: monotone integers if state empty
        "U8:U18":  _list(8, 18, "U", lambda i: i),
        # Source concs: 0 by default
        "V8:V18":  _list(8, 18, "V", lambda i: 0.0),
        "X8:X18":  _list(8, 18, "X", lambda i: 0.0),
        "Z8:Z18":  _list(8, 18, "Z", lambda i: 0.0),
        "AB8:AB18":_list(8, 18, "AB", lambda i: 0.0),

        # Monitoring well names: placeholder if blank
        "U34:U40": _list(34, 40, "U", lambda i: f"Well {i+1}"),
        # MW concentrations: blank = no observed data point
        "V34:V40": _list(34, 40, "V", lambda i: 0.0),
        "X34:X40": _list(34, 40, "X", lambda i: 0.0),
    }

    out_path = os.path.join(project, "dashboard_state.json")
    try:
        with open(out_path, "w", encoding="utf-8") as fh:
            json.dump(payload, fh, default=str, indent=2)
    except Exception as e:
        try:
            print(f"[run_model] could not write dashboard_state.json: {e}")
        except Exception:
            pass
    return out_path


def _check_dashboard_imports(py_exec: str):
    """Run a tiny child-Python import test for everything
    generate_dashboard.py needs.  Returns (ok, missing_modules_list).
    Fast — typically <0.5 s.

    v99: skipped when running as a frozen .exe — PyInstaller has
    already bundled every required module and `[py_exec, "-c", ...]`
    would re-launch the .exe with "-c" argv (which the multi-mode
    dispatcher doesn't recognise → confused main()).  Validating
    the frozen bundle's imports up-front would also have to do an
    in-process import test, which we trust the bundle to satisfy.
    """
    if getattr(sys, "frozen", False):
        return True, []
    test_code = (
        "import importlib, sys\n"
        "missing = []\n"
        "for m in ('pandas','numpy','plotly','dash'):\n"
        "    try: importlib.import_module(m)\n"
        "    except Exception: missing.append(m)\n"
        "print('MISSING:' + ','.join(missing))\n"
    )
    try:
        out = subprocess.run(
            [py_exec, "-c", test_code],
            capture_output=True, text=True, timeout=15,
        )
    except Exception:
        return True, []   # if the probe itself fails, don't block
    txt = (out.stdout or "") + (out.stderr or "")
    for line in txt.splitlines():
        if line.startswith("MISSING:"):
            miss = [m for m in line[len("MISSING:"):].split(",") if m]
            return (len(miss) == 0), miss
    return True, []


def _launch_dashboard_async(workbook_path: str, sheet_name: str,
                            label: tk.Label, root: tk.Toplevel,
                            mins: int, secs: int):
    """Spawn `python generate_dashboard.py <xlsm> <sheet>` in a daemon
    thread.  Plotly-Dash opens the user's default browser to the
    dashboard URL.

    Captures stdout/stderr to dashboard.log in the project dir so we
    can diagnose launch failures (missing pywin32 / dash / plotly).
    Polls the child for ~5 s; if it exits with non-zero, surfaces the
    log tail in the runtime-clock label and a messagebox.
    """
    def _worker():
        script = _dashboard_script()
        proj   = _run_dir()
        # Timestamped log per run so a leftover OS handle (e.g. after a
        # crash) can never lock the fixed name and block deletion/rebuild.
        # _prune_dashboard_logs() trims old ones.  See _prune_dashboard_logs.
        log_path = os.path.join(
            proj, f"dashboard_{time.strftime('%Y%m%d_%H%M%S')}.log")
        try:
            _prune_dashboard_logs(proj, keep=10)
        except Exception:
            pass

        # In a frozen .exe the dashboard is launched via the
        # `--mode=dashboard` re-exec of the .exe itself (see below), so
        # the .py source need NOT exist on disk — and it doesn't, because
        # PyInstaller compiles it into the bundle.  Only require the
        # on-disk script when running from source (dev mode).
        if not getattr(sys, "frozen", False) and not os.path.exists(script):
            try:
                root.after(0, lambda: label.config(
                    text=f"Finished: {mins:02d}:{secs:02d}\n"
                         f"Dashboard script not found:\n{script}"))
            except Exception:
                pass
            return

        try:
            root.after(0, lambda: label.config(
                text=f"Finished: {mins:02d}:{secs:02d}\n"
                     "Launching Dashboard..."))
        except Exception:
            pass

        # Use the same Python interpreter the app is running under.
        py = sys.executable or "python"

        # Pre-flight: verify all dashboard imports resolve in this env.
        ok, missing = _check_dashboard_imports(py)
        if not ok:
            pip_names = {
                "win32com.client": "pywin32",
                "pythoncom":       "pywin32",
            }
            pkgs = sorted({pip_names.get(m, m) for m in missing})
            install = "pip install " + " ".join(pkgs)
            err_short = (f"Finished: {mins:02d}:{secs:02d}\n"
                         f"Dashboard NOT launched.\n"
                         f"Missing modules: {', '.join(missing)}\n"
                         f"Run:  {install}")
            err_full = (
                "The Plotly-Dash dashboard cannot start because the "
                "current Python environment is missing modules.\n\n"
                f"Missing imports: {', '.join(missing)}\n\n"
                f"To fix, open a terminal in the `remfluor` conda env and run:\n"
                f"    {install}\n\n"
                f"Then click Run Model again."
            )
            try:
                root.after(0, lambda: label.config(text=err_short))
                root.after(0, lambda: messagebox.showerror(
                    "Dashboard Dependencies Missing",
                    err_full, parent=root))
            except Exception:
                pass
            return

        # v99: PyInstaller multi-mode dispatcher.
        # In a frozen .exe build, sys.executable IS the .exe — there
        # is no separate python interpreter to invoke `script` with.
        # We instead re-launch the .exe with a `--mode=dashboard`
        # flag, which the dispatcher at the top of main.py
        # intercepts and routes into generate_dashboard.main().
        # In dev mode (sys.frozen is False) we keep the original
        # `[python, -u, script, ...]` invocation.
        if getattr(sys, "frozen", False):
            cmd = [py, "--mode=dashboard", workbook_path, sheet_name]
        else:
            cmd = [py, "-u", script, workbook_path, sheet_name]

        # Kill any prior dashboard process before spawning the new one.
        _kill_previous_dashboard()
        time.sleep(0.5)

        # Open log file for combined stdout+stderr capture.
        try:
            log_fh = open(log_path, "w", encoding="utf-8", errors="replace")
        except Exception:
            log_fh = None

        kwargs = {
            "cwd":    proj,
            "stdout": log_fh if log_fh else subprocess.DEVNULL,
            "stderr": subprocess.STDOUT,
            "stdin":  subprocess.DEVNULL,
        }
        if os.name == "nt":
            kwargs["creationflags"] = 0x00000200
        else:
            kwargs["start_new_session"] = True

        try:
            proc = subprocess.Popen(cmd, **kwargs)
            global _LAST_DASHBOARD_PROC
            _LAST_DASHBOARD_PROC = proc
            # Close the PARENT's copy of the log handle immediately.
            # The child (dashboard) has inherited its own handle and
            # keeps writing; leaving the parent's copy open is what
            # left dashboard.log locked for the GUI's whole lifetime
            # (undeletable file).  See bpo-19575 / bpo-24909.
            if log_fh:
                try:
                    log_fh.close()
                except Exception:
                    pass
                log_fh = None
        except Exception as e:
            if log_fh:
                try: log_fh.close()
                except Exception: pass
            try:
                root.after(0, lambda: label.config(
                    text=f"Finished: {mins:02d}:{secs:02d}\nDashboard error:\n{e}"))
                root.after(0, lambda: messagebox.showerror(
                    "Dashboard Error",
                    f"Failed to launch dashboard:\n{e}\n\nLog: {log_path}",
                    parent=root))
            except Exception:
                pass
            return

        # Poll briefly — dashboards typically take a couple of seconds
        # to start serving; if the process dies in <5 s it's a crash.
        for _ in range(50):  # 5 s @ 100 ms
            if proc.poll() is not None:
                break
            time.sleep(0.1)

        rc = proc.poll()
        if rc is None:
            try:
                root.after(0, lambda: label.config(
                    text=f"Finished in {mins:02d}:{secs:02d}\n"
                         "Dashboard launched — see your browser"))
            except Exception:
                pass
            return

        # Process exited early — read tail of log and show error.
        tail = ""
        try:
            if log_fh:
                try: log_fh.flush()
                except Exception: pass
            with open(log_path, "r", encoding="utf-8", errors="replace") as fp:
                lines = fp.readlines()
                tail = "".join(lines[-20:]).strip()
        except Exception:
            pass

        # v106: user-facing wording.  Lead with what happened in plain
        # terms; keep the exit code + log tail lower down for whoever
        # needs to debug.  Avoid raw "exit 1 / traceback" as the headline.
        msg_short = (f"Model finished in {mins:02d}:{secs:02d}.\n"
                     "The results dashboard couldn't open.\n"
                     "See the pop-up for details.")
        msg_full = (
            "The model run finished, but the results dashboard could "
            "not be opened.\n\n"
            "This usually means the model's output files could not be "
            "read — often because the model did not finish producing a "
            "complete set of results.\n\n"
            "What to try:\n"
            "  1. Run the model again.\n"
            "  2. If it keeps happening, send the log file below to "
            "support.\n\n"
            f"Log file:\n{log_path}\n\n"
            f"--- technical details (dashboard exit code {rc}) ---\n"
            f"{tail or '(empty log)'}"
        )
        try:
            root.after(0, lambda: label.config(text=msg_short))
            root.after(0, lambda: messagebox.showerror(
                "Results Dashboard Could Not Open", msg_full, parent=root))
        except Exception:
            pass

    threading.Thread(target=_worker, daemon=False).start()


# Result files the Plotly-Dash dashboard reads from the project dir.
# REMFluor-MD.out + dashboard_state.json are REQUIRED; the rest enrich
# the plots (observation wells, mass discharge, plume mass).
DASHBOARD_RESULT_FILES = [
    "REMFluor-MD.out",
    "discharge.out",
    "plume_mass.out",
    "output.out",
    "dashboard_state.json",
]
DASHBOARD_RESULT_GLOBS = ["obs_well*.out"]


def launch_dashboard(parent=None, sheet_name="Simple"):
    """Launch the Plotly-Dash dashboard for whatever result files are
    already sitting in the project dir — WITHOUT re-running the solver.

    Used by the §11 "Visualize Results" button.  Returns
    (ok: bool, message: str)."""
    proj = _run_dir()
    main_out = os.path.join(proj, "REMFluor-MD.out")
    if not os.path.exists(main_out):
        return False, ("No model results found to visualize in:\n"
                       f"{proj}\n\n(REMFluor-MD.out is missing.)")

    py = sys.executable or "python"
    ok, missing = _check_dashboard_imports(py)
    if not ok:
        pip_names = {"win32com.client": "pywin32", "pythoncom": "pywin32"}
        pkgs = sorted({pip_names.get(m, m) for m in missing})
        return False, ("The dashboard cannot start — missing Python "
                       f"modules: {', '.join(missing)}\n\n"
                       f"Run:  pip install {' '.join(pkgs)}")

    workbook_path = _xlsm_path()
    if getattr(sys, "frozen", False):
        cmd = [py, "--mode=dashboard", workbook_path, sheet_name]
    else:
        script = _dashboard_script()
        if not os.path.exists(script):
            return False, f"Dashboard script not found:\n{script}"
        cmd = [py, "-u", script, workbook_path, sheet_name]

    _kill_previous_dashboard()
    time.sleep(0.5)

    log_path = os.path.join(
        proj, f"dashboard_{time.strftime('%Y%m%d_%H%M%S')}.log")
    try:
        _prune_dashboard_logs(proj, keep=10)
    except Exception:
        pass
    try:
        log_fh = open(log_path, "w", encoding="utf-8", errors="replace")
    except Exception:
        log_fh = None

    kwargs = {
        "cwd":    proj,
        "stdout": log_fh if log_fh else subprocess.DEVNULL,
        "stderr": subprocess.STDOUT,
        "stdin":  subprocess.DEVNULL,
    }
    if os.name == "nt":
        kwargs["creationflags"] = 0x00000200
    else:
        kwargs["start_new_session"] = True

    try:
        proc = subprocess.Popen(cmd, **kwargs)
        global _LAST_DASHBOARD_PROC
        _LAST_DASHBOARD_PROC = proc
    except Exception as e:
        if log_fh:
            try: log_fh.close()
            except Exception: pass
        return False, f"Failed to launch dashboard:\n{e}\n\nLog: {log_path}"
    finally:
        if log_fh:
            try: log_fh.close()
            except Exception: pass

    return True, ("Dashboard launching — it will open in your default "
                  "browser in a few seconds.")


def run(app, parent=None) -> bool:
    """Run the model pipeline.  Returns True if launch succeeded."""
    # Step 0 (v105): validation gate.  Refuse to run on a partially-filled
    # form and tell the user EXACTLY which required inputs are missing for
    # their current options — no silent fallback to built-in defaults.
    validator = getattr(app, "_collect_missing_inputs", None)
    if callable(validator):
        try:
            missing = validator()
        except Exception:
            missing = []
        if missing:
            try:
                msg = app._format_missing_message(missing)
            except Exception:
                msg = ("Some required inputs are missing. Please complete "
                       "the form before running the model.")
            messagebox.showwarning("Run Model — Missing Inputs", msg,
                                   parent=parent or app)
            return False

    # Step 1: build input.inp from current state ---------------------------
    if not generate_input_file.run(app):
        return False

    # v106: run in the ACTIVE model folder (work_dir) so input.inp and the
    # .out outputs land in the saved model's own folder, not the shared app
    # folder.  generate_input_file already wrote input.inp to work_dir, so
    # this keeps the solver cwd consistent with where input.inp lives.
    project = _run_dir()
    # v100: input.inp is a write target → run dir (writable).
    # remfluor_v8a.exe is a bundled binary → resolve via _resolve_asset
    # which checks the app folder (so a user-supplied override works) then
    # falls back to the PyInstaller bundle (sys._MEIPASS).
    inp_path = os.path.join(project, "input.inp")
    # v107: solver renamed to remfluor_v9a.exe; keep v8a as a fallback so
    # older installs / model folders still run.
    exe_path = _resolve_asset("remfluor_v9a.exe")
    if not os.path.exists(exe_path):
        _old = _resolve_asset("remfluor_v8a.exe")
        if os.path.exists(_old):
            exe_path = _old

    if not os.path.exists(inp_path):
        messagebox.showerror(
            "Run Model",
            f"input.inp not found in:\n{project}\n\n"
            "generate_input_file.run() did not produce the file.")
        return False
    if not os.path.exists(exe_path):
        messagebox.showerror(
            "Run Model",
            f"Model executable not found:\n{exe_path}")
        return False

    # Step 1b (v86): dump dashboard_state.json with bulletproof
    # defaults so the dashboard subprocess has zero .xlsm dependency
    # and never sees a None for any field it touches.
    try:
        _dump_dashboard_state(app, project)
    except Exception:
        pass

    # Step 2: spawn the Fortran solver with shell redirection --------------
    cmd = '"{}" < "input.inp" > "output.out"'.format(exe_path)
    try:
        process = subprocess.Popen(cmd, shell=True, cwd=project)
    except Exception as e:
        messagebox.showerror("Run Model",
                             f"Could not start {os.path.basename(exe_path)}:\n{e}")
        return False

    sheet_name = getattr(app, "active_sheet", "Simple")
    workbook_path = _xlsm_path()

    # Step 3: runtime-clock popup ------------------------------------------
    root = tk.Toplevel(parent or app)
    root.title("REMFluor Runtime Clock")
    root.configure(bg="#F0F0F0")
    try:
        root.transient(parent or app)
    except Exception:
        pass

    label = tk.Label(root, text="Starting...", font=("Arial", 22, "bold"),
                     bg="#F0F0F0", fg="#222", padx=24, pady=18,
                     justify="center", wraplength=1080)
    label.pack(padx=32, pady=(40, 14), fill="x")

    detail = tk.Label(root,
                      text=f"Running {os.path.basename(exe_path)} ...\n"
                           "Dashboard will open in your browser when finished.",
                      font=("Arial", 12), bg="#F0F0F0", fg="#555",
                      justify="center", wraplength=1080)
    detail.pack(padx=32, pady=(0, 28))

    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 64, 1180)
    h = max(root.winfo_reqheight() + 48, 360)
    try:
        sw = root.winfo_screenwidth(); sh = root.winfo_screenheight()
        x = max(0, (sw - w) // 2); y = max(0, (sh - h) // 2 - 40)
        root.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        root.geometry(f"{w}x{h}")
    root.minsize(w, h)
    root.resizable(True, True)
    try:
        root.lift(); root.focus_force()
        root.attributes("-topmost", True)
        root.after(800, lambda: root.attributes("-topmost", False))
    except Exception:
        pass

    start = time.time()
    state = {"dashboard_started": False}

    def _tick():
        elapsed = time.time() - start
        mins, secs = divmod(int(elapsed), 60)
        if process.poll() is None:
            label.config(text=f"Runtime: {mins:02d}:{secs:02d}")
            root.after(1000, _tick)
        else:
            label.config(text=f"Finished: {mins:02d}:{secs:02d}")
            root.update_idletasks()
            if not state["dashboard_started"]:
                state["dashboard_started"] = True
                _launch_dashboard_async(workbook_path, sheet_name,
                                        label, root, mins, secs)

    def _on_close():
        # Don't kill the model if user closes the clock — let it finish
        try: root.withdraw()
        except Exception: pass

    root.protocol("WM_DELETE_WINDOW", _on_close)
    _tick()
    return True
