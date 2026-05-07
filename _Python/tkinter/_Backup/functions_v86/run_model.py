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
_DASHBOARD_PORT = 8050


def _project_dir() -> str:
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(here, "..", "..", ".."))


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

    # Belt + suspenders: kill anything else listening on port 8050.
    # This catches dashboards left over from a prior app session that
    # crashed (no _LAST_DASHBOARD_PROC handle), or zombies bash didn't
    # reap.  Cross-platform via psutil if available, else OS fallback.
    try:
        import psutil   # type: ignore
        for c in psutil.net_connections(kind="inet"):
            if c.status != psutil.CONN_LISTEN:
                continue
            laddr = getattr(c, "laddr", None)
            if laddr and laddr.port == _DASHBOARD_PORT and c.pid:
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
            pattern = re.compile(rf":{_DASHBOARD_PORT}\s+.*LISTENING\s+(\d+)")
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
            subprocess.run(
                ["fuser", "-k", f"{_DASHBOARD_PORT}/tcp"],
                capture_output=True, text=True, timeout=5,
            )
    except Exception:
        pass


def _xlsm_path() -> str:
    return os.path.join(_project_dir(),
                        "REMFluor-MD Interface Storyboard v2.6.xlsm")


def _dashboard_script() -> str:
    # Lives next to this file (tkinter/functions/) — kept in functions/
    # for consistency with every other ported module.  Source_Py copy
    # remains in the repo as historical reference only.
    return os.path.join(os.path.dirname(os.path.abspath(__file__)),
                        "generate_dashboard.py")


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
    state = getattr(app, "state", None)

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

    # Pull a few values up-front so we can use them in defaults below.
    start_year = _int("E18", 2025)
    end_year   = _int("E19", start_year + 100)

    payload = {
        # ── Scalars used for plot context ──────────────────────────
        "A8":   _int("A8", 1),       # version flag (1=Simple, 2=Detailed)
        "AD1":  _int("AD1", 1),      # unit flag (1=feet, 2=meters)
        "K38":  _str("K38", ""),     # Precursor 1 name (blank = ipre=0)
        "E38":  _str("E38", "PFAA-1"),
        "M38":  _str("M38", ""),     # Precursor 2 name
        "G38":  _str("G38", ""),     # PFAA 2 name (blank = ncomp=1)

        # See Results Every (yr) — int avoids modulo issues
        "V47":  _int("V47", 5),

        # Time
        "E18":  start_year,
        "E19":  end_year,

        # Geometry / source — non-zero defaults so mass calc isn't 0
        "E13":  _num("E13", 10.0),   # Total depth
        "E15":  _num("E15", 10.0),   # Source width
        "E16":  _num("E16", 5.0),    # Source thickness
        "E22":  _num("E22", 1.0),    # Velocity (Darcy)

        # Sample years (only one is used depending on version_flag)
        "R36":  _int("R36", start_year),
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
    Fast — typically <0.5 s."""
    # v86: dashboard reads cell values from dashboard_state.json
    # (written by _dump_dashboard_state with hard-coded defaults), so
    # pywin32 is no longer required.  openpyxl is still required only
    # for the user's optional calibration .xlsx (Model Data sheet) —
    # but we don't gate on it here; the dashboard imports it lazily.
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
        proj   = _project_dir()
        log_path = os.path.join(proj, "dashboard.log")

        if not os.path.exists(script):
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

        rc = proc.