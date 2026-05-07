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
    """Write a JSON snapshot of every state cell the dashboard needs.

    v82: replaces the dashboard's previous openpyxl/win32com dependency
    on the .xlsm Storyboard workbook.  Called right before the dashboard
    subprocess is spawned, so the dashboard can run with zero Excel
    dependencies — the .exe build no longer needs xlsm bundled.
    """
    state = getattr(app, "state", None)

    def _g(addr):
        try:
            return state.get(addr) if state is not None else None
        except Exception:
            return None

    def _list(start_row, end_row, col):
        return [_g(f"{col}{r}") for r in range(start_row, end_row + 1)]

    payload = {
        # Scalars consumed by the dashboard's main()
        "A8":   _g("A8"),    # version flag (1 = simple, 2 = detailed)
        "AD1":  _g("AD1"),   # unit flag (1 = ft, 2 = m)
        "K38":  _g("K38"),   # Precursor 1 name
        "E38":  _g("E38"),   # PFAA 1 name
        "M38":  _g("M38"),   # Precursor 2 name
        "G38":  _g("G38"),   # PFAA 2 name
        "V47":  _g("V47"),   # See Results Every (yr)
        "E18":  _g("E18"),   # Start Year
        "E19":  _g("E19"),   # End Year
        "E13":  _g("E13"),   # Total Depth
        "E15":  _g("E15"),   # Source Width
        "E16":  _g("E16"),   # Source Thickness
        "E22":  _g("E22"),   # Velocity (Darcy)
        "R36":  _g("R36"),   # Sample year (Simple version)
        "Y74":  _g("Y74"),   # Sample year (Detailed §10)

        # Ranges — pre-flattened for the dashboard
        "U8:U18":  _list(8, 18, "U"),    # Source loading time
        "V8:V18":  _list(8, 18, "V"),    # Source conc PFAA1
        "X8:X18":  _list(8, 18, "X"),    # Source conc PFAA2
        "Z8:Z18":  _list(8, 18, "Z"),    # Source conc Precursor1
        "AB8:AB18":_list(8, 18, "AB"),   # Source conc Precursor2
        "U34:U40": _list(34, 40, "U"),   # MW names
        "V34:V40": _list(34, 40, "V"),   # MW PFAA1 conc
        "X34:X40": _list(34, 40, "X"),   # MW PFAA2 conc
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
    # v82: pywin32 + openpyxl are no longer required up-front.  The
    # dashboard reads workbook state from dashboard_state.json and only
    # uses openpyxl as a deferred optional import for the user's
    # calibration .xlsx (Model Data sheet).
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
        # If something's missing we abort immediately and tell the user
        # which `pip install` they're missing — far more helpful than a
        # silent crash a few seconds later.
        ok, missing = _check_dashboard_imports(py)
        if not ok:
            # Translate module names → pip package names users care about
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

        # Kill any prior dashboard process (or any other listener on
        # port 8050) before spawning the new one.  Without this, a
        # stale dashboard from a previous Run Model click hogs the
        # port and the user keeps seeing OLD output even after the
        # solver wrote fresh .out files — which matches the user's
        # "I changed See Results Every and it didn't work" report.
        _kill_previous_dashboard()
        # Tiny grace period so the OS releases the socket before the
        # new dash binds.  500 ms is plenty on Win10/11.
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
            # CREATE_NEW_PROCESS_GROUP only — keep stdout redirection so we
            # can read errors.  DETACHED_PROCESS would close the file
            # handles we just opened.
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
                    text=f"Finished: {mins:02d}:{secs:02d}\n"
                         f"Dashboard error:\n{e}"))
                root.after(0, lambda: messagebox.showerror(
                    "Dashboard Error",
                    f"Failed to launch dashboard:\n{e}\n\n"
                    f"Log: {log_path}", parent=root))
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
            # Still running — assume the Dash server is up; browser
            # auto-opens via webbrowser.open() in the script.
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

        msg_short = (f"Finished: {mins:02d}:{secs:02d}\n"
                     f"Dashboard FAILED (exit {rc}).\n"
                     f"See: {log_path}")
        msg_full = (
            f"Dashboard process exited with code {rc}.\n\n"
            f"Log file:\n{log_path}\n\n"
            f"--- last lines of log ---\n{tail or '(empty log)'}"
        )
        try:
            root.after(0, lambda: label.config(text=msg_short))
            root.after(0, lambda: messagebox.showerror(
                "Dashboard Error", msg_full, parent=root))
        except Exception:
            pass

    threading.Thread(target=_worker, daemon=False).start()


def run(app, parent=None) -> bool:
    """Run the model pipeline.  Returns True if launch succeeded."""
    # Step 1: build input.inp from current state ---------------------------
    if not generate_input_file.run(app):
        return False

    project = _project_dir()
    inp_path = os.path.join(project, "input.inp")
    exe_path = os.path.join(project, "remfluor_v8a.exe")

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

    # Step 1b (v82): dump dashboard state to JSON so the dashboard
    # subprocess has zero dependency on the .xlsm Storyboard workbook.
    # Replaces the openpyxl/win32com reads that used to live in
    # generate_dashboard.py.  Written into the project dir alongside
    # input.inp so the dashboard finds it in cwd.
    try:
        _dump_dashboard_state(app, project)
    except Exception:
        pass

    # Step 2: spawn the Fortran solver with shell redirection --------------
    # Equivalent to: remfluor_v8a.exe < input.inp > output.out
    cmd = '"{}" < "input.inp" > "output.out"'.format(exe_path)
    try:
        process = subprocess.Popen(cmd, shell=True, cwd=project)
    except Exception as e:
        messagebox.showerror("Run Model",
 