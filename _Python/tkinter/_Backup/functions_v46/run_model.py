"""
run_model.py — Run Model pipeline (replaces the broken inline launcher).

Mirrors the Source_Py/input_variables.py __main__ block:

    1.  Build/refresh input.inp via generate_input_file.run(app)
    2.  Run remfluor_v8a.exe in a shell with stdin/stdout redirection:
            remfluor_v8a.exe < input.inp > output.out
    3.  Show a tkinter Toplevel with a live "Runtime Clock" while the
        Fortran solver runs.
    4.  When the model finishes, launch the Plotly-Dash dashboard
        (Source_Py/generate_dashboard.py) — the dashboard auto-opens
        the user's default browser (Edge/Chrome) at http://localhost:8050.

Called from main.run_script() when RunPythonScript is dispatched.
"""
from __future__ import annotations
import os
import sys
import time
import threading
import subprocess
import tkinter as tk
from tkinter import messagebox

from . import generate_input_file


def _project_dir() -> str:
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(here, "..", "..", ".."))


def _xlsm_path() -> str:
    return os.path.join(_project_dir(),
                        "REMFluor-MD Interface Storyboard v2.6.xlsm")


def _dashboard_script() -> str:
    return os.path.join(_project_dir(), "Source_Py", "generate_dashboard.py")


def _check_dashboard_imports(py_exec: str):
    """Run a tiny child-Python import test for everything
    generate_dashboard.py needs.  Returns (ok, missing_modules_list).
    Fast — typically <0.5 s."""
    test_code = (
        "import importlib, sys\n"
        "missing = []\n"
        "for m in ('pandas','numpy','plotly','dash','openpyxl',"
        "'win32com.client','pythoncom'):\n"
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
                    text=f"Finished: {mins:02d}:{secs:02d}\n"
                         "Dashboard Launched\n"
                         "(check your browser — http://localhost:8050)"))
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

    # Step 2: spawn the Fortran solver with shell redirection --------------
    # Equivalent to: remfluor_v8a.exe < input.inp > output.out
    cmd = '"{}" < "input.inp" > "output.out"'.format(exe_path)
    try:
        process = subprocess.Popen(cmd, shell=True, cwd=project)
    except Exception as e:
        messagebox.showerror("Run Model",
                             f"Could not start remfluor_v8a.exe:\n{e}")
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

    label = tk.Label(root, text="Starting...", font=("Arial", 36, "bold"),
                     bg="#F0F0F0", fg="#222", padx=40, pady=32,
                     justify="center", wraplength=900)
    label.pack(padx=48, pady=(56, 18), fill="x")

    detail = tk.Label(root,
                      text="Running remfluor_v8a.exe ...\n"
                           "Dashboard will open in your browser when finished.",
                      font=("Arial", 14), bg="#F0F0F0", fg="#555",
                      justify="center", wraplength=900)
    detail.pack(padx=48, pady=(0, 40))

    # Center & size after layout — much bigger so user can read clock at a glance
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 64, 1000)
    h = max(root.winfo_reqheight() + 48, 520)
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
            # Don't auto-hide: keep the runtime clock visible so the
            # user can read the dashboard launch status (or any error
            # message).  They can close it manually.

    def _safe_withdraw(win):
        try: win.withdraw()
        except Exception: pass

    def _on_close():
        # Don't kill the model if user closes the clock — let it finish
        try: root.withdraw()
        except Exception: pass

    root.protocol("WM_DELETE_WINDOW", _on_close)
    _tick()
    return True
