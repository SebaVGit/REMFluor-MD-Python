"""
popups_numerical.py — pure-Python port.

Adapted from Source_Py/popups_numerical.py.  Persists Timestep Size,
Convergence Tolerance and the iTVD flag (TVD vs Upstream weighting) to
numerical_inputs.txt — same format as the .exe pipeline.

Called from main.run_script() when ChangeNumericalParameters is dispatched.
"""
from __future__ import annotations
import os
import platform
import subprocess
import webbrowser
import tkinter as tk
from tkinter import messagebox

from .state import get_state


PARAMETERS = ["Timestep Size (yr) ", "Convergence Tolerance (ug/L)"]
HELP_SECTION = {
    "Timestep Size (yr) ":           "timestep-size",
    "Convergence Tolerance (ug/L)":  "convergence-tolerance",
}

FONT_TITLE  = ("Arial", 16, "bold")
FONT_LABEL  = ("Arial", 11)
FONT_HEADER = ("Arial", 11, "bold")
FONT_BTN    = ("Arial", 11)
FONT_HELP   = ("Arial", 10, "bold")
FONT_VAL    = ("Arial", 11)


def _docs_root():
    """v100: in a frozen --onefile build, walking up from __file__
    lands inside sys._MEIPASS\functions which is wrong.  Use the
    state singleton's bundle_dir / work_dir set by main.py at
    startup; fall back to the dev-tree walk if neither is set."""
    try:
        from .state import get_state
        st = get_state()
        for base in (getattr(st, "bundle_dir", ""),
                     getattr(st, "work_dir", "")):
            if base:
                cand = os.path.join(base, "docs", "_site")
                if os.path.isdir(cand):
                    return cand
    except Exception:
        pass
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.abspath(os.path.join(here, "..", "..", "..",
                                        "docs", "_site"))


def _open_help_section(section_id):
    f = os.path.join(_docs_root(),
                     "data_chicklets", "Step11_ModelingParameters.html")
    if not os.path.exists(f):
        messagebox.showerror("Help Not Found", f"Help file not found:\n{f}")
        return
    abs_p = os.path.abspath(f).replace("\\", "/")
    anchor = f"#{section_id}" if section_id else ""
    url = (f"file:///{abs_p}{anchor}" if os.name == "nt" and abs_p[1] == ":"
           else f"file://{abs_p}{anchor}")
    try:
        if platform.system() == "Windows":
            for exe in (r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"):
                if os.path.exists(exe):
                    subprocess.Popen([exe, url]); return
        webbrowser.open(url)
    except Exception:
        webbrowser.open(url)


def _load_defaults(path: str):
    out = {}
    if not os.path.exists(path):
        return out, 1
    iTVD = 1
    try:
        with open(path) as f:
            for ln in f:
                ln = ln.strip()
                if "," not in ln:
                    continue
                parts = ln.split(",")
                if len(parts) != 2 or parts[0].strip() == "Parameter":
                    continue
                key = parts[0].strip()
                try:
                    val = float(parts[1].strip())
                except ValueError:
                    continue
                if key.lower() == "itvd":
                    iTVD = int(val)
                else:
                    out[parts[0]] = val   # keep original key (with trailing space)
    except Exception:
        pass
    return out, iTVD


def run(app, parent=None):
    state = get_state()
    work_dir = state.work_dir or os.getcwd()
    txt_path = os.path.join(work_dir, "numerical_inputs.txt")
    defaults, default_iTVD = _load_defaults(txt_path)

    root = tk.Toplevel(parent or app)
    root.title("Change Numerical Parameters")
    root.configure(bg="#F0F0F0")
    try: root.withdraw()
    except Exception: pass
    # Skip transient() so the popup keeps a maximize button on Windows.
    try:
        root.grab_set()
    except Exception:
        pass
    outer = tk.Frame(root, bg="#F0F0F0", padx=24, pady=18)
    outer.pack(fill="both", expand=True)

    tk.Label(outer, text="Change Numerical Parameters", font=FONT_TITLE,
             bg="#F0F0F0").grid(row=0, column=0, columnspan=4,
                                pady=(0, 18), sticky="w")

    weighting_var = tk.StringVar(value="TVD" if default_iTVD == 1 else "Upstream")
    tk.Label(outer, text="Numerical Method:", font=FONT_HEADER, bg="#F0F0F0"
             ).grid(row=1, column=0, sticky="w", pady=(0, 8))
    tk.Radiobutton(outer, text="TVD", variable=weighting_var, value="TVD",
                   font=FONT_LABEL, bg="#F0F0F0"
                   ).grid(row=1, column=1, sticky="w")
    tk.Radiobutton(outer, text="Upstream weighting method",
                   variable=weighting_var, value="Upstream",
                   font=FONT_LABEL, bg="#F0F0F0"
                   ).grid(row=1, column=2, columnspan=2, sticky="w")

    entries = {}
    r = 2
    for param in PARAMETERS:
        tk.Label(outer, text=param.strip(), font=FONT_LABEL, bg="#F0F0F0",
                 anchor="w").grid(row=r, column=0, sticky="w", pady=4)
        sec_id = HELP_SECTION.get(param, "")
        if sec_id:
            tk.Button(outer, text="?", width=2, height=1, fg="red",
                      font=FONT_HELP,
                      command=lambda sid=sec_id: _open_help_section(sid)
                      ).grid(row=r, column=1, padx=(8, 12))
        e = tk.Entry(outer, width=12, font=FONT_VAL, justify="left")
        e.grid(row=r, column=2, sticky="w", pady=4)
        if param in defaults:
            e.insert(0, str(defaults[param]))
        entries[param] = e
        r += 1

    tk.Label(outer, text="", bg="#F0F0F0").grid(row=r, column=0, pady=8)
    r += 1

    def _save_and_exit():
        try:
            data = {}
            for param in PARAMETERS:
                s = entries[param].get().strip()
                if not s:
                    messagebox.showerror("Error", f"Missing value: {param}",
                                         parent=root); return
                try:
                    data[param] = float(s)
                except ValueError:
                    messagebox.showerror("Error", f"Invalid value: {param}",
                                         parent=root); return
            iTVD_val = 1 if weighting_var.get() == "TVD" else 0
            if os.path.exists(txt_path):
                try:
                    os.chmod(txt_path, 0o666); os.remove(txt_path)
                except Exception: pass
            with open(txt_path, "w") as f:
                f.write("iTVD\n")
                f.write(f"iTVD, {iTVD_val}\n\n")
                f.write(" Zone\n")
                f.write("Parameter,\n")
                for param in PARAMETERS:
                    f.write(f"{param},{data[param]}\n")
                f.write("\n")
            messagebox.showinfo("Success",
                                f"Saved: {os.path.basename(txt_path)}",
                                parent=root)
            try: root.grab_release()
            except Exception: pass
            root.destroy()
        except Exception as e:
            messagebox.showerror("Unexpected Error", str(e), parent=root)

    def _cancel():
        try: root.grab_release()
        except Exception: pass
        root.destroy()

    bar = tk.Frame(outer, bg="#F0F0F0")
    bar.grid(row=r, column=0, columnspan=4, pady=(8, 0))
    tk.Button(bar, text="OK",     width=10, font=FONT_BTN,
              command=_save_and_exit).pack(side="left", padx=6)
    tk.Button(bar, text="Cancel", width=10, font=FONT_BTN,
              command=_cancel).pack(side="left", padx=6)

    root.protocol("WM_DELETE_WINDOW", _cancel)
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 32, 600)
    h = max(root.winfo_reqheight() + 24, 320)
    try:
        sw = root.winfo_screenwidth(); sh = root.winfo_screenheight()
        x = max(0, (sw - w) // 2); y = max(0, (sh - h) // 2 - 30)
        root.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        root.geometry(f"{w}x{h}")
    root.minsize(w, h); root.resizable(True, True)
    try:
        root.deiconify(); root.lift(); root.focus_force()
    except Exception: pass
    root.wait_window()
