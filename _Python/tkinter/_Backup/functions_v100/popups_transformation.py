"""
popups_transformation.py — pure-Python port.

Adapted from Source_Py/popups_transformation.py.  Persists Yield Factor
and Decay half-life for Precursors 1 / 2 in the Low-K zone to
transformation_inputs.txt (same format as the .exe pipeline).

Called from main.run_script() when ModelingTransformationLowK is dispatched.
"""
from __future__ import annotations
import os
import platform
import subprocess
import webbrowser
import tkinter as tk
from tkinter import messagebox

from .state import get_state


PRECURSORS = ["Precursors 1", "Precursors 2"]
PARAMETERS = ["Yield Factor (-)", "Decay half-life (years)"]
ZONES      = ["Low-K"]

FONT_TITLE  = ("Arial", 16, "bold")
FONT_LABEL  = ("Arial", 11)
FONT_HEADER = ("Arial", 11, "bold")
FONT_ZONE   = ("Arial", 13, "underline")
FONT_BTN    = ("Arial", 11)
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


def _open_help():
    f = os.path.join(_docs_root(), "appendix", "appendix_5_3.html")
    if not os.path.exists(f):
        messagebox.showerror("Help Not Found", f"Help file not found:\n{f}")
        return
    abs_p = os.path.abspath(f).replace("\\", "/")
    url = (f"file:///{abs_p}" if os.name == "nt" and abs_p[1] == ":"
           else f"file://{abs_p}")
    try:
        if platform.system() == "Windows":
            for exe in (r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"):
                if os.path.exists(exe):
                    subprocess.Popen([exe, url]); return
        webbrowser.open(url)
    except Exception:
        webbrowser.open(url)


def _read_existing(path: str):
    if not os.path.exists(path):
        return {}
    try:
        with open(path) as f:
            lines = [ln.strip() for ln in f if ln.strip()]
    except Exception:
        return {}
    out = {}
    for ln in lines:
        for param in PARAMETERS:
            if ln.startswith(param):
                parts = ln.split(",")
                vals = []
                for p in parts[1:]:
                    try: vals.append(float(p.strip()))
                    except ValueError: vals.append(0.0)
                out[param] = vals
    return out


def run(app, parent=None):
    state = get_state()
    work_dir = state.work_dir or os.getcwd()
    txt_path = os.path.join(work_dir, "transformation_inputs.txt")
    existing = _read_existing(txt_path)

    root = tk.Toplevel(parent or app)
    root.title("Enter Parameters for Modeling Transformation Low-K")
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

    tk.Label(outer,
             text="Enter Parameters for Modeling Transformation Low-K",
             font=FONT_TITLE, bg="#F0F0F0"
             ).grid(row=0, column=0, columnspan=4, pady=(0, 18), sticky="w")

    r = 1
    tk.Label(outer, text="Low-K Zone", font=FONT_ZONE, bg="#F0F0F0"
             ).grid(row=r, column=0, sticky="w", pady=(8, 4))
    r += 1
    for col, p in enumerate(PRECURSORS, start=1):
        tk.Label(outer, text=p, font=FONT_HEADER, bg="#F0F0F0"
                 ).grid(row=r, column=col, padx=8)
    r += 1

    entries = {}
    for param in PARAMETERS:
        tk.Label(outer, text=param, font=FONT_LABEL, bg="#F0F0F0", anchor="w"
                 ).grid(row=r, column=0, sticky="w", pady=4)
        for col, p in enumerate(PRECURSORS, start=1):
            e = tk.Entry(outer, width=12, font=FONT_VAL, justify="left")
            e.grid(row=r, column=col, padx=8, pady=4)
            vals = existing.get(param, [])
            if col - 1 < len(vals):
                e.insert(0, str(vals[col - 1]))
            else:
                e.insert(0, "0")
            entries[(param, p)] = e
        r += 1

    tk.Label(outer, text="", bg="#F0F0F0").grid(row=r, column=0, pady=8)
    r += 1

    def _save_and_exit():
        try:
            data = {}
            for param in PARAMETERS:
                data[param] = []
                for p in PRECURSORS:
                    s = entries[(param, p)].get().strip()
                    try:
                        data[param].append(float(s) if s else 0.0)
                    except ValueError:
                        data[param].append(0.0)
            if os.path.exists(txt_path):
                try:
                    os.chmod(txt_path, 0o666); os.remove(txt_path)
                except Exception: pass
            with open(txt_path, "w") as f:
                f.write("Low-K Zone\n")
                f.write("Parameter," + ",".join(PRECURSORS) + "\n")
                for param in PARAMETERS:
                    f.write(param + "," +
                            ",".join(str(v) for v in data[param]) + "\n")
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
    tk.Button(bar, text="Help",   width=10, font=FONT_BTN,
              command=_open_help).pack(side="left", padx=6)

    root.protocol("WM_DELETE_WINDOW", _cancel)
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 32, 560)
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
