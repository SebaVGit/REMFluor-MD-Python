"""
popups_GWvelocity.py — pure-Python port.

Adapted from Source_Py/popups_GWvelocity.py.  Computes the groundwater
bulk Darcy velocity v = K * i, exposes 9 unit options for K, persists
to gwvelocity_inputs.txt, and writes the result back into the live app's
Section-3 Vd field (v_darcy) in m/year (or ft/year if units = feet).

Called from main.run_script() when GWVelocityCalculator is dispatched.
"""
from __future__ import annotations
import os
import platform
import subprocess
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox

from .state import get_state


UNIT_TO_MDAY = {
    "m/s":     86400.0,
    "m/day":   1.0,
    "m/year":  1.0 / 365.25,
    "cm/s":    864.0,
    "cm/day":  0.01,
    "mm/day":  0.001,
    "ft/day":  0.3048,
    "ft/s":    26334.72,
    "in/day":  0.0254,
}

FONT_TITLE = ("Arial", 16, "bold")
FONT_LABEL = ("Arial", 11)
FONT_BOLD  = ("Arial", 11, "bold")
FONT_SMALL = ("Arial", 9)
FONT_BTN   = ("Arial", 11)
FONT_VAL   = ("Arial", 11)


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
    f = os.path.join(_docs_root(), "appendix", "appendix_3_1.html")
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


def _load_existing(path: str):
    out = {}
    if not os.path.exists(path):
        return out
    try:
        with open(path) as f:
            for ln in f:
                ln = ln.strip()
                if "Bulk Hydraulic Conductivity Value" in ln and ":" in ln:
                    try: out["k_value"] = float(ln.split(":", 1)[1].strip())
                    except ValueError: pass
                elif "Bulk Hydraulic Conductivity Unit" in ln and ":" in ln:
                    out["k_unit"] = ln.split(":", 1)[1].strip()
                elif ln.startswith("Bulk Hydraulic Gradient") and ":" in ln:
                    try: out["gradient"] = float(ln.split(":", 1)[1].strip())
                    except ValueError: pass
    except Exception:
        pass
    return out


def run(app, parent=None):
    state = get_state()
    work_dir = state.work_dir or os.getcwd()
    txt_path = os.path.join(work_dir, "gwvelocity_inputs.txt")
    existing = _load_existing(txt_path)

    units_str = "feet" if (getattr(app, "v_units", None)
                           and app.v_units.get() == "feet") else "meters"

    root = tk.Toplevel(parent or app)
    root.title("Groundwater Bulk Darcy Velocity Calculator")
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

    tk.Label(outer, text="Groundwater Bulk Darcy Velocity Calculator",
             font=FONT_TITLE, bg="#F0F0F0"
             ).pack(anchor="w", pady=(0, 12))

    formula = tk.Frame(outer, bg="#F0F0F0")
    formula.pack(anchor="w", pady=(0, 12), fill="x")
    tk.Label(formula, text="Formula:  v = K * i",
             font=("Arial", 12, "italic"), bg="#F0F0F0"
             ).pack(anchor="w")
    tk.Label(formula, text="where:  v = Bulk Darcy velocity,  "
                            "K = Bulk hydraulic conductivity,  i = gradient",
             font=FONT_SMALL, bg="#F0F0F0", fg="#333"
             ).pack(anchor="w")
    tk.Label(formula,
             text='Note: "Bulk K" includes the effects of any low-k layers '
                  'and lenses in the aquifer cross-section\n(but not '
                  'aquitards on the top or bottom).',
             font=FONT_SMALL, bg="#F0F0F0", fg="gray", justify="left"
             ).pack(anchor="w", pady=(4, 0))

    # K row
    k_row = tk.Frame(outer, bg="#F0F0F0"); k_row.pack(anchor="w",
                                                      pady=(8, 4), fill="x")
    tk.Label(k_row, text="Bulk Hydraulic Conductivity (K):",
             font=FONT_BOLD, bg="#F0F0F0", width=30, anchor="w"
             ).pack(side="left")
    sv_k = tk.StringVar(value=str(existing.get("k_value", "")))
    tk.Entry(k_row, textvariable=sv_k, font=FONT_VAL, width=15
             ).pack(side="left", padx=4)
    sv_unit = tk.StringVar(value=existing.get("k_unit", "m/day"))
    ttk.Combobox(k_row, textvariable=sv_unit, state="readonly",
                 values=list(UNIT_TO_MDAY.keys()), width=10,
                 font=FONT_LABEL).pack(side="left", padx=4)

    # Gradient row
    g_row = tk.Frame(outer, bg="#F0F0F0"); g_row.pack(anchor="w",
                                                      pady=4, fill="x")
    tk.Label(g_row, text="Bulk Hydraulic Gradient (i):",
             font=FONT_BOLD, bg="#F0F0F0", width=30, anchor="w"
             ).pack(side="left")
    sv_grad = tk.StringVar(value=str(existing.get("gradient", "")))
    tk.Entry(g_row, textvariable=sv_grad, font=FONT_VAL, width=15
             ).pack(side="left", padx=4)
    tk.Label(g_row, text="(dimensionless)", font=FONT_SMALL,
             fg="gray", bg="#F0F0F0").pack(side="left", padx=8)

    # Result rows
    tk.Label(outer, text="Calculated Bulk Darcy Velocity:",
             font=("Arial", 12, "bold"), bg="#F0F0F0"
             ).pack(anchor="w", pady=(16, 4))
    sv_my = tk.StringVar()
    sv_fy = tk.StringVar()
    for label_text, sv in (("Darcy Velocity (m/year):", sv_my),
                           ("Darcy Velocity (ft/year):", sv_fy)):
        row = tk.Frame(outer, bg="#F0F0F0")
        row.pack(anchor="w", padx=20, pady=2, fill="x")
        tk.Label(row, text=label_text, font=FONT_LABEL, bg="#F0F0F0",
                 width=24, anchor="w").pack(side="left")
        tk.Entry(row, textvariable=sv, font=FONT_VAL, width=15,
                 state="readonly").pack(side="left", padx=10)

    def _calc():
        try:
            k = float(sv_k.get())
            i = float(sv_grad.get())
        except ValueError:
            messagebox.showerror("Error",
                                 "Please enter valid numeric K and gradient.",
                                 parent=root); return None, None
        if k <= 0 or i <= 0:
            messagebox.showerror("Error",
                                 "K and gradient must both be positive.",
                                 parent=root); return None, None
        unit = sv_unit.get()
        if unit not in UNIT_TO_MDAY:
            messagebox.showerror("Error", f"Unknown unit: {unit}",
                                 parent=root); return None, None
        v_mday = k * UNIT_TO_MDAY[unit] * i
        v_my = v_mday * 365.25
        v_fy = v_my / 0.3048
        sv_my.set(f"{v_my:.6f}")
        sv_fy.set(f"{v_fy:.6f}")
        return v_my, v_fy

    def _apply():
        v_my, v_fy = _calc()
        if v_my is None:
            return
        # Persist
        try:
            if os.path.exists(txt_path):
                try:
                    os.chmod(txt_path, 0o666); os.remove(txt_path)
                except Exception: pass
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write("Groundwater BulkDarcy Velocity Calculator Results\n")
                f.write(f"Bulk Hydraulic Conductivity Value: {sv_k.get()}\n")
                f.write(f"Bulk Hydraulic Conductivity Unit: {sv_unit.get()}\n")
                f.write(f"Bulk Hydraulic Gradient: {sv_grad.get()}\n")
                f.write(f"Bulk Darcy Velocity (m/year): {v_my:.6f}\n")
                f.write(f"Bulk Darcy Velocity (ft/year): {v_fy:.6f}\n")
        except Exception as e:
            messagebox.showerror("Error",
                                 f"Could not write {os.path.basename(txt_path)}:\n{e}",
                                 parent=root); return
        # Push to app: Section 3 Vd (m/year if meters, ft/year if feet)
        try:
            target = v_fy if units_str == "feet" else v_my
            app.v_darcy.set(f"{target:.2f}")
        except Exception:
            pass
        messagebox.showinfo(
            "Success",
            f"Bulk Darcy Velocity ({units_str}/year) applied to Section 3.\n\n"
            f"K  : {sv_k.get()} {sv_unit.get()}\n"
            f"i  : {sv_grad.get()}\n"
            f"v  : {v_my:.6f} m/year   ({v_fy:.6f} ft/year)",
            parent=root)
        try: root.grab_release()
        except Exception: pass
        root.destroy()

    def _cancel():
        try: root.grab_release()
        except Exception: pass
        root.destroy()

    bar = tk.Frame(outer, bg="#F0F0F0"); bar.pack(pady=(20, 0))
    tk.Button(bar, text="Calculate", width=10, font=FONT_BTN,
              command=_calc).pack(side="left", padx=6)
    tk.Button(bar, text="Apply",     width=10, font=FONT_BTN,
              command=_apply).pack(side="left", padx=6)
    tk.Button(bar, text="Cancel",    width=10, font=FONT_BTN,
              command=_cancel).pack(side="left", padx=6)
    tk.Button(bar, text="Help",      width=10, font=FONT_BTN,
              command=_open_help).pack(side="left", padx=6)

    root.protocol("WM_DELETE_WINDOW", _cancel)
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 32, 700)
    h = max(root.winfo_reqheight() + 24, 480)
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
