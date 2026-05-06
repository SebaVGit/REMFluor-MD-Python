"""
popups_longevity.py — pure-Python port.

Adapted from Source_Py/popups_longevity.py.  Estimates breakthrough time
for a CAC barrier as longevity = coefficient * fcac * W * (rho_b * 1000) *
C^exponent / (Vd / 365).  Reads inputs from the live app StringVars
(no xlsm required) and persists results to longevity_inputs.txt.

Called from main.run_script() when LongevityTool is dispatched.
"""
from __future__ import annotations
import os
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox

from .state import get_state


CAC_PAPER_URL = "https://onlinelibrary.wiley.com/doi/pdf/10.1002/rem.70017"
PFAS_OPTIONS  = ["PFOA", "PFOS", "PFHxS"]
COEFFS = {
    "PFOA":  (0.245, -0.775),
    "PFOS":  (2.60,  -0.675),
    "PFHxS": (0.608, -0.804),
}

# Match the Excel "Longevity Calculator" popup styling exactly:
# - Centered, large bold title
# - Bold section labels, regular body text
# - Indented param rows
FONT_TITLE = ("Arial", 22, "bold")
FONT_HDR   = ("Arial", 14, "bold")
FONT_LABEL = ("Arial", 12)
FONT_BOLD  = ("Arial", 12, "bold")
FONT_SMALL = ("Arial", 10)
FONT_BTN   = ("Arial", 12)
FONT_VAL   = ("Arial", 12)
FONT_RES   = ("Arial", 14, "bold")


def _safe_float(s, default=None):
    try: return float(str(s).strip())
    except (ValueError, TypeError): return default


def _read_bulk_density(work_dir):
    """Try retardation_inputs.txt → Transmissive Zone Soil Bulk Density."""
    f = os.path.join(work_dir, "retardation_inputs.txt")
    if not os.path.exists(f):
        return None
    try:
        with open(f) as fp:
            lines = [ln.strip() for ln in fp if ln.strip()]
        in_trans = False
        for ln in lines:
            if ln == "Transmissive Zone":
                in_trans = True; continue
            if ln.endswith("Zone"):
                in_trans = False
            if in_trans and ln.startswith("Soil Bulk Density"):
                try: return float(ln.split(",")[1])
                except (IndexError, ValueError): pass
    except Exception:
        pass
    return None


def _get_var(app, name, default=""):
    """Safely fetch a StringVar's value off the app, falling back to
    `default` if the attribute is missing or .get() raises."""
    v = getattr(app, name, None)
    if v is None:
        return default
    try:
        return v.get()
    except Exception:
        return default


def run(app, parent=None):
    state = get_state()
    work_dir = state.work_dir or os.getcwd()
    txt_path = os.path.join(work_dir, "longevity_inputs.txt")

    # Inputs from app -------------------------------------------------------
    # fcac comes from Section 9 PSB Loading (v_psb_load) — entered as %
    # in storyboard; we need it as a fraction (0..1).
    fcac_raw = _safe_float(_get_var(app, "v_psb_load"), None)
    if fcac_raw is None or fcac_raw == 0:
        fcac = 0.01            # 1 % default when blank / zero
    elif fcac_raw > 1:
        fcac = fcac_raw / 100  # value entered as percent
    else:
        fcac = fcac_raw        # already a fraction
    # Width comes from Section 9 v_psb_width
    width = _safe_float(_get_var(app, "v_psb_width"), 4.0)
    # Vd from Section 3 (m/yr)
    gw_velocity = _safe_float(_get_var(app, "v_darcy"), 50.0)
    # Bulk density from retardation_inputs.txt → fallback 1.6
    bulk_density = _read_bulk_density(work_dir)
    if bulk_density is None:
        bulk_density = 1.6

    root = tk.Toplevel(parent or app)
    root.title("Longevity Calculator")
    root.configure(bg="#F0F0F0")
    # Mark transient so the popup tracks the main window; this also
    # makes Windows treat it as a child dialog (gets focus reliably).
    try:
        root.transient(parent or app)
    except Exception:
        pass
    outer = tk.Frame(root, bg="#F0F0F0", padx=32, pady=24)
    outer.pack(fill="both", expand=True)

    # Centered title — matches the Excel popup
    tk.Label(outer, text="Longevity Calculator",
             font=FONT_TITLE, bg="#F0F0F0", fg="#000"
             ).pack(anchor="center", pady=(0, 24))

    # Compound + concentration inputs
    pfas_var = tk.StringVar(value=PFAS_OPTIONS[0])
    conc_var = tk.StringVar()

    tk.Label(outer, text="Select PFAS Compound:", font=FONT_HDR,
             bg="#F0F0F0").pack(anchor="w", pady=(4, 6))
    ttk.Combobox(outer, textvariable=pfas_var, values=PFAS_OPTIONS,
                 state="readonly", width=18, font=FONT_LABEL
                 ).pack(anchor="w", padx=24)

    tk.Label(outer, text="Concentration (ug/L):", font=FONT_HDR,
             bg="#F0F0F0").pack(anchor="w", pady=(20, 6))
    tk.Entry(outer, textvariable=conc_var, font=FONT_VAL, width=18,
             relief="solid", bd=1
             ).pack(anchor="w", padx=24)

    # Live-app parameters -----------------------------------------------
    # Excel popup labels this "Parameters from Excel:" — keep that wording
    tk.Label(outer, text="Parameters from Excel:",
             font=FONT_HDR, bg="#F0F0F0"
             ).pack(anchor="w", pady=(24, 6))
    params_frame = tk.Frame(outer, bg="#F0F0F0")
    params_frame.pack(anchor="w", padx=24, fill="x")
    rows = [
        ("fcac:",                        f"{fcac:.6f}"),
        ("Width (m):",                   f"{width:.2f}"),
        ("Bulk Density (kg/m³):",   f"{bulk_density * 1000:.3f}"),
        ("Groundwater Velocity (m/yr):", f"{gw_velocity:.2f}"),
    ]
    for lbl, val in rows:
        row = tk.Frame(params_frame, bg="#F0F0F0")
        row.pack(anchor="w", pady=3)
        tk.Label(row, text=lbl, font=FONT_LABEL, bg="#F0F0F0",
                 width=28, anchor="w").pack(side="left")
        tk.Label(row, text=val, font=FONT_LABEL, bg="#F0F0F0"
                 ).pack(side="left")

    # Separator line + Results section
    sep = tk.Frame(outer, height=2, bg="#888")
    sep.pack(fill="x", pady=(28, 14))
    res_label = tk.Label(outer, text="Calculation Results:",
                          font=FONT_RES, bg="#F0F0F0")
    res_label.pack(anchor="w")
    res_content = tk.Frame(outer, bg="#F0F0F0")
    res_content.pack(anchor="w", padx=24, fill="x", pady=(8, 0))

    def _calc():
        try:
            c = float(conc_var.get())
            if c <= 0:
                messagebox.showerror("Error",
                                     "Concentration must be greater than zero.",
                                     parent=root); return
        except ValueError:
            messagebox.showerror("Error",
                                 "Please enter a valid concentration.",
                                 parent=root); return
        if gw_velocity <= 0:
            messagebox.showerror("Error",
                                 "Groundwater velocity must be greater than zero.",
                                 parent=root); return

        coeff, exp = COEFFS.get(pfas_var.get(), COEFFS["PFOA"])
        longevity = (coeff * fcac * width * (bulk_density * 1000) *
                     (c ** exp)) / (gw_velocity / 365.0)

        # Persist
        try:
            if os.path.exists(txt_path):
                try:
                    os.chmod(txt_path, 0o666); os.remove(txt_path)
                except Exception: pass
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write("Longevity Calculator Results\n")
                f.write(f"PFAS Compound: {pfas_var.get()}\n")
                f.write(f"Concentration (ug/L): {c:.2f}\n")
                f.write(f"Longevity (years): {longevity:.2f}\n")
                f.write(f"fcac: {fcac:.6f}\n")
                f.write(f"Width (m): {width:.2f}\n")
                f.write(f"Bulk Density (kg/m^3): {bulk_density * 1000:.3f}\n")
                f.write(f"Groundwater Velocity (m/yr): {gw_velocity:.2f}\n")
                f.write(f"Coefficient: {coeff:.3f}\n")
                f.write(f"Concentration Exponent: {exp:.3f}\n")
        except Exception as e:
            print(f"[longevity] write failed: {e}")

        # Render
        for w in res_content.winfo_children():
            w.destroy()
        long_int = int(round(longevity))
        half_int = long_int // 2
        for txt, val in (
            ("Estimated longevity (no safety factor):", f"{long_int} years"),
            ("Estimated longevity (2x safety factor):", f"{half_int} years"),
        ):
            r = tk.Frame(res_content, bg="#F0F0F0")
            r.pack(anchor="w", pady=4)
            tk.Label(r, text=txt, font=FONT_LABEL, bg="#F0F0F0",
                     width=44, anchor="w").pack(side="left")
            tk.Label(r, text=val, font=FONT_BOLD, bg="#F0F0F0"
                     ).pack(side="left")
        link = tk.Label(res_content,
                        text="(Click here to see CAC paper for more details)",
                        font=FONT_SMALL, fg="#1B6EC2", cursor="hand2",
                        bg="#F0F0F0")
        link.pack(anchor="w", pady=(12, 0))
        link.bind("<Button-1>", lambda e: webbrowser.open(CAC_PAPER_URL))

    def _cancel():
        try: root.grab_release()
        except Exception: pass
        root.destroy()

    def _help():
        messagebox.showinfo(
            "Longevity Calculator Help",
            "Estimates breakthrough time for a Colloidal Activated Carbon\n"
            "(CAC) barrier for PFAS compounds.\n\n"
            "Inputs:\n"
            "  - PFAS Compound (PFOA / PFOS / PFHxS)\n"
            "  - Influent concentration (ug/L)\n\n"
            "Auto-loaded from current model:\n"
            "  - fcac (mass fraction of CAC)\n"
            "  - Width (m)\n"
            "  - Bulk density (g/ml)\n"
            "  - Groundwater velocity (m/yr)\n\n"
            "Compound-specific formula:\n"
            "  Longevity = coefficient * fcac * W * pb * C^exponent / Vd",
            parent=root)

    bar = tk.Frame(outer, bg="#F0F0F0")
    bar.pack(pady=(28, 8), anchor="center")
    tk.Button(bar, text="Calculate", width=11, font=FONT_BTN,
              command=_calc).pack(side="left", padx=10)
    tk.Button(bar, text="Cancel",    width=11, font=FONT_BTN,
              command=_cancel).pack(side="left", padx=10)
    tk.Button(bar, text="Help",      width=11, font=FONT_BTN,
              command=_help).pack(side="left", padx=10)

    root.protocol("WM_DELETE_WINDOW", _cancel)
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 48, 720)
    h = max(root.winfo_reqheight() + 36, 820)
    try:
        sw = root.winfo_screenwidth(); sh = root.winfo_screenheight()
        x = max(0, (sw - w) // 2); y = max(0, (sh - h) // 2 - 30)
        root.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        root.geometry(f"{w}x{h}")
    root.minsize(w, h); root.resizable(True, True)
    # Force the popup to the front and grab focus.  grab_set() must run
    # AFTER the window is mapped (deiconified), or Tk raises
    # "grab failed: window not viewable" and the popup ends up modeless.
    try:
        root.lift()
        root.focus_force()
        root.attributes("-topmost", True)
        root.after(400, lambda: root.attributes("-topmost", False))
    except Exception:
        pass
    try:
        root.grab_set()
    except Exception:
        pass
    root.wait_window()
