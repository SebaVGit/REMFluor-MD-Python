"""
popups_mass_discharge_import.py — pure-Python port.

Adapted from Source_Py/popups_mass_discharge_import.py.  Loads PFAS mass
discharge rate from a CSV (PFAS-LEACH style), converts to concentration
using v * area and writes 11 concentrations into Section-7 PFAA-1 or
PFAA-2 source-term cells of the live app (no xlsm coupling).

Called from main.run_script() when SourceOption2 is dispatched.
"""
from __future__ import annotations
import os
import csv
import tkinter as tk
from tkinter import ttk, messagebox, filedialog


RATE_COLUMN_NAMES = [
    "Vadose-zone mass discharge rate (ug/yr)",
    "Mean mass discharge rate (ug/yr)",
]
TIME_COLUMN = "Time (yr)"

FONT_TITLE = ("Arial", 14, "bold")
FONT_LABEL = ("Arial", 10)
FONT_BTN   = ("Arial", 11)


def _norm(s): return (s or "").replace(" ", "")
def _safe_float(v, d=None):
    try: return float(v)
    except (ValueError, TypeError): return d


def _find_header(headers, target):
    t = _norm(target)
    for h in headers:
        if h is not None and _norm(h) == t:
            return h
    return None


def _load_discharge_csv(path):
    if not path or not os.path.isfile(path):
        return None, None
    try:
        with open(path, "r", newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            headers = reader.fieldnames or []
            time_col = _find_header(headers, TIME_COLUMN)
            rate_col = None
            for n in RATE_COLUMN_NAMES:
                rate_col = _find_header(headers, n)
                if rate_col: break
            if not (time_col and rate_col):
                return None, None
            ts, rs = [], []
            for row in reader:
                t = _safe_float(row.get(time_col))
                r = _safe_float(row.get(rate_col))
                if t is not None and r is not None:
                    ts.append(t); rs.append(r)
            return (ts, rs) if ts else (None, None)
    except Exception:
        return None, None


def run(app, parent=None):
    # Inputs from app -------------------------------------------------------
    E12 = _safe_float(app.v_y_size.get(), 1.0)
    E13 = _safe_float(app.v_z_size.get())
    E22 = _safe_float(app.v_darcy.get())
    E18 = _safe_float(app.v_yr_start.get(), 1977)
    E19 = _safe_float(app.v_yr_end.get(),   2010)
    units_str = "feet" if (getattr(app, "v_units", None)
                           and app.v_units.get() == "feet") else "meters"

    if E13 is None or E22 is None or E22 == 0:
        messagebox.showerror(
            "Source Concentrations",
            "Section 2 'Z size' and Section 3 Darcy velocity must be set.",
            parent=parent or app)
        return False
    area = E12 * E13
    conv = 28.317 if units_str == "feet" else 1000.0
    denom = E22 * area * conv

    pfaa1_name = (app.v_pfaa1.get() or "PFAA1") if hasattr(app, "v_pfaa1") else "PFAA1"
    pfaa2_name = (app.v_pfaa2.get() or "PFAA2") if hasattr(app, "v_pfaa2") else "PFAA2"
    ncomp = 1 if str(pfaa2_name).strip().lower() in ("", "none") else 2

    root = tk.Toplevel(parent or app)
    root.title("Import Mass Discharge Rate to Source Term")
    root.configure(bg="#F0F0F0")
    try: root.withdraw()
    except Exception: pass
    # Skip transient() so the popup keeps a maximize button on Windows.
    try:
        root.grab_set()
    except Exception:
        pass
    main = tk.Frame(root, bg="#F0F0F0", padx=16, pady=14)
    main.pack(fill="both", expand=True)

    tk.Label(main,
             text="Import PFAS mass discharge rate (CSV) into source concentration",
             font=FONT_TITLE, bg="#F0F0F0").pack(anchor="w", pady=(0, 12))

    # CSV selector
    csv_var = tk.StringVar()
    row1 = tk.Frame(main, bg="#F0F0F0"); row1.pack(fill="x", pady=4)
    tk.Label(row1, text="Discharge CSV:", font=FONT_LABEL, bg="#F0F0F0",
             width=18, anchor="w").pack(side="left")
    tk.Entry(row1, textvariable=csv_var, width=50, font=FONT_LABEL
             ).pack(side="left", padx=4, fill="x", expand=True)
    def _browse():
        p = filedialog.askopenfilename(
            title="Select mass discharge rate CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            parent=root)
        if p: csv_var.set(p)
    tk.Button(row1, text="Browse...", command=_browse,
              font=FONT_LABEL).pack(side="left")

    # Target column
    target_var = tk.StringVar(value="V")
    row2 = tk.Frame(main, bg="#F0F0F0"); row2.pack(fill="x", pady=8)
    tk.Label(row2, text="Fill PFAS column:", font=FONT_LABEL, bg="#F0F0F0",
             width=18, anchor="w").pack(side="left")
    tk.Radiobutton(row2, text=str(pfaa1_name), variable=target_var, value="V",
                   font=FONT_LABEL, bg="#F0F0F0").pack(side="left", padx=6)
    tk.Radiobutton(row2,
                   text=str(pfaa2_name) if ncomp == 2 else "PFAA2 (none)",
                   variable=target_var, value="X",
                   state="normal" if ncomp == 2 else "disabled",
                   font=FONT_LABEL, bg="#F0F0F0").pack(side="left", padx=6)

    span = int(E19) - int(E18) + 1 if (E18 and E19) else 33
    tk.Label(main,
             text=f"Model run: {int(E18)} - {int(E19)} ({span} years).",
             font=("Arial", 10, "bold"), bg="#F0F0F0"
             ).pack(anchor="w", pady=(12, 4))
    tk.Label(main,
             text="Choose which CSV time segment to use for the source term:",
             font=FONT_LABEL, bg="#F0F0F0").pack(anchor="w")

    yr_var = tk.StringVar(value="first")
    custom_start = tk.StringVar(value="1")
    row3 = tk.Frame(main, bg="#F0F0F0"); row3.pack(fill="x", pady=4)
    tk.Radiobutton(row3, text=f"First {span} years (from CSV start)",
                   variable=yr_var, value="first", font=FONT_LABEL,
                   bg="#F0F0F0").pack(anchor="w")
    tk.Radiobutton(row3, text=f"Last {span} years (from CSV end)",
                   variable=yr_var, value="last", font=FONT_LABEL,
                   bg="#F0F0F0").pack(anchor="w")
    row3b = tk.Frame(main, bg="#F0F0F0"); row3b.pack(fill="x", pady=2)
    tk.Radiobutton(row3b, text="Custom start:", variable=yr_var, value="custom",
                   font=FONT_LABEL, bg="#F0F0F0").pack(side="left")
    tk.Entry(row3b, textvariable=custom_start, width=8, font=FONT_LABEL
             ).pack(side="left", padx=4)

    def _apply():
        csv_path = csv_var.get().strip()
        if not csv_path or not os.path.isfile(csv_path):
            messagebox.showerror("Error",
                                 "Please select a valid mass discharge CSV.",
                                 parent=root); return
        ts, rs = _load_discharge_csv(csv_path)
        if not ts or not rs:
            messagebox.showerror(
                "Error",
                "CSV must have 'Time (yr)' and one of:\n  "
                + "\n  ".join(RATE_COLUMN_NAMES),
                parent=root); return

        csv_min = min(ts); csv_max = max(ts)
        choice = yr_var.get()
        if choice == "first":
            csv_start = csv_min
        elif choice == "last":
            csv_start = max(csv_min, csv_max - span)
        else:
            try:
                csv_start = float(custom_start.get())
            except ValueError:
                messagebox.showerror("Error", "Start must be numeric.",
                                     parent=root); return
        csv_start = max(csv_start, csv_min)
        csv_end = csv_max
        if csv_start >= csv_end:
            messagebox.showerror("Error",
                                 "CSV start is at or past the CSV end.",
                                 parent=root); return

        # Read the 11 source years from app.v_src_years
        try:
            excel_times = [float(v.get()) for v in app.v_src_years[:11]]
        except (ValueError, AttributeError):
            messagebox.showerror("Error",
                                 "Section 7 source years must all be numeric.",
                                 parent=root); return
        excel_updated = [t - excel_times[0] + csv_start for t in excel_times]

        # Restrict CSV to selected segment, sort
        seg = sorted([(t, r) for t, r in zip(ts, rs) if csv_start <= t <= csv_end])
        if not seg:
            messagebox.showerror("Error",
                                 "No CSV rows in selected segment.",
                                 parent=root); return
        seg_t = [p[0] for p in seg]; seg_r = [p[1] for p in seg]

        def _closest(target):
            best_i, best_d = 0, abs(seg_t[0] - target)
            for i in range(1, len(seg_t)):
                d = abs(seg_t[i] - target)
                if d < best_d: best_i, best_d = i, d
            return best_i

        concs = [seg_r[_closest(t)] / denom for t in excel_updated]

        # Write into the matching app source list
        target = app.v_src_pfaa1 if target_var.get() == "V" else app.v_src_pfaa2
        for i, c in enumerate(concs):
            if i < len(target):
                target[i].set(f"{c:,.3f}")

        messagebox.showinfo(
            "Done",
            f"11 concentrations written to {target_var.get()}-column "
            f"(Section 7 rows 8-18).",
            parent=root)
        try: root.grab_release()
        except Exception: pass
        root.destroy()

    def _cancel():
        try: root.grab_release()
        except Exception: pass
        root.destroy()

    bar = tk.Frame(main, bg="#F0F0F0"); bar.pack(pady=(20, 0))
    tk.Button(bar, text="Apply",  width=10, font=FONT_BTN,
              command=_apply).pack(side="left", padx=6)
    tk.Button(bar, text="Cancel", width=10, font=FONT_BTN,
              command=_cancel).pack(side="left", padx=6)

    root.protocol("WM_DELETE_WINDOW", _cancel)
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 32, 760)
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
