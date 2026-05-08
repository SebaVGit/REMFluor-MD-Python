"""
popups_cellsize.py — pure-Python port.

Adapted from Source_Py/popups_cellsize.py.  Persists user-defined grid
cell sizes (dX, dY, dZ) to cellsize_input.txt.  Reads model size and unit
flag from the live app StringVars (E11/E12/E13, AD1, A8) instead of xlsm.

Called from main.run_script() when OpenAppendix_2_1_Relative is dispatched
(via the EXE shortcut) — this is the user's "Optional: Enter user defined
size of grid cells" button in Section 2.
"""
from __future__ import annotations
import os
import tkinter as tk
from tkinter import messagebox

from .state import get_state


FONT_TITLE  = ("Arial", 16, "bold")
FONT_LABEL  = ("Arial", 11)
FONT_BTN    = ("Arial", 11)
FONT_VAL    = ("Arial", 11)
FONT_SMALL  = ("Arial", 10)


def _safe_float(s, default=None):
    try: return float(str(s).strip())
    except (ValueError, TypeError): return default


def _load_existing(path: str):
    out = {}
    if not os.path.exists(path):
        return out
    try:
        with open(path) as f:
            for ln in f:
                ln = ln.strip()
                if "," not in ln:
                    continue
                parts = ln.split(",")
                if len(parts) != 2 or parts[0] == "Parameter":
                    continue
                key = parts[0].strip()
                try: out[key] = float(parts[1].strip())
                except ValueError: pass
    except Exception:
        pass
    return out


def run(app, parent=None):
    state = get_state()
    work_dir = state.work_dir or os.getcwd()
    txt_path = os.path.join(work_dir, "cellsize_input.txt")
    existing = _load_existing(txt_path)

    # Inputs from app -------------------------------------------------------
    is_simple = (getattr(app, "active_sheet", "Simple") != "Detailed_2")
    units_str = "feet" if (getattr(app, "v_units", None)
                           and app.v_units.get() == "feet") else "meters"
    Lx = _safe_float(app.v_x_size.get())
    Ly = _safe_float(app.v_y_size.get())
    Lz = _safe_float(app.v_z_size.get())

    root = tk.Toplevel(parent or app)
    root.title("Enter Grid Cell Sizes")
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

    tk.Label(outer, text="Enter Grid Cell Sizes",
             font=FONT_TITLE, bg="#F0F0F0"
             ).grid(row=0, column=0, columnspan=4, pady=(0, 10), sticky="w")
    tk.Label(outer, text=f"Units: {units_str}",
             font=FONT_SMALL, fg="gray", bg="#F0F0F0"
             ).grid(row=1, column=0, columnspan=4, pady=(0, 12), sticky="w")

    sv_dx = tk.StringVar(value=str(existing.get("Cell Size X:", "")))
    sv_dy = tk.StringVar(value=str(existing.get("Cell Size Y:", "")))
    sv_dz = tk.StringVar(value=str(existing.get("Cell Size Z:", "")))

    label_nx = tk.Label(outer, text="", font=FONT_SMALL, fg="gray", bg="#F0F0F0")
    label_ny = tk.Label(outer, text="", font=FONT_SMALL, fg="gray", bg="#F0F0F0")
    label_nz = tk.Label(outer, text="", font=FONT_SMALL, fg="gray", bg="#F0F0F0")

    # Row X ---------------------------------------------------------------
    tk.Label(outer, text="Cell Size X:", font=FONT_LABEL, bg="#F0F0F0"
             ).grid(row=2, column=0, sticky="w", padx=4, pady=8)
    tk.Entry(outer, textvariable=sv_dx, width=12, font=FONT_VAL
             ).grid(row=2, column=1, padx=4, pady=8)
    tk.Label(outer, text=units_str, font=FONT_LABEL, bg="#F0F0F0"
             ).grid(row=2, column=2, padx=4, sticky="w")
    label_nx.grid(row=2, column=3, padx=4, sticky="w")

    if not is_simple:
        tk.Label(outer, text="Cell Size Y:", font=FONT_LABEL, bg="#F0F0F0"
                 ).grid(row=3, column=0, sticky="w", padx=4, pady=8)
        tk.Entry(outer, textvariable=sv_dy, width=12, font=FONT_VAL
                 ).grid(row=3, column=1, padx=4, pady=8)
        tk.Label(outer, text=units_str, font=FONT_LABEL, bg="#F0F0F0"
                 ).grid(row=3, column=2, padx=4, sticky="w")
        label_ny.grid(row=3, column=3, padx=4, sticky="w")
        row_z = 4
    else:
        row_z = 3

    tk.Label(outer, text="Cell Size Z:", font=FONT_LABEL, bg="#F0F0F0"
             ).grid(row=row_z, column=0, sticky="w", padx=4, pady=8)
    tk.Entry(outer, textvariable=sv_dz, width=12, font=FONT_VAL
             ).grid(row=row_z, column=1, padx=4, pady=8)
    tk.Label(outer, text=units_str, font=FONT_LABEL, bg="#F0F0F0"
             ).grid(row=row_z, column=2, padx=4, sticky="w")
    label_nz.grid(row=row_z, column=3, padx=4, sticky="w")

    def _update_counts(*_):
        try:
            dx = _safe_float(sv_dx.get(), 0); dy = _safe_float(sv_dy.get(), 0)
            dz = _safe_float(sv_dz.get(), 0)
            label_nx.config(text=f"-> {int(round(Lx/dx))} cells"
                            if Lx and dx > 0 else "")
            if is_simple:
                label_ny.config(text="-")
            else:
                label_ny.config(text=f"-> {int(round(Ly/dy))} cells"
                                if Ly and dy > 0 else "")
            label_nz.config(text=f"-> {int(round(Lz/dz))} cells"
                            if Lz and dz > 0 else "")
        except Exception:
            pass

    sv_dx.trace_add("write", _update_counts)
    sv_dy.trace_add("write", _update_counts)
    sv_dz.trace_add("write", _update_counts)
    _update_counts()

    def _save_and_exit():
        dx_s = sv_dx.get().strip(); dy_s = sv_dy.get().strip()
        dz_s = sv_dz.get().strip()
        if not dx_s or not dz_s:
            messagebox.showerror("Error",
                "Please enter values for cell sizes X and Z.", parent=root)
            return
        if not is_simple and not dy_s:
            messagebox.showerror("Error",
                "Please enter values for all cell sizes (X, Y, Z).",
                parent=root); return
        try:
            dx = float(dx_s)
            default_dy = 16.40 if units_str == "feet" else 5.0
            dy = float(dy_s) if (not is_simple and dy_s) else default_dy
            dz = float(dz_s)
        except ValueError:
            messagebox.showerror("Error",
                "Invalid numeric values.  Please enter valid numbers.",
                parent=root); return

        unit_flag = 1 if units_str == "feet" else 2
        if os.path.exists(txt_path):
            try:
                os.chmod(txt_path, 0o666); os.remove(txt_path)
            except Exception: pass
        with open(txt_path, "w") as f:
            f.write("Grid Cell Sizes\n")
            f.write("Parameter,Value\n")
            f.write(f"Cell Size X:,{dx}\n")
            f.write(f"Cell Size Y:,{dy}\n")
            f.write(f"Cell Size Z:,{dz}\n")
            f.write(f"Unit Flag:,{unit_flag}\n")

        msg = (f"Cell sizes saved.\n\n"
               f"Cell Size X: {dx} {units_str}\n")
        if not is_simple:
            msg += f"Cell Size Y: {dy} {units_str}\n"
        msg += f"Cell Size Z: {dz} {units_str}"
        messagebox.showinfo("Success", msg, parent=root)
        try: root.grab_release()
        except Exception: pass
        root.destroy()

    def _cancel():
        try: root.grab_release()
        except Exception: pass
        root.destroy()

    bar = tk.Frame(outer, bg="#F0F0F0")
    bar.grid(row=row_z + 2, column=0, columnspan=4, pady=(20, 0))
    tk.Button(bar, text="OK",     width=10, font=FONT_BTN,
              command=_save_and_exit).pack(side="left", padx=6)
    tk.Button(bar, text="Cancel", width=10, font=FONT_BTN,
              command=_cancel).pack(side="left", padx=6)

    root.protocol("WM_DELETE_WINDOW", _cancel)
    root.update_idletasks()
    w = max(root.winfo_reqwidth() + 32, 580)
    h = max(root.winfo_reqheight() + 24, 260)
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
