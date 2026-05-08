"""
popups_retardation.py — standalone replacement for popups_retardation.exe.

Adapted from Source_Py/popups_retardation.py.  Layout, fonts, button set, and
help-link wiring mirror the original .exe popup 1:1, but no xlwings/openpyxl
coupling — PFAA names, porosities, and computed R-values are read/written
through the live tkinter app StringVars instead.

Called from main.run_script() when CalculrateRetardationFactors is dispatched.
"""
from __future__ import annotations
import os
import platform
import subprocess
import webbrowser
import tkinter as tk
from tkinter import messagebox

from .state import get_state


# ── Compound lookup (mirrors Source_Py/popups_retardation.py) ─────────────
KOC_VALUES = {
    "PFOS":           631,
    "PFOA":           200,
    "PFHxS":          126,
    "PFHxA":          40,
    "PFBS":           35,
    "PFNA":           398,
    "USER-SPECIFIED": "",
}

PARAMETERS = ["Soil Bulk Density (g/ml)", "foc (-)", "Koc (L/kg)"]
ZONES      = ["Transmissive", "Low-K"]


# ── Fonts (sized for high-DPI displays — matches original .exe scale) ─────
FONT_TITLE  = ("Arial", 16, "bold")
FONT_LABEL  = ("Arial", 11)
FONT_HEADER = ("Arial", 11, "bold")
FONT_ZONE   = ("Arial", 13, "underline")
FONT_HELP   = ("Arial", 10, "bold")
FONT_BTN    = ("Arial", 11)
FONT_VAL    = ("Arial", 11)


# ── helpers ────────────────────────────────────────────────────────────────
def _detect_compound(value):
    if not value:
        return None
    s = str(value).upper().strip()
    if "USER" in s and "SPECIFIED" in s:
        return "USER-SPECIFIED"
    if "PFOS" in s:  return "PFOS"
    if "PFOA" in s:  return "PFOA"
    if "PFHXS" in s or "PFHX S" in s.replace("-", " "):
        return "PFHxS"
    if "PFHXA" in s or "PFHX A" in s.replace("-", " "):
        return "PFHxA"
    if "PFBS" in s:  return "PFBS"
    if "PFNA" in s:  return "PFNA"
    return None


def _read_existing(retard_file, precursors, pfas_map):
    foc_t = foc_l = rho_b = None
    koc_by_p = {}
    file_pfas = {}
    use = False
    if not os.path.exists(retard_file):
        return use, foc_t, foc_l, rho_b, koc_by_p, file_pfas
    try:
        with open(retard_file) as f:
            lines = [ln.strip() for ln in f if ln.strip()]
    except Exception:
        return use, foc_t, foc_l, rho_b, koc_by_p, file_pfas

    in_pfas = False
    for ln in lines:
        if "PFAS Names from Excel" in ln:
            in_pfas = True; continue
        if in_pfas and ":" in ln:
            for tag, key in (("PFAA 1 (E38):", "E38"),
                             ("PFAA 2 (G38):", "G38"),
                             ("Precursor 1 (K38):", "K38"),
                             ("Precursor 2 (M38):", "M38")):
                if ln.startswith(tag):
                    parts = ln.split(":", 1)
                    if len(parts) == 2:
                        file_pfas[key] = parts[1].strip().lstrip(",")

    def _norm(v):
        return "" if v is None else str(v).upper().strip()

    keys_to_check = list(pfas_map.keys())
    if keys_to_check:
        use = all(_norm(file_pfas.get(k, "")) == _norm(pfas_map.get(k, ""))
                  for k in keys_to_check)

    if not use:
        return use, foc_t, foc_l, rho_b, koc_by_p, file_pfas

    cur = None
    for ln in lines:
        if ln.endswith("Zone"):
            cur = ln.replace(" Zone", ""); continue
        if cur == "Transmissive":
            if ln.startswith("Soil Bulk Density"):
                try: rho_b = float(ln.split(",")[1])
                except Exception: pass
            elif ln.startswith("foc (-)"):
                try: foc_t = float(ln.split(",")[1])
                except Exception: pass
            elif ln.startswith("Koc (L/kg)"):
                parts = ln.split(",")
                for idx, p in enumerate(precursors):
                    if idx + 1 < len(parts):
                        try:
                            v = parts[idx + 1].strip()
                            if v:
                                koc_by_p[p] = float(v)
                        except Exception:
                            pass
        elif cur == "Low-K":
            if ln.startswith("foc (-)"):
                try: foc_l = float(ln.split(",")[1])
                except Exception: pass
    return use, foc_t, foc_l, rho_b, koc_by_p, file_pfas


def _docs_root(app):
    here = os.path.dirname(os.path.abspath(__file__))
    project = os.path.abspath(os.path.join(here, "..", "..", ".."))
    return os.path.join(project, "docs", "_site")


def _open_url(url):
    try:
        if platform.system() == "Windows":
            for exe in (r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"):
                if os.path.exists(exe):
                    subprocess.Popen([exe, url])
                    return
        webbrowser.open(url)
    except Exception:
        webbrowser.open(url)


def _open_help_section(app, section_id):
    f = os.path.join(_docs_root(app),
                     "data_chicklets", "Step5_PFASTransportProperties.html")
    if not os.path.exists(f):
        messagebox.showerror("Help Not Found",
                             f"Help file not found:\n{f}")
        return
    abs_p = os.path.abspath(f).replace("\\", "/")
    anchor = f"#{section_id}" if section_id else ""
    if os.name == "nt" and len(abs_p) > 1 and abs_p[1] == ":":
        url = f"file:///{abs_p}{anchor}"
    else:
        url = f"file://{abs_p}{anchor}"
    _open_url(url)


def _open_appendix(app, name):
    f = os.path.join(_docs_root(app), "appendix", name)
    if not os.path.exists(f):
        messagebox.showerror("Help Not Found",
                             f"Help file not found:\n{f}")
        return
    abs_p = os.path.abspath(f).replace("\\", "/")
    if os.name == "nt" and len(abs_p) > 1 and abs_p[1] == ":":
        url = f"file:///{abs_p}"
    else:
        url = f"file://{abs_p}"
    _open_url(url)


def _help_btn(parent, app, section_id):
    return tk.Button(parent, text="?", width=2, height=1, fg="red",
                     font=FONT_HELP, relief="raised",
                     command=lambda sid=section_id: _open_help_section(app, sid))


# ── Main entry ─────────────────────────────────────────────────────────────
def run(app, parent=None):
    state = get_state()
    work_dir = state.work_dir or os.getcwd()
    retard_file = os.path.join(work_dir, "retardation_inputs.txt")

    is_detailed = (getattr(app, "active_sheet", "Simple") == "Detailed_2")
    precursors = ["PFAA 1", "PFAA 2", "Precursors 1", "Precursors 2"] \
                 if is_detailed else ["PFAA 1", "PFAA 2"]

    pfas_map = {"E38": app.v_pfaa1.get(), "G38": app.v_pfaa2.get()}
    if is_detailed:
        pfas_map["K38"] = app.v_pfaa3.get()
        pfas_map["M38"] = app.v_pfaa4.get()

    cell_to_p = {"E38": "PFAA 1", "G38": "PFAA 2",
                 "K38": "Precursors 1", "M38": "Precursors 2"}
    precursor_koc_default = {}
    for cell, pname in cell_to_p.items():
        if cell in pfas_map:
            comp = _detect_compound(pfas_map[cell])
            if comp:
                precursor_koc_default[pname] = (
                    "" if comp == "USER-SPECIFIED" else KOC_VALUES.get(comp, "")
                )

    use_existing, foc_t_ex, foc_l_ex, rho_b_ex, koc_ex, _ = \
        _read_existing(retard_file, precursors, pfas_map)

    def _safe_float(s, default=0.0):
        try: return float(s)
        except (ValueError, TypeError): return default
    porosity_tzone = _safe_float(app.v_porf.get(),     0.20)
    porosity_lowk  = _safe_float(app.v_lowk_por.get(), 0.48)

    # ── Build popup (hidden until sized) ───────────────────────────────
    root = tk.Toplevel(parent or app)
    root.title("Enter Parameters to Calculate Retardation Factor")
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

    n_cols = 2 + len(precursors)
    tk.Label(outer,
             text="Enter Parameters to Calculate Retardation Factor",
             font=FONT_TITLE, bg="#F0F0F0"
             ).grid(row=0, column=0, columnspan=n_cols, pady=(0, 18), sticky="w")

    input_entries = {z: {} for z in ZONES}
    r = 1

    tk.Label(outer, text="Soil Bulk Density (g/ml)",
             font=FONT_LABEL, bg="#F0F0F0", anchor="w"
             ).grid(row=r, column=0, sticky="w", pady=4)
    _help_btn(outer, app, "soil-bulk-density-in-t-zone"
              ).grid(row=r, column=1, padx=(8, 12))
    e_bd = tk.Entry(outer, width=12, font=FONT_VAL, justify="left")
    e_bd.grid(row=r, column=2, sticky="w", pady=4)
    default_bd = (str(rho_b_ex) if (use_existing and rho_b_ex is not None) else "1.7")
    e_bd.insert(0, default_bd)
    input_entries["Transmissive"][("Soil Bulk Density (g/ml)", None)] = e_bd
    r += 1

    for col, pname in enumerate(precursors, start=2):
        tk.Label(outer, text=pname, font=FONT_HEADER, bg="#F0F0F0"
                 ).grid(row=r, column=col, padx=4, pady=(8, 0))
    r += 1

    tk.Label(outer, text="Koc (L/kg)",
             font=FONT_LABEL, bg="#F0F0F0", anchor="w"
             ).grid(row=r, column=0, sticky="w", pady=4)
    _help_btn(outer, app, "organic-carbon-partitioning-coefficient"
              ).grid(row=r, column=1, padx=(8, 12))
    for col, pname in enumerate(precursors, start=2):
        e = tk.Entry(outer, width=12, font=FONT_VAL, justify="left")
        e.grid(row=r, column=col, padx=4, pady=4)
        if use_existing and pname in koc_ex:
            e.insert(0, str(koc_ex[pname]))
        elif pname in precursor_koc_default and precursor_koc_default[pname] != "":
            e.insert(0, str(precursor_koc_default[pname]))
        else:
            e.insert(0, "0.0")
        input_entries["Transmissive"][("Koc (L/kg)", pname)] = e
    r += 1

    tk.Label(outer, text="Transmissive Zone", font=FONT_ZONE, bg="#F0F0F0"
             ).grid(row=r, column=0, sticky="w", pady=(16, 4))
    r += 1

    tk.Label(outer, text="foc (-)", font=FONT_LABEL, bg="#F0F0F0", anchor="w"
             ).grid(row=r, column=0, sticky="w", pady=4)
    _help_btn(outer, app, "fraction-organic-carbon-in-t-zone"
              ).grid(row=r, column=1, padx=(8, 12))
    e_foc_t = tk.Entry(outer, width=12, font=FONT_VAL, justify="left")
    e_foc_t.grid(row=r, column=2, sticky="w", pady=4)
    default_foc_t = (str(foc_t_ex) if (use_existing and foc_t_ex is not None) else "0.001")
    e_foc_t.insert(0, default_foc_t)
    input_entries["Transmissive"][("foc (-)", None)] = e_foc_t
    r += 1

    tk.Label(outer, text="T-Zone porosity (-)",
             font=FONT_LABEL, bg="#F0F0F0", anchor="w"
             ).grid(row=r, column=0, sticky="w", pady=4)
    tk.Label(outer, text=f"{porosity_tzone}", fg="gray",
             font=FONT_VAL, bg="#F0F0F0"
             ).grid(row=r, column=2, sticky="w", pady=4)
    r += 1

    tk.Label(outer, text="Low-K Zone", font=FONT_ZONE, bg="#F0F0F0"
             ).grid(row=r, column=0, sticky="w", pady=(16, 4))
    r += 1

    tk.Label(outer, text="foc (-)", font=FONT_LABEL, bg="#F0F0F0", anchor="w"
             ).grid(row=r, column=0, sticky="w", pady=4)
    _help_btn(outer, app, "fraction-organic-carbon-in-low-k"
              ).grid(row=r, column=1, padx=(8, 12))
    e_foc_l = tk.Entry(outer, width=12, font=FONT_VAL, justify="left")
    e_foc_l.grid(row=r, column=2, sticky="w", pady=4)
    default_foc_l = (str(foc_l_ex) if (use_existing and foc_l_ex is not None) else "0.002")
    e_foc_l.insert(0, default_foc_l)
    input_entries["Low-K"][("foc (-)", None)] = e_foc_l
    r += 1

    tk.Label(outer, text="Low-K Zone porosity (-)",
             font=FONT_LABEL, bg="#F0F0F0", anchor="w"
             ).grid(row=r, column=0, sticky="w", pady=4)
    tk.Label(outer, text=f"{porosity_lowk}", fg="gray",
             font=FONT_VAL, bg="#F0F0F0"
             ).grid(row=r, column=2, sticky="w", pady=4)
    r += 1

    tk.Label(outer, text="", bg="#F0F0F0").grid(row=r, column=0, pady=8)
    r += 1

    def _read_float(entry):
        s = entry.get().strip()
        if not s: return 0.0
        try: return float(s)
        except ValueError: return 0.0

    def _save_and_apply():
        rho_b = _read_float(input_entries["Transmissive"][("Soil Bulk Density (g/ml)", None)])
        foc_t = _read_float(input_entries["Transmissive"][("foc (-)", None)])
        foc_l = _read_float(input_entries["Low-K"][("foc (-)", None)])
        koc_by_p = {p: _read_float(input_entries["Transmissive"][("Koc (L/kg)", p)])
                    for p in precursors}

        try:
            if os.path.exists(retard_file):
                try:
                    os.chmod(retard_file, 0o666); os.remove(retard_file)
                except Exception: pass
            with open(retard_file, "w") as f:
                f.write("PFAS Names from Excel\n")
                f.write(f"PFAA 1 (E38):,{pfas_map.get('E38','')}\n")
                f.write(f"PFAA 2 (G38):,{pfas_map.get('G38','')}\n")
                if is_detailed:
                    f.write(f"Precursor 1 (K38):,{pfas_map.get('K38','')}\n")
                    f.write(f"Precursor 2 (M38):,{pfas_map.get('M38','')}\n")
                f.write("\n")
                f.write("Transmissive Zone\n")
                f.write("Parameter," + ",".join(precursors) + "\n")
                f.write(f"Soil Bulk Density (g/ml),{rho_b}\n\n")
                f.write(f"foc (-),{foc_t}\n\n")
                f.write("Koc (L/kg)," + ",".join(str(koc_by_p[p]) for p in precursors) + "\n\n")
                f.write("Low-K Zone\n")
                f.write("Parameter," + ",".join(precursors) + "\n")
                f.write(f"foc (-),{foc_l}\n\n")
        except Exception as e:
            messagebox.showerror("Error",
                                 f"Could not write retardation_inputs.txt:\n{e}",
                                 parent=root)
            return

        ret_pairs = [
            (app.v_ret_trans1, app.v_ret_lowk1, "PFAA 1"),
            (app.v_ret_trans2, app.v_ret_lowk2, "PFAA 2"),
        ]
        if is_detailed:
            ret_pairs += [
                (app.v_ret_trans3, app.v_ret_lowk3, "Precursors 1"),
                (app.v_ret_trans4, app.v_ret_lowk4, "Precursors 2"),
            ]
        for ret_t, ret_l, pname in ret_pairs:
            koc = koc_by_p.get(pname, 0.0)
            if koc <= 0:
                ret_t.set(""); ret_l.set("")
                continue
            r_t = 1.0 + (rho_b * foc_t * koc) / porosity_tzone if porosity_tzone > 0 else 1.0
            r_l = 1.0 + (rho_b * foc_l * koc) / porosity_lowk  if porosity_lowk  > 0 else 1.0
            ret_t.set(f"{r_t:.1f}" if r_t > 1.0 else "")
            ret_l.set(f"{r_l:.1f}" if r_l > 1.0 else "")

        messagebox.showinfo("Success",
                            "Retardation factors updated.\n\n"
                            f"Saved: {os.path.basename(retard_file)}",
                            parent=root)
        try: root.grab_release()
        except Exception: pass
        root.destroy()

    def _cancel():
        try: root.grab_release()
        except Exception: pass
        root.destroy()

    def _input_help():  _open_appendix(app, "appendix_5_1.html")
    def _theory_help(): _open_appendix(app, "appendix_5_2.html")

    btn_bar = tk.Frame(outer, bg="#F0F0F0")
    btn_bar.grid(row=r, column=0, columnspan=n_cols, pady=(8, 0))
    tk.Button(btn_bar, text="OK",          width=10, font=FONT_BTN,
              command=_save_and_apply).pack(side="left", padx=6)
    tk.Button(btn_bar, text="Cancel",      width=10, font=FONT_BTN,
              command=_cancel).pack(side="left", padx=6)
    tk.Button(btn_bar, text="Input Help",  width=12, font=FONT_BTN,
              command=_input_help).pack(side="left", padx=6)
    tk.Button(btn_bar, text="Theory Help", width=12, font=FONT_BTN,
              command=_theory_help).pack(side="left", padx=6)

    root.protocol("WM_DELETE_WINDOW", _cancel)

    # ── Size & center NOW that all widgets are laid out ─────────────────
    root.update_idletasks()
    req_w = root.winfo_reqwidth()  + 32
    req_h = root.winfo_reqheight() + 24
    design_w = 980 if is_detailed else 760
    design_h = 600
    final_w = max(req_w, design_w)
    final_h = max(req_h, design_h)
    try:
        sw = root.winfo_screenwidth(); sh = root.winfo_screenheight()
        final_w = min(final_w, int(sw * 0.95))
        final_h = min(final_h, int(sh * 0.92))
        x = max(0, (sw - final_w) // 2)
        y = max(0, (sh - final_h) // 2 - 30)
        root.geometry(f"{final_w}x{final_h}+{x}+{y}")
    except Exception:
        root.geometry(f"{final_w}x{final_h}")
    root.minsize(final_w, final_h)
    root.resizable(True, True)
    try:
        root.deiconify(); root.lift(); root.focus_force()
    except Exception: pass

    root.wait_window()
