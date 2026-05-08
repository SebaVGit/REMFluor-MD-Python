"""
popups_source_remediation.py — pure-Python port of the Apply Remediation macro.

Source: Source_Py/popups_source_remediation.py.

Workflow (matches the .exe):
  1. Read v_src_rem_yr (start year) and v_src_conc_red (% reduction).
  2. Find the decade-row in v_src_years[0..10] whose year is closest to
     the start year — that's the first row to modify.
  3. From that row onward, multiply the corresponding PFAA-1 (v_src_pfaa1)
     and PFAA-2 (v_src_pfaa2) source concentrations by (1 - reduction%).
  4. (Detailed mode would also touch precursor columns Z and AB, but the
     standalone UI doesn't expose those lists yet — Simple-mode behaviour
     is identical to the original .exe.)

Called from main.run_script() when SourceRemediation is dispatched.
"""
from __future__ import annotations
from tkinter import messagebox


def _safe_float(s, default=None):
    if s is None:
        return default
    try:
        return float(str(s).strip())
    except (ValueError, TypeError):
        return default


def run(app, parent=None):
    # Inputs ------------------------------------------------------------
    start_year = _safe_float(app.v_src_rem_yr.get())
    reduction  = _safe_float(app.v_src_conc_red.get())

    if start_year is None:
        messagebox.showerror("Apply Remediation",
                             "Please enter Source Treatment Start Year (D27).",
                             parent=parent or app)
        return False
    if reduction is None:
        messagebox.showerror("Apply Remediation",
                             "Please enter Source Concentration Reduction % (D28).",
                             parent=parent or app)
        return False
    if reduction < 0 or reduction > 100:
        messagebox.showerror("Apply Remediation",
                             "Reduction percentage must be between 0 and 100.",
                             parent=parent or app)
        return False

    # Reduction is ALWAYS entered as a percent (the UI label literally
    # says "(%)"), so divide by 100 unconditionally.  The previous
    # branch
    #   factor = 1.0 - (reduction/100 if reduction > 1 else reduction)
    # collapsed §7 to zero whenever the user typed a value <= 1: e.g.
    # "1" was treated as the fraction 1.0, so factor = 1 - 1 = 0 and
    # every PFAA concentration multiplied by 0.  "0.5" got factor=0.5
    # — silently a 50% reduction instead of the intended 0.5%.
    factor = 1.0 - (reduction / 100.0)
    # Clamp for safety — validation above already guards 0..100, but
    # belt-and-suspenders against any future edge case.
    if factor < 0:
        factor = 0.0
    elif factor > 1:
        factor = 1.0

    # Find decade-row closest to start_year ----------------------------
    years = []
    for i, v in enumerate(app.v_src_years):
        y = _safe_float(v.get())
        if y is not None:
            years.append((i, y))
    if not years:
        messagebox.showerror("Apply Remediation",
                             "No source-term years found (Section 7).",
                             parent=parent or app)
        return False
    start_idx = min(years, key=lambda t: abs(t[1] - start_year))[0]

    # Apply reduction to PFAA-1 and PFAA-2 from start_idx onward -------
    n_changed = 0
    for col_attr in ("v_src_pfaa1", "v_src_pfaa2"):
        col_list = getattr(app, col_attr, None)
        if not col_list:
            continue
        for i in range(start_idx, len(col_list)):
            current = _safe_float(col_list[i].get())
            if current is None:
                continue
            new_val = current * factor
            # 2 decimals + comma grouping for cleaner display
            # ("1,584.00" instead of "1,584.000").
            col_list[i].set(f"{new_val:,.2f}")
            n_changed += 1

    messagebox.showinfo(
        "Apply Remediation",
        f"Reduction of {reduction:g}% applied to PFAA-1 / PFAA-2 source\n"
        f"concentrations from year {int(years[start_idx][1])} onward.\n\n"
        f"{n_changed} value(s) updated.",
        parent=parent or app)
    return True
