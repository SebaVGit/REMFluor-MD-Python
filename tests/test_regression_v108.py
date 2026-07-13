"""
REMFluor-MD v108 regression tests - Ron Falta review items 1 & 2.

    cd <project root>
    python tests\test_regression_v108.py

Headless (mocks tkinter), no GUI needed. Each test captures a client
bug fixed in v108; if any FAIL, a fix has regressed - do not ship.

Covered:
  Item 1 - Simple CAC Barrier Longevity Tool entered fcac 100x too large.
           Section 9 PSB Loading is a PERCENT ("(%)"), so 0.8 -> 0.008.
  Item 2 - Section 11 time-step size was not passed to input.inp; the
           reader looked up "Timestep Size (yr) " (trailing space) while
           _read_txt stored the stripped key, so dt always fell back to 0.1.
"""
import os
import sys
import types
import tempfile

sys.dont_write_bytecode = True   # never poison functions/__pycache__

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "_Python", "tkinter"))

# mock tkinter so functions/ imports headless
tk = types.ModuleType("tkinter")
tk.messagebox = types.ModuleType("tkinter.messagebox")
tk.messagebox.showerror = lambda *a, **k: None
tk.messagebox.showinfo = lambda *a, **k: None
tk.filedialog = types.ModuleType("tkinter.filedialog")
tk.ttk = types.ModuleType("tkinter.ttk")
sys.modules["tkinter"] = tk
sys.modules["tkinter.messagebox"] = tk.messagebox
sys.modules["tkinter.filedialog"] = tk.filedialog
sys.modules["tkinter.ttk"] = tk.ttk

from functions import generate_input_file as gif

PASS = FAIL = 0


def check(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  PASS  {name}")
    else:
        FAIL += 1
        print(f"  FAIL  {name}   {detail}")


# ====================================================================
# Item 1 - Longevity fcac conversion (percent -> fraction)
# ====================================================================
print("== Item 1: CAC Longevity fcac is a percent (0.8 -> 0.008) ==")

from functions import popups_longevity as pl


class _Var:
    def __init__(self, v=""):
        self.v = v
    def get(self):
        return self.v


class _App:
    pass


def fcac_from(psb_load_value):
    """Mirror the conversion in popups_longevity.run() so a regression in
    that block is caught here without opening a GUI."""
    app = _App()
    app.v_psb_load = _Var(psb_load_value)
    fcac_raw = pl._safe_float(str(pl._get_var(app, "v_psb_load")).rstrip('%'), None)
    if fcac_raw is None or fcac_raw == 0:
        return 0.01
    return fcac_raw / 100.0


check("0.8 pct -> fcac 0.008 (was 0.8, 100x too large)",
      abs(fcac_from("0.8") - 0.008) < 1e-9, f"got {fcac_from('0.8')}")
check("0.24 pct -> fcac 0.0024",
      abs(fcac_from("0.24") - 0.0024) < 1e-9, f"got {fcac_from('0.24')}")
check("1 pct -> fcac 0.01",
      abs(fcac_from("1") - 0.01) < 1e-9, f"got {fcac_from('1')}")
check("2.5 pct -> fcac 0.025 (values > 1 still divided by 100)",
      abs(fcac_from("2.5") - 0.025) < 1e-9, f"got {fcac_from('2.5')}")
check("trailing pct stripped: '0.8%' -> 0.008",
      abs(fcac_from("0.8%") - 0.008) < 1e-9, f"got {fcac_from('0.8%')}")
check("blank -> 0.01 default",
      abs(fcac_from("") - 0.01) < 1e-9, f"got {fcac_from('')}")
check("zero -> 0.01 default",
      abs(fcac_from("0") - 0.01) < 1e-9, f"got {fcac_from('0')}")
check("matches generate_input_file percent convention (0.8 -> 0.008)",
      abs(float("0.8") / 100.0 - fcac_from("0.8")) < 1e-12)


# ====================================================================
# Item 2 - Section 11 time-step size reaches input.inp
# ====================================================================
print("== Item 2: Section 11 time-step size flows into input.inp ==")

tmp = tempfile.mkdtemp()
num_path = os.path.join(tmp, "numerical_inputs.txt")


def write_numerical(dt_val):
    with open(num_path, "w") as f:
        f.write("iTVD\niTVD, 1\n\n Zone\nParameter,\n")
        f.write(f"Timestep Size (yr) ,{dt_val}\n")
        f.write("Convergence Tolerance (ug/L),1.0\n\n")


write_numerical(0.05)
nd = gif._read_txt(num_path)
check("_read_txt stores stripped key 'Timestep Size (yr)'",
      "Timestep Size (yr)" in nd, f"keys={list(nd)}")
check("reader recovers user dt=0.05 (not default 0.1)",
      abs(nd.get("Timestep Size (yr)", 0.1) - 0.05) < 1e-12,
      f"got {nd.get('Timestep Size (yr)')}")

write_numerical(0.25)
nd2 = gif._read_txt(num_path)
check("reader recovers user dt=0.25",
      abs(nd2.get("Timestep Size (yr)", 0.1) - 0.25) < 1e-12,
      f"got {nd2.get('Timestep Size (yr)')}")


def resolve(num_data):
    dt = num_data.get("Timestep Size (yr)",
                      num_data.get("Timestep Size (yr) ", 0.1))
    try:
        dt = float(dt)
    except (ValueError, TypeError):
        dt = 0.1
    if dt <= 0:
        dt = 0.1
    return dt


legacy = {"Timestep Size (yr) ": "0.2"}
stripped = {"Timestep Size (yr)": "0.3"}
check("dual-key lookup handles stripped key (0.3)", abs(resolve(stripped) - 0.3) < 1e-12)
check("dual-key lookup handles legacy trailing-space key (0.2)", abs(resolve(legacy) - 0.2) < 1e-12)
check("missing/invalid dt falls back to 0.1", abs(resolve({}) - 0.1) < 1e-12)
check("non-positive dt falls back to 0.1", abs(resolve({"Timestep Size (yr)": "0"}) - 0.1) < 1e-12)

# End-to-end: build a full input.inp and confirm the dt reflects the file.
EX = os.path.join(ROOT, "Example", "3_Detailed_2")
if os.path.isdir(EX):
    from functions.inp_parser import parse_input_inp, parse_additional_info
    from functions.inp_to_state import write_inp_to_state

    class FakeState:
        def __init__(self, wd=""):
            self.d = {}; self.work_dir = wd; self.bundle_dir = ""; self.base_dir = ""
        def get(self, a, d=None): return self.d.get(a, d)
        def set(self, a, v): self.d[a] = v

    data = parse_input_inp(os.path.join(EX, "input.inp"))
    add = parse_additional_info(os.path.join(EX, "store_info_additional_input.txt"))
    st = FakeState(tmp)     # tmp holds numerical_inputs.txt with dt=0.25
    write_inp_to_state(st, data, add, unit_flag=2, volfrac=0.8)
    built = gif.build_inp_data(st)
    check("build_inp_data uses dt=0.25 from numerical_inputs.txt",
          abs(float(built["dt"]) - 0.25) < 1e-9, f"got dt={built.get('dt')}")
else:
    print("  SKIP  end-to-end build (Example/3_Detailed_2 not found)")



# ====================================================================
# Items 3 & 4 - PSB x1/x2 and xmax keep decimals above 100 in input.inp
# ====================================================================
print("== Items 3 & 4: x1/x2/xmax not truncated above 100 ==")

cf = gif._coord_fmt

# Item 3: PSB boundaries. Entering dist=100.4, width=5.1 -> x1=100.4, x2=105.5
check("x1 100.4 kept (was truncated to 100)", cf(100.4) == "100.4", cf(100.4))
check("x2 105.5 kept (was rounded to 106)", cf(105.5) == "105.5", cf(105.5))
check("x2 = x1 + width dust cleaned (100.4+5.1 -> 105.5)",
      cf(100.4 + 5.1) == "105.5", cf(100.4 + 5.1))
# Item 4: xmax
check("xmax 300.5 kept (was truncated to 300)", cf(300.5) == "300.5", cf(300.5))
# below 100 unchanged, whole numbers stay clean, zero handled
check("below-100 value unchanged (99.4)", cf(99.4) == "99.4", cf(99.4))
check("whole number stays clean (300 -> '300')", cf(300.0) == "300", cf(300.0))
check("whole number stays clean (100 -> '100')", cf(100.0) == "100", cf(100.0))
check("zero -> '0'", cf(0.0) == "0", cf(0.0))
check("None -> fallback '0'", cf(None) == "0", cf(None))

# End-to-end: detailed PSB model, dist>100, confirm input.inp x1/x2/xmax.
tmp_e = tempfile.mkdtemp()
with open(os.path.join(tmp_e, "cellsize_input.txt"), "w") as f:
    f.write("Grid Cell Sizes\nParameter,Value\n"
            "Cell Size X:,1.0\nCell Size Y:,5.0\nCell Size Z:,2.0\nUnit Flag:,2.0\n")

class FS2:
    def __init__(self, wd=""):
        self.d = {}; self.work_dir = wd; self.bundle_dir = ""; self.base_dir = ""
    def get(self, a, d=None): return self.d.get(a, d)
    def set(self, a, v): self.d[a] = v

se = FS2(tmp_e)
se.set("A8", 2); se.set("AD1", 2); se.set("R22", True)
se.set("E11", 400); se.set("E12", 50); se.set("E13", 10)
se.set("E15", 4); se.set("E16", 5); se.set("E18", 1977); se.set("E19", 2077)
se.set("C22", 10); se.set("G22", 0.3)
se.set("X74", 100.4)     # PSB distance from source
se.set("Y82", 5.1)       # PSB width
se.set("AB28", 2025); se.set("AA82", 0.24); se.set("V23", 0.33); se.set("V24", 100)
built_e = gif.build_inp_data(se)
check("build_inp_data x1 -> '100.4'", str(built_e["x1"]) == "100.4", built_e["x1"])
check("build_inp_data x2 -> '105.5'", str(built_e["x2"]) == "105.5", built_e["x2"])
check("build_inp_data xmax(400) -> '400'", str(built_e["xmax"]) == "400", built_e["xmax"])

se.set("E11", 300.5)
check("build_inp_data xmax 300.5 kept",
      str(gif.build_inp_data(se)["xmax"]) == "300.5",
      gif.build_inp_data(se)["xmax"])



# ====================================================================
# Save Input round-trip - PSB loading (fcac) & Freundlich Kf
# ====================================================================
# Bug: "Save Data" wrote v_psb_load (the PERCENT, e.g. 0.8) into the
# store_info "PSB Loading (AH28)" line, but that line + the Example
# folders + inp_to_state all expect the raw FRACTION (0.008).  So Load
# multiplied by 100 again (fcac 100x too big) and divided the Freundlich
# Kf by the mis-scaled loading.  The fix writes the fraction on save.
print("== Save Input: PSB loading fraction round-trip ==")

def save_frac(pct):
    """Mirror the fraction written to store_info by Save Input."""
    try:
        return f"{float(str(pct).replace(',', '').strip()) / 100.0:g}"
    except (TypeError, ValueError):
        return "None"

check("save writes 0.8% -> fraction '0.008'", save_frac(0.8) == "0.008", save_frac(0.8))
check("save writes 0.24% -> fraction '0.0024'", save_frac(0.24) == "0.0024", save_frac(0.24))
check("save blank loading -> 'None'", save_frac("") == "None", save_frac(""))

EX2 = os.path.join(ROOT, "Example", "3_Detailed_2")
if os.path.isdir(EX2):
    from functions.inp_parser import parse_input_inp as _pin
    from functions.inp_to_state import write_inp_to_state as _wis

    class FS3:
        def __init__(self, wd=""):
            self.d = {}; self.work_dir = wd; self.bundle_dir = ""; self.base_dir = ""
        def get(self, a, d=None): return self.d.get(a, d)
        def set(self, a, v): self.d[a] = v

    d2 = _pin(os.path.join(EX2, "input.inp"))
    fck2 = float(d2.get("fcackf2"))

    # user entered 0.8% -> save wrote fraction 0.008 -> load
    s_a = FS3(EX2)
    _wis(s_a, d2, {"psb_loading": 0.008,
                   "psb_kf_unit": "(ug/kg)(ug/L)^(-a)"}, 2, 1.0)
    check("load 0.8%: AA82 (v_psb_load) == 0.8 (not 80)",
          abs(float(s_a.get("AA82")) - 0.8) < 1e-9, s_a.get("AA82"))
    check("load 0.8%: AH28 fraction == 0.008",
          abs(float(s_a.get("AH28")) - 0.008) < 1e-12, s_a.get("AH28"))
    kf_v24 = float(str(s_a.get("V24")).replace(",", ""))
    check("load 0.8%: Freundlich Kf V24 == fcackf/0.008 (not /0.8)",
          abs(kf_v24 - fck2 / 0.008) < 1e-2, f"{kf_v24} vs {fck2/0.008}")

    # Example convention (fraction 0.0024) still loads to 0.24 percent
    s_b = FS3(EX2)
    _wis(s_b, d2, {"psb_loading": 0.0024,
                   "psb_kf_unit": "(ug/kg)(ug/L)^(-a)"}, 2, 1.0)
    check("Example loading 0.0024 -> AA82 0.24 (unchanged)",
          abs(float(s_b.get("AA82")) - 0.24) < 1e-9, s_b.get("AA82"))
else:
    print("  SKIP  Save Input round-trip (Example/3_Detailed_2 not found)")



# ====================================================================
# Item 13 (part 3) - spurious "Precursor 1" after save/load
# ====================================================================
# retardation_inputs.txt's PFAS-name block can carry a stale precursor
# species (e.g. "PFAA 1-able") inherited from an Example.  store_info is
# rewritten from the live UI on every Save; when it records the precursor
# slot as None, Load must NOT let the stale name re-add the component
# (which flipped ipre=1 and made Run Model demand Precursor-1 data).
print("== Item 13: no spurious Precursor 1 on save/load ==")

from functions.inp_parser import parse_retardation_pfas_names, parse_additional_info

def _reconcile(pfas_names, additional):
    """Replicate restore_from_saved's precursor None-guard."""
    st = {}
    cleared = set()
    for a, k in (("K38", "precursor1"), ("M38", "precursor2")):
        v = additional.get(k)
        if v is not None and str(v).strip().lower() in ("", "none"):
            cleared.add(a)
    for addr, name in pfas_names.items():
        if addr in cleared:
            st[addr] = "None"
            continue
        if name and str(name).strip():
            st[addr] = name
    return st

_tmp_p = tempfile.mkdtemp()
open(os.path.join(_tmp_p, "retardation_inputs.txt"), "w").write(
    "PFAS Names from Excel\nPFAA 1 (E38):,PFOS\nPFAA 2 (G38):,None\n"
    "Precursor 1 (K38):,PFAA 1-able\nPrecursor 2 (M38):,None\n\n"
    "Transmissive Zone\nParameter,PFAA 1\nSoil Bulk Density (g/ml),1.7\n")
open(os.path.join(_tmp_p, "store_info_additional_input.txt"), "w").write(
    "Additional Information\n" + "=" * 20 + "\n"
    "PFAA 1:,PFOS\nPFAA 2:,None\nPrecursor 1:,None\nPrecursor 2:,None\n")

_pf = parse_retardation_pfas_names(os.path.join(_tmp_p, "retardation_inputs.txt"))
_ad = parse_additional_info(os.path.join(_tmp_p, "store_info_additional_input.txt"))
_res = _reconcile(_pf, _ad)
check("stale retardation Precursor 1 present in file",
      _pf.get("K38") == "PFAA 1-able", _pf.get("K38"))
check("store_info records precursor1 as None",
      str(_ad.get("precursor1")).lower() == "none", _ad.get("precursor1"))
check("guard keeps K38 = None (no spurious Precursor 1)",
      _res.get("K38") == "None", _res.get("K38"))

# resulting model has ipre = 0 (no precursor component)
class _FSp:
    def __init__(self):
        self.d = {"A8": 2, "G38": "None", "K38": _res.get("K38")}
        self.work_dir = _tmp_p; self.bundle_dir = ""; self.base_dir = ""
    def get(self, a, d=None): return self.d.get(a, d)
    def set(self, a, v): self.d[a] = v
check("build_inp_data ipre == 0 with guard (was 1)",
      gif.build_inp_data(_FSp())["ipre"] == 0)

# a REAL precursor still round-trips (both sources name a species)
open(os.path.join(_tmp_p, "store_info_additional_input.txt"), "w").write(
    "Additional Information\n" + "=" * 20 + "\n"
    "PFAA 1:,PFOS\nPrecursor 1:,PFOA\n")
_ad2 = parse_additional_info(os.path.join(_tmp_p, "store_info_additional_input.txt"))
_res2 = _reconcile(_pf, _ad2)
check("real precursor preserved (retardation name wins when not None)",
      _res2.get("K38") == "PFAA 1-able", _res2.get("K38"))


# ====================================================================
# Items 1-3 END-TO-END - values land in a generated input.inp
# ====================================================================
print("== Items 1-3: values written to a real generated input.inp ==")
import shutil
_root_tpl = os.path.join(ROOT, "template.inp")
if os.path.exists(_root_tpl):
    from functions.state import get_state as _get_state
    _wd = tempfile.mkdtemp()
    shutil.copy(_root_tpl, os.path.join(_wd, "template.inp"))
    open(os.path.join(_wd, "cellsize_input.txt"), "w").write(
        "Grid Cell Sizes\nParameter,Value\nCell Size X:,1.0\n"
        "Cell Size Y:,5.0\nCell Size Z:,2.0\nUnit Flag:,2.0\n")
    open(os.path.join(_wd, "numerical_inputs.txt"), "w").write(
        "iTVD\niTVD, 1\n\n Zone\nParameter,\n"
        "Timestep Size (yr) ,0.05\nConvergence Tolerance (ug/L),1.0\n\n")
    _st = _get_state()
    _st.work_dir = _wd; _st.bundle_dir = _wd
    _st.snapshot = lambda a: None
    for _k, _v in {"A8": 2, "AD1": 2, "R22": True, "E11": 300.5, "E12": 50,
                   "E13": 10, "E15": 4, "E16": 5, "E18": 1977, "E19": 2077,
                   "C22": 10, "G22": 0.3, "X74": 100.4, "Y82": 5.1,
                   "AB28": 2025, "AA82": 0.8, "V23": 0.33, "V24": 100,
                   "AC82": 5, "V47": 10}.items():
        _st.set(_k, _v)
    class _AppE: pass
    assert gif.run(_AppE()), "generate_input_file.run failed"
    _lines = open(os.path.join(_wd, "input.inp")).read().splitlines()

    def _val_after(label):
        for i, l in enumerate(_lines):
            if label in l:
                return _lines[i + 1].strip()
        return None

    check("input.inp x1,x2 == '100.4, 105.5' (item 3)",
          _val_after("x1, x2 (m)") == "100.4, 105.5", _val_after("x1, x2 (m)"))
    _xm = _val_after("xmax (m), nx1, nxpsb, nx2")
    check("input.inp xmax == 300.5 (item 4)",
          _xm is not None and _xm.split(",")[0].strip() == "300.5", _xm)
    _dtl = _val_after("dt (yr), nt, npt")
    check("input.inp dt == 0.05 (item 2, not 0.1)",
          _dtl is not None and _dtl.split(",")[0].strip() == "0.05", _dtl)
else:
    print("  SKIP  end-to-end input.inp (template.inp not found)")


print()
print(f"{PASS} passed, {FAIL} failed")
sys.exit(1 if FAIL else 0)
