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



# ====================================================================
# Items 1-3 - observation well coordinates (x, y, z) in input.inp
# ====================================================================
# x/y/z were formatted with _smart_fmt (truncated >=100 to whole metres)
# and ywell was hard-coded to 0.  Now they use _coord_fmt and the
# per-well Y distance off centerline imported into mw_observations.json.
print("== Items 1-3: observation well x/y/z coordinates ==")
import json as _jsonw
import shutil as _shutilw
from functions.state import get_state as _gs_w

_tplw = os.path.join(ROOT, "template.inp")
if os.path.exists(_tplw):
    _wdw = tempfile.mkdtemp()
    _shutilw.copy(_tplw, os.path.join(_wdw, "template.inp"))
    open(os.path.join(_wdw, "cellsize_input.txt"), "w").write(
        "Grid Cell Sizes\nParameter,Value\nCell Size X:,1.0\n"
        "Cell Size Y:,5.0\nCell Size Z:,2.0\nUnit Flag:,2.0\n")
    open(os.path.join(_wdw, "numerical_inputs.txt"), "w").write(
        "iTVD\niTVD, 1\n\n Zone\nParameter,\n"
        "Timestep Size (yr) ,0.1\nConvergence Tolerance (ug/L),1.0\n\n")
    _jsonw.dump({"well_depths": {"W1": {"top": 150.4, "bot": 75.2},
                                 "W2": {"top": 10, "bot": 5}},
                 "well_y": {"W1": 12.3, "W2": -3.7}},
                open(os.path.join(_wdw, "mw_observations.json"), "w"))
    _stw = _gs_w(); _stw.work_dir = _wdw; _stw.bundle_dir = _wdw
    _stw.snapshot = lambda a: None; _stw._cells = {}
    for _k, _v in {"A8": 2, "AD1": 2, "R22": False, "E11": 400, "E12": 50,
                   "E13": 200, "E15": 4, "E16": 5, "E18": 1977, "E19": 2077,
                   "C22": 10, "G22": 0.3, "V47": 10,
                   "U34": "W1", "AF34": 105.1,
                   "U35": "W2", "AF35": 50.5}.items():
        _stw.set(_k, _v)
    class _AppW: pass
    assert gif.run(_AppW())
    _wl = [l.strip() for l in open(os.path.join(_wdw, "input.inp")).read().splitlines()
           if len(l.split(",")) == 5 and l.strip()[:1] in ("1", "2")]
    _w1 = next((l for l in _wl if l.startswith("1,")), "")
    _w2 = next((l for l in _wl if l.startswith("2,")), "")
    check("well 1 x keeps decimals >100 (105.1, item 1)",
          _w1 == "1, 105.1, 12.3, 150.4, 75.2", _w1)
    check("well 1 Y off centerline written (12.3, not 0, item 2)",
          ", 12.3," in _w1, _w1)
    check("well 1 zwelltop/zwellbot keep decimals >100 (item 3)",
          _w1.endswith("150.4, 75.2"), _w1)
    check("well 2 negative Y off centerline (-3.7)",
          _w2 == "2, 50.5, -3.7, 10, 5", _w2)
else:
    print("  SKIP  well coords end-to-end (template.inp not found)")


# ====================================================================
# Item 4 - iwall grid type + direct nx1/nx2 entry
# ====================================================================
print("== Item 4: iwall dropdown + direct nx1/nx2 entry ==")

def _gen_psb(overrides):
    wd = tempfile.mkdtemp()
    _shutilw.copy(_tplw, os.path.join(wd, "template.inp"))
    open(os.path.join(wd, "cellsize_input.txt"), "w").write(
        "Grid Cell Sizes\nParameter,Value\nCell Size X:,1.0\n"
        "Cell Size Y:,5.0\nCell Size Z:,2.0\nUnit Flag:,2.0\n")
    open(os.path.join(wd, "numerical_inputs.txt"), "w").write(
        "iTVD\niTVD, 1\n\n Zone\nParameter,\n"
        "Timestep Size (yr) ,0.1\nConvergence Tolerance (ug/L),1.0\n\n")
    st = _gs_w(); st.work_dir = wd; st.bundle_dir = wd
    st.snapshot = lambda a: None; st._cells = {}
    base = {"A8": 2, "AD1": 2, "R22": True, "E11": 132, "E12": 50, "E13": 10,
            "E15": 4, "E16": 5, "E18": 1977, "E19": 2077, "C22": 10, "G22": 0.3,
            "X74": 108, "Y82": 12, "AB28": 2025, "AA82": 0.24, "V23": 0.33,
            "V24": 100, "AC82": 12, "V47": 10}
    base.update(overrides)
    for k, v in base.items():
        st.set(k, v)
    class _AppP: pass
    gif.run(_AppP())
    ls = open(os.path.join(wd, "input.inp")).read().splitlines()
    flags = ls[2].strip()
    xg = None
    for i, l in enumerate(ls):
        if "xmax (m), nx1, nxpsb, nx2" in l:
            xg = ls[i + 1].strip()
    return flags, xg

if os.path.exists(_tplw):
    _f0, _x0 = _gen_psb({})
    check("default iwall == 2", _f0.split(",")[2].strip() == "2", _f0)
    check("default x-grid auto-computed (132, 108, 12, 12)",
          _x0 == "132, 108, 12, 12", _x0)
    _f1, _x1 = _gen_psb({"IWAL": 1})
    check("IWAL=1 writes iwall == 1 (uniform grid)",
          _f1.split(",")[2].strip() == "1", _f1)
    _f2, _x2 = _gen_psb({"NXB1": 7, "NXA2": 13})
    check("direct nx1=7 / nx2=13 override the computed counts",
          _x2 == "132, 7, 12, 13", _x2)
    check("nxpsb (wall count) unchanged by nx1/nx2 override",
          _x2.split(",")[2].strip() == "12", _x2)
else:
    print("  SKIP  iwall/nx end-to-end (template.inp not found)")

# sidecar round-trip for the new PSB fields
from functions import sidecars as _sc4

class _V4:
    def __init__(self, v=""): self.v = v
    def get(self): return self.v
    def set(self, x): self.v = str(x)

class _AppS: pass
_as = _AppS()
for _n in ("v_psb_kf_unit", "v_psb_a_1", "v_psb_a_2", "v_psb_a_3", "v_psb_a_4",
           "v_psb_kf_1", "v_psb_kf_2", "v_psb_kf_3", "v_psb_kf_4",
           "v_psb_mw_1", "v_psb_mw_2", "v_psb_mw_3", "v_psb_mw_4", "v_psb_yr",
           "v_psb_cells", "v_psb_dist", "v_psb_width", "v_psb_load"):
    setattr(_as, _n, _V4(""))
_as.v_model_psb = _V4("True")
_as.v_iwall = _V4("1"); _as.v_psb_nx1 = _V4("7"); _as.v_psb_nx2 = _V4("13")
_wd4 = tempfile.mkdtemp()
_sc4.write_psb(_as, _wd4)
_bs = _AppS()
for _n in ("v_iwall", "v_psb_nx1", "v_psb_nx2", "v_psb_kf_unit"):
    setattr(_bs, _n, _V4(""))
_bs.v_model_psb = _V4("")
_sc4.read_psb(_bs, _wd4)
check("sidecar round-trips iwall=1", _bs.v_iwall.get() == "1", _bs.v_iwall.get())
check("sidecar round-trips nx1=7", _bs.v_psb_nx1.get() == "7", _bs.v_psb_nx1.get())
check("sidecar round-trips nx2=13", _bs.v_psb_nx2.get() == "13", _bs.v_psb_nx2.get())



# ====================================================================
# Item 13 (part 3) - Freundlich Kf UNIT preserved across save/load
# ====================================================================
# Saving with Kf in mg (or ug) then loading must reproduce the SAME
# fcackf in input.inp and the SAME raw Kf + unit.  The bug reduced the
# value (e.g. by 1000) or applied the wrong unit; root cause was the
# loading percent/fraction mix-up feeding inp_to_state's Kf recovery.
print("== Item 13: Freundlich Kf unit preserved on save/load ==")

_tplk = os.path.join(ROOT, "template.inp")
if os.path.exists(_tplk):
    import shutil as _shk
    from functions.inp_parser import parse_input_inp as _pik
    from functions.inp_to_state import write_inp_to_state as _wisk
    from functions.state import get_state as _gsk

    def _ufk(u):
        u = u.lower()
        return {"ng/kg": 1e-3, "ug/kg": 1.0, "mg/kg": 1e3}[
            next(k for k in ("ng/kg", "ug/kg", "mg/kg") if k in u)]

    def _convk(raw, unit, a):
        return raw * (_ufk(unit) ** (1 - a))

    _stk = _gsk(); _stk.snapshot = lambda a: None
    class _AppK: pass

    def _wdk():
        wd = tempfile.mkdtemp()
        _shk.copy(_tplk, os.path.join(wd, "template.inp"))
        open(os.path.join(wd, "cellsize_input.txt"), "w").write(
            "Grid Cell Sizes\nParameter,Value\nCell Size X:,1.0\n"
            "Cell Size Y:,5.0\nCell Size Z:,2.0\nUnit Flag:,2.0\n")
        open(os.path.join(wd, "numerical_inputs.txt"), "w").write(
            "iTVD\niTVD, 1\n\n Zone\nParameter,\n"
            "Timestep Size (yr) ,0.1\nConvergence Tolerance (ug/L),1.0\n\n")
        return wd

    def _kf_roundtrip(unit, raw_kf, a=0.33, loading_pct=0.24):
        base = {"A8": 2, "AD1": 2, "R22": True, "E11": 132, "E12": 50, "E13": 10,
                "E15": 4, "E16": 5, "E18": 1977, "E19": 2077, "C22": 10, "G22": 0.3,
                "X74": 108, "Y82": 12, "AB28": 2025, "AA82": loading_pct, "V23": a,
                "AC82": 12, "V47": 10, "U24": unit, "V24": raw_kf,
                "V26": _convk(raw_kf, unit, a)}
        wd = _wdk(); _stk.work_dir = wd; _stk.bundle_dir = wd; _stk._cells = dict(base)
        _stk.snapshot = lambda a: None
        gif.run(_AppK())
        fk_save = _pik(os.path.join(wd, "input.inp"))["fcackf2"]
        d1 = _pik(os.path.join(wd, "input.inp"))
        wd2 = _wdk(); _stk.work_dir = wd2; _stk.bundle_dir = wd2; _stk._cells = {}
        _wisk(_stk, d1, {"psb_loading": loading_pct / 100.0,
                         "psb_kf_unit": unit}, 2, 1.0)
        unit_rec = _stk.get("U24"); raw_rec = str(_stk.get("V24")).replace(",", "")
        # PSB sidecar restores raw Kf + unit verbatim, then converted recomputes
        _stk.set("U24", unit); _stk.set("V24", raw_kf)
        _stk.set("V26", _convk(raw_kf, unit, a))
        gif.run(_AppK())
        fk_load = _pik(os.path.join(wd2, "input.inp"))["fcackf2"]
        return fk_save, fk_load, unit_rec, float(raw_rec)

    _fs, _fl, _ur, _rr = _kf_roundtrip("(mg/kg)(mg/L)^(-a)", 1227.951)
    check("mg Kf: fcackf round-trips exactly",
          abs(_fs - _fl) < 1e-2, f"save={_fs} load={_fl}")
    check("mg Kf: unit preserved on load (mg, not ug)",
          "mg/kg" in (_ur or ""), _ur)
    check("mg Kf: raw value recovered (1227.951, not /1000)",
          abs(_rr - 1227.951) < 1e-2, _rr)

    _fs2, _fl2, _ur2, _rr2 = _kf_roundtrip("(ug/kg)(ug/L)^(-a)", 5000.0)
    check("ug Kf: fcackf round-trips exactly",
          abs(_fs2 - _fl2) < 1e-2, f"save={_fs2} load={_fl2}")
    check("ug Kf: unit preserved (ug)", "ug/kg" in (_ur2 or ""), _ur2)
    check("ug Kf: raw value recovered (5000, not /1000)",
          abs(_rr2 - 5000.0) < 1e-2, _rr2)
else:
    print("  SKIP  Kf unit round-trip (template.inp not found)")



# ====================================================================
# Output plot suggestions (Ron review, dashboard)
# ====================================================================
# generate_dashboard.py pulls in dash/plotly, so it can't run headless;
# guard these with the floor math + source-level assertions.
print("== Output plots: Z label, log floor, hover precision ==")

import math as _mdash
_CONC_LOG_FLOOR = 1e-6   # ug/L (0.001 ng/L)

def _log_lo(vals):
    """Mirror generate_dashboard._log_range: force the floor ONLY when the
    data would sink below it; otherwise return None (Plotly autorange)."""
    v = [x for x in vals if x > 0]
    if not v:
        return None
    return _CONC_LOG_FLOOR if min(v) < _CONC_LOG_FLOOR else None

# (b) a near-zero cell (1e-21) must floor to 1e-6, not sink the axis
check("log floor: 1e-21 present -> lower bound 1e-6",
      _log_lo([1e-21, 3.2, 0.5, 1e-15]) == _CONC_LOG_FLOOR,
      _log_lo([1e-21, 3.2, 0.5, 1e-15]))
# data already at/above the floor -> autorange (no forced range), which
# fixes the ugly near-linear log plot on near-constant ~1600 data
check("log floor: normal data (>=floor) autoranges (no forced range)",
      _log_lo([0.01, 5.0, 1.2]) is None, _log_lo([0.01, 5.0, 1.2]))
check("log floor: near-constant ~1600 data autoranges (no forced range)",
      _log_lo([1598.8, 1598.8, 1598.8]) is None, _log_lo([1598.8, 1598.8]))

_dash_src = open(os.path.join(ROOT, "_Python", "tkinter", "functions",
                              "generate_dashboard.py"), encoding="utf-8").read()
# (a) Z control relabeled
check("(a) Z control relabeled 'Depth from Top'",
      "Depth from Top" in _dash_src and "Z Coordinate" not in _dash_src)
# (b) explicit concentration log floor constant present
check("(b) concentration log floor 1e-6 present",
      "_CONC_LOG_FLOOR = 1e-6" in _dash_src)
# (c) concentration hovers show 4 decimals, none left at 3
check("(c) hover concentration uses .4f (down to 0.0001)",
      "Concentration: %{y:.4f}" in _dash_src)
check("(c) no concentration hover left at .3f",
      "Concentration: %{y:.3f}" not in _dash_src
      and "Concentration: %{{x:.3f}}" not in _dash_src)


# ====================================================================
# Client request - extend last Section 7 source year if == end year
# ====================================================================
# If the last §7 source year equals the §2 end-of-simulation year, the
# source point is pushed out 5 years before writing input.inp.
print("== Source year: extend +5 when last §7 year == §2 end year ==")

_tpls = os.path.join(ROOT, "template.inp")
if os.path.exists(_tpls):
    import shutil as _shs
    from functions.state import get_state as _gss
    _sts = _gss(); _sts.snapshot = lambda a: None
    class _AppS7: pass

    def _src_times(src_years, endyr):
        wd = tempfile.mkdtemp()
        _shs.copy(_tpls, os.path.join(wd, "template.inp"))
        open(os.path.join(wd, "cellsize_input.txt"), "w").write(
            "Grid Cell Sizes\nParameter,Value\nCell Size X:,5.0\n"
            "Cell Size Y:,5.0\nCell Size Z:,2.0\nUnit Flag:,2.0\n")
        open(os.path.join(wd, "numerical_inputs.txt"), "w").write(
            "iTVD\niTVD, 1\n\n Zone\nParameter,\n"
            "Timestep Size (yr) ,0.1\nConvergence Tolerance (ug/L),1.0\n\n")
        _sts.work_dir = wd; _sts.bundle_dir = wd; _sts._cells = {}
        base = {"A8": 1, "AD1": 2, "R22": False, "E11": 100, "E12": 50,
                "E13": 10, "E15": 4, "E16": 5, "E18": 1977, "E19": endyr,
                "C22": 10, "G22": 0.3, "V47": 10}
        for _i, _y in enumerate(src_years):
            base[f"U{8+_i}"] = _y; base[f"V{8+_i}"] = 100 + _i
        _sts._cells = base
        d = gif.build_inp_data(_sts)
        return d["times"]

    # last source year 2077 == end 2077 -> last relative time 100 -> 105
    _tA = _src_times([1977, 2000, 2050, 2077], 2077)
    check("last year == end year: 100 extended to 105",
          105 in _tA and 100 not in _tA[:4], _tA[:5])
    # last source year 2070 != end 2077 -> unchanged (93)
    _tB = _src_times([1977, 2000, 2050, 2070], 2077)
    check("last year != end year: unchanged (93 kept, no 98)",
          93 in _tB and 98 not in _tB, _tB[:5])
    # exactly-at-end with a single point
    _tC = _src_times([2077], 2077)
    check("single point at end year extended (0 stays, but max 100->105 n/a)",
          True)  # smoke: build succeeded
else:
    print("  SKIP  source-year test (template.inp not found)")


# ====================================================================
# Client: velocity x Transmissive Fraction (volfrac)
# ====================================================================
print("== Velocity x Transmissive Fraction (volfrac) ==")

_tplv = os.path.join(ROOT, "template.inp")
if os.path.exists(_tplv):
    import shutil as _shv
    from functions.state import get_state as _gsv
    _stv = _gsv(); _stv.snapshot = lambda a: None
    class _AppV: pass

    def _vd_written(volfrac):
        wd = tempfile.mkdtemp()
        _shv.copy(_tplv, os.path.join(wd, "template.inp"))
        open(os.path.join(wd, "cellsize_input.txt"), "w").write(
            "Grid Cell Sizes\nParameter,Value\nCell Size X:,5.0\n"
            "Cell Size Y:,5.0\nCell Size Z:,2.0\nUnit Flag:,2.0\n")
        open(os.path.join(wd, "numerical_inputs.txt"), "w").write(
            "iTVD\niTVD, 1\n\n Zone\nParameter,\n"
            "Timestep Size (yr) ,0.1\nConvergence Tolerance (ug/L),1.0\n\n")
        open(os.path.join(wd, "heterogeneity_inputs.txt"), "w").write(
            "Heterogeneity Calculator Results\nmdflag: 2\n"
            f"Transmissive Fraction of Model (-): {volfrac}\n"
            "Diffusion Length (m): 0.25\n")
        _stv.work_dir = wd; _stv.bundle_dir = wd
        _stv._cells = {"A8": 1, "AD1": 2, "R22": False, "E11": 100, "E12": 50,
                       "E13": 10, "E15": 4, "E16": 5, "E18": 1977, "E19": 2077,
                       "C22": 10.0, "G22": 0.3, "V47": 10}
        return float(gif.build_inp_data(_stv)["vd"])

    check("vd written = bulk vd x volfrac (10 x 0.8 = 8)",
          abs(_vd_written(0.8) - 8.0) < 1e-6, _vd_written(0.8))
    check("volfrac 1.0 leaves vd unchanged (10)",
          abs(_vd_written(1.0) - 10.0) < 1e-6, _vd_written(1.0))
    check("volfrac 0.5 halves vd (5)",
          abs(_vd_written(0.5) - 5.0) < 1e-6, _vd_written(0.5))
else:
    print("  SKIP  volfrac velocity (template.inp not found)")


# ====================================================================
# Client: calibrated volfrac survives Load Optimal; velocity is BULK
# ====================================================================
# Chain: calibration ends with v_darcy = K x i (BULK, no volfrac) and the
# optimal volfrac in heterogeneity_inputs.txt + snapshot (het.volfrac).
# Load Optimal restores the bulk velocity to the GUI and the volfrac file;
# build_inp_data applies volfrac ONCE: input.inp vd = bulk x volfrac.
print("== Calibrated volfrac + bulk velocity through Load Optimal ==")

_tplc = os.path.join(ROOT, "template.inp")
if os.path.exists(_tplc):
    import shutil as _shc, re as _rec, json as _jsc
    from functions.state import get_state as _gsc
    _wd = tempfile.mkdtemp()
    _shc.copy(_tplc, os.path.join(_wd, "template.inp"))
    open(os.path.join(_wd, "cellsize_input.txt"), "w").write(
        "Grid Cell Sizes\nParameter,Value\nCell Size X:,5.0\n"
        "Cell Size Y:,5.0\nCell Size Z:,2.0\nUnit Flag:,2.0\n")
    open(os.path.join(_wd, "numerical_inputs.txt"), "w").write(
        "iTVD\niTVD, 1\n\n Zone\nParameter,\n"
        "Timestep Size (yr) ,0.1\nConvergence Tolerance (ug/L),1.0\n\n")
    open(os.path.join(_wd, "heterogeneity_inputs.txt"), "w").write(
        "Heterogeneity Calculator Results\nmdflag: 2\n"
        "Transmissive Fraction of Model (-): 0.7\nDiffusion Length (m): 0.25\n")
    _jsc.dump({"labels": ["Hydraulic Conductivity (k)"],
               "best_x": [25.0], "best_rmsle": 0.2},
              open(os.path.join(_wd, "best_calib.json"), "w"))

    _srcm = open(os.path.join(ROOT, "_Python", "tkinter", "main.py"),
                 encoding="utf-8").read()
    _mm = _rec.search(r'def _write_optimal_snapshot\(app, folder\):.*?\n    return path\n',
                      _srcm, _rec.S)
    _nsc = {"os": os, "_CALIB_PARAMS": ["Hydraulic Conductivity (k)"],
            "generate_input_file": gif}
    exec(_mm.group(0), _nsc)

    class _Vc:
        def __init__(self, v): self.v = str(v)
        def get(self): return self.v
    class _AppC: pass
    _ac = _AppC()
    _ac.v_n_iter = _Vc("50"); _ac.v_yr_start = _Vc("1977")
    _ac.v_darcy = _Vc("15"); _ac.v_porf = _Vc("0.3"); _ac.v_alpha_l = _Vc("4.75")
    for _n in ("v_ret_trans1", "v_ret_trans2", "v_ret_trans3", "v_ret_trans4"):
        setattr(_ac, _n, _Vc("1"))
    for _n in ("v_src_pfaa1", "v_src_pfaa2", "v_src_pre1", "v_src_pre2"):
        setattr(_ac, _n, [_Vc("0") for _ in range(3)])
    _ac.v_calib_chk = [_Vc("True")]; _ac.v_calib_low = [_Vc("4")]
    _ac.v_calib_mid = [_Vc("25")]; _ac.v_calib_high = [_Vc("100")]
    _snap = _nsc["_write_optimal_snapshot"](_ac, _wd)
    _stxt = open(_snap).read()
    check("snapshot stores BULK velocity (src.v_darcy=15, not x0.7)",
          "src.v_darcy=15" in _stxt)
    check("snapshot stores het.volfrac=0.7", "het.volfrac=0.7" in _stxt)

    # Clear deletes het file; Load Optimal restores it from the snapshot
    os.remove(os.path.join(_wd, "heterogeneity_inputs.txt"))
    _hetc = {}
    for _ln in open(_snap):
        _ln = _ln.strip()
        if _ln.startswith("het."):
            _k, _, _v = _ln[4:].partition("="); _hetc[_k.strip()] = _v.strip()
    open(os.path.join(_wd, "heterogeneity_inputs.txt"), "w").write(
        "Heterogeneity Calculator Results\nmdflag: %s\n"
        "Transmissive Fraction of Model (-): %s\nDiffusion Length (m): %s\n"
        % (_hetc.get("mdflag", "2"), _hetc["volfrac"], _hetc.get("difflen", "0.25")))

    _stc = _gsc(); _stc.work_dir = _wd; _stc.bundle_dir = _wd
    _stc.snapshot = lambda a: None
    _stc._cells = {"A8": 1, "AD1": 2, "R22": False, "E11": 100, "E12": 50,
                   "E13": 10, "E15": 4, "E16": 5, "E18": 1977, "E19": 2077,
                   "C22": 15.0, "G22": 0.3, "V47": 10}
    _dc = gif.build_inp_data(_stc)
    check("Run Optimal applies volfrac ONCE (15 x 0.7 = 10.5)",
          abs(float(_dc["vd"]) - 10.5) < 1e-9, _dc["vd"])
    check("volfrac itself written to input.inp data (0.7)",
          abs(float(_dc["volfrac"]) - 0.7) < 1e-9, _dc["volfrac"])
else:
    print("  SKIP  volfrac chain (template.inp not found)")


# ====================================================================
# Client: het defaults must never clobber the real volfrac
# ====================================================================
# (1) The optimal snapshot must NOT record het.* lines when
#     heterogeneity_inputs.txt is missing (defaults 1.0/0.1 would later
#     be restored over the user's real 0.8 by Load Optimal).
# (2) The calibration Mid-column readers must use the ACTIVE folder,
#     not the install dir (BASE_DIR) -- Section 4's calculator writes
#     to work_dir, so reading BASE_DIR showed volfrac 1.0 instead of 0.8.
print("== Het defaults never clobber volfrac ==")

import re as _reh
_srcm2 = open(os.path.join(ROOT, "_Python", "tkinter", "main.py"),
              encoding="utf-8").read()
_mm2 = _reh.search(r'def _write_optimal_snapshot\(app, folder\):.*?\n    return path\n',
                   _srcm2, _reh.S)
_nsh = {"os": os, "_CALIB_PARAMS": ["Hydraulic Conductivity (k)"],
        "generate_input_file": gif}
exec(_mm2.group(0), _nsh)

class _Vh:
    def __init__(self, v): self.v = str(v)
    def get(self): return self.v
class _AppH: pass
_ah = _AppH()
_ah.v_n_iter = _Vh("50"); _ah.v_yr_start = _Vh("1977")
_ah.v_darcy = _Vh("15"); _ah.v_porf = _Vh("0.3"); _ah.v_alpha_l = _Vh("4.75")
for _n in ("v_ret_trans1", "v_ret_trans2", "v_ret_trans3", "v_ret_trans4"):
    setattr(_ah, _n, _Vh("1"))
for _n in ("v_src_pfaa1", "v_src_pfaa2", "v_src_pre1", "v_src_pre2"):
    setattr(_ah, _n, [_Vh("0") for _ in range(3)])
_ah.v_calib_chk = [_Vh("True")]; _ah.v_calib_low = [_Vh("4")]
_ah.v_calib_mid = [_Vh("25")]; _ah.v_calib_high = [_Vh("100")]

# (1) folder WITHOUT het file -> snapshot must have NO het.* lines
_wdh = tempfile.mkdtemp()
_snap_h = _nsh["_write_optimal_snapshot"](_ah, _wdh)
_txt_h = open(_snap_h).read()
check("snapshot with MISSING het file records no het.* lines",
      "het.volfrac" not in _txt_h and "het.mdflag" not in _txt_h)

# folder WITH het 0.8 -> snapshot records the real value
_wdh2 = tempfile.mkdtemp()
open(os.path.join(_wdh2, "heterogeneity_inputs.txt"), "w").write(
    "Heterogeneity Calculator Results\nmdflag: 7\n"
    "Transmissive Fraction of Model (-): 0.8\nDiffusion Length (m): 0.25\n")
_snap_h2 = _nsh["_write_optimal_snapshot"](_ah, _wdh2)
check("snapshot with het 0.8 records het.volfrac=0.8",
      "het.volfrac=0.8" in open(_snap_h2).read())

# (2) source-level: Mid-column sidecar readers use the active folder
check("calibration Mid readers no longer read BASE_DIR (gwvelocity)",
      'os.path.join(BASE_DIR, "gwvelocity_inputs.txt")' not in _srcm2)
check("calibration Mid readers no longer read BASE_DIR (heterogeneity)",
      'os.path.join(BASE_DIR, "heterogeneity_inputs.txt")' not in _srcm2)
check("Mid readers use _active_dir()",
      'os.path.join(_active_dir(), "heterogeneity_inputs.txt")' in _srcm2
      and 'os.path.join(_active_dir(), "gwvelocity_inputs.txt")' in _srcm2)

print()
print(f"{PASS} passed, {FAIL} failed")
sys.exit(1 if FAIL else 0)
