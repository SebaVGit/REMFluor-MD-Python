"""
REMFluor-MD v107 regression tests — run BEFORE every build:

    cd <project root>
    python tests\test_regression_v107.py

Headless (mocks tkinter), no GUI needed.  Every test captures a client
bug fixed in v107; if any FAIL, a fix has regressed — do not ship.
"""
import os
import sys
import types
import tempfile

sys.dont_write_bytecode = True   # never poison functions/__pycache__

# ── locate the package ─────────────────────────────────────────────
HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
sys.path.insert(0, os.path.join(ROOT, "_Python", "tkinter"))

# ── mock tkinter so functions/ imports headless ─────────────────────
tk = types.ModuleType("tkinter")
tk.messagebox = types.ModuleType("tkinter.messagebox")
tk.messagebox.showerror = lambda *a, **k: None
tk.messagebox.showinfo = lambda *a, **k: None
tk.filedialog = types.ModuleType("tkinter.filedialog")
sys.modules["tkinter"] = tk
sys.modules["tkinter.messagebox"] = tk.messagebox
sys.modules["tkinter.filedialog"] = tk.filedialog

from functions.inp_parser import parse_input_inp, parse_additional_info
from functions.inp_to_state import write_inp_to_state
from functions import generate_input_file as gif
from functions import sidecars


class FakeState:
    def __init__(self, wd=""):
        self.d = {}
        self.work_dir = wd
        self.bundle_dir = ""
        self.base_dir = ""

    def get(self, a, d=None):
        return self.d.get(a, d)

    def set(self, a, v):
        self.d[a] = v


class V:
    def __init__(self, v=""):
        self.v = v

    def get(self):
        return self.v

    def set(self, x):
        self.v = x


class App:
    pass


PASS = FAIL = 0


def check(name, cond, detail=""):
    global PASS, FAIL
    if cond:
        PASS += 1
        print(f"  PASS  {name}")
    else:
        FAIL += 1
        print(f"  FAIL  {name}   {detail}")


def _cellsize(tmp, dx=5.0, dy=5.0, dz=2.0, unit=2.0):
    with open(os.path.join(tmp, "cellsize_input.txt"), "w") as f:
        f.write("Grid Cell Sizes\nParameter,Value\n"
                f"Cell Size X:,{dx}\nCell Size Y:,{dy}\n"
                f"Cell Size Z:,{dz}\nUnit Flag:,{unit}\n")


# ════════════════════════════════════════════════════════════════════
print("== Load Data fixes (detailed example, UI mismatched on Simple) ==")
EX = os.path.join(ROOT, "Example", "3_Detailed_2")
data = parse_input_inp(os.path.join(EX, "input.inp"))
add = parse_additional_info(os.path.join(EX, "store_info_additional_input.txt"))
st = FakeState(EX)
st.set("A8", 1)          # UI on Simple pre-load = historical failure mode
write_inp_to_state(st, data, add, unit_flag=2, volfrac=0.8)

check("velocity un-scaled by transmissive fraction (0.011992/0.8)",
      abs(st.get("C22") - 0.014990) < 1e-6, f"got {st.get('C22')}")
check("PSB year at AB28, as int", st.get("AB28") == 2025,
      f"got {st.get('AB28')!r}")
check("PSB x-cells at AC82", st.get("AC82") == 4, f"got {st.get('AC82')}")
check("Kf unit consistent with derived values",
      st.get("U24") == "(ug/kg)(ug/L)^(-a)", f"got {st.get('U24')}")

# PSB off -> on must reproduce identical solver params
for a, b in (("V24", "V26"), ("X24", "X26"), ("Z24", "Z26"), ("AB24", "AB26")):
    st.set(b, st.get(a))
d1 = gif.build_inp_data(st)
st.set("R22", False)
gif.build_inp_data(st)
st.set("R22", True)
d3 = gif.build_inp_data(st)
check("PSB off->on round trip identical",
      all(d1[k] == d3[k] for k in ("iwall", "nxpsb", "tplume1", "fcackf2")))
check("fcackf2 matches original input.inp",
      abs(float(d3["fcackf2"]) - 2947.0824) < 1e-3, f"got {d3['fcackf2']}")

# ════════════════════════════════════════════════════════════════════
print("== PSB Freundlich Kf unit round trip ==")
raw_mg, a_exp, load = 1227.951, 0.33, 0.0024
fc = raw_mg * (1e3 ** (1 - a_exp)) * load       # what save-side writes
dpsb = dict(data, fcackf2=fc, fra2=a_exp, iwall=2, ipre=0, ncomp=1)
apsb = {"psb_loading": load, "start_year": 1977,
        "psb_kf_unit": "(mg/kg)(mg/L)^(-a)"}
s2 = FakeState()
s2.set("A8", 2)
write_inp_to_state(s2, dpsb, apsb, 2, 1.0)
check("mg unit respected on load", s2.get("U24") == "(mg/kg)(mg/L)^(-a)")
check("mg raw value recovered exactly", s2.get("V24") == "1,227.951",
      f"got {s2.get('V24')}")
apsb["psb_kf_unit"] = "(ug/kg)(ug/L)^(-a)"
dpsb["fcackf2"] = 5000.0 * load
s3 = FakeState()
s3.set("A8", 2)
write_inp_to_state(s3, dpsb, apsb, 2, 1.0)
check("ug unit: raw == converted (5,000)", s3.get("V24") == "5,000",
      f"got {s3.get('V24')}")

# ════════════════════════════════════════════════════════════════════
print("== Grid: dy/ny span, well-z clamp, x-direction ==")
tmp = tempfile.mkdtemp()
_cellsize(tmp, dz=0.67)
sg = FakeState(tmp)
sg.set("A8", 1); sg.set("AD1", 2); sg.set("R22", False)
sg.set("E11", 150); sg.set("E12", 6.096); sg.set("E13", 6.7056)
sg.set("E15", 4); sg.set("E16", 3.35); sg.set("E18", 1977); sg.set("E19", 2077)
sg.set("C22", 10); sg.set("G22", 0.3)
for i, x in enumerate((125, 128, 131, 134, 137, 140, 143)):
    sg.set(f"AF{34+i}", x)
dg = gif.build_inp_data(sg)
check("dy*ny spans Y (6.1 x 1 for Y=6.096)",
      abs(float(dg["dy"]) * dg["ny"] - 6.1) < 1e-6,
      f"got dy={dg['dy']} ny={dg['ny']}")
check("well z clamped to grid (6.7, not 6.71)",
      dg["wells"][0].endswith("6.7, 3.35"), f"got {dg['wells'][0]}")

tmp2 = tempfile.mkdtemp()
_cellsize(tmp2, dx=1.0, dy=5.0, dz=2.0)
sx = FakeState(tmp2)
sx.set("A8", 2); sx.set("AD1", 2); sx.set("R22", True)
sx.set("E11", 132); sx.set("E12", 50); sx.set("E13", 10)
sx.set("E15", 4); sx.set("E16", 5); sx.set("E18", 1977); sx.set("E19", 2077)
sx.set("C22", 10); sx.set("G22", 0.3)
sx.set("X74", 108); sx.set("Y82", 12); sx.set("AB28", 2025)
sx.set("AA82", 0.24); sx.set("V23", 0.33); sx.set("V24", 100)
sx.set("AC82", 12)
dx1 = gif.build_inp_data(sx)
check("x-grid 1m: 132 = 108 + 12 + 12",
      (dx1["nx1"], dx1["nxpsb"], dx1["nx2"]) == (108, 12, 12),
      f"got {dx1['nx1']},{dx1['nxpsb']},{dx1['nx2']}")
sx.set("AC82", 48)
check("x-grid 0.25m PSB: nxpsb=48",
      gif.build_inp_data(sx)["nxpsb"] == 48)
sx.set("AC82", None)
check("blank PSB cells defaults to width/dx (12), not 0",
      gif.build_inp_data(sx)["nxpsb"] == 12)

# ════════════════════════════════════════════════════════════════════
print("== Calibration overrides must not leak into manual runs ==")
sx.set("_calib_volfrac", 0.312)
v_with = gif.build_inp_data(sx)["volfrac"]
sx.set("_calib_volfrac", None)
v_clear = gif.build_inp_data(sx)["volfrac"]
check("override honored while set / ignored when cleared",
      v_with == 0.312 and v_clear != 0.312,
      f"with={v_with} cleared={v_clear}")

# ════════════════════════════════════════════════════════════════════
print("== Sidecars: dispersivity + mol diff ==")
tmp3 = tempfile.mkdtemp()
ap = App()
ap.v_het = V("Enter Your Own Value Below"); ap.v_units = V("meters")
ap.v_alpha_l = V("4.75"); ap.v_alpha_t = V("0.09"); ap.v_alpha_v = V("0.0071")
ap.v_mol_diff = V("4.20E-10")
sidecars.write_dispersivity(ap, tmp3)
b = App(); b.v_units = V("meters")
for n in ("v_het", "v_alpha_l", "v_alpha_t", "v_alpha_v", "v_mol_diff", "v_pfaa1"):
    setattr(b, n, V())
sidecars.read_dispersivity(b, tmp3)
check("custom dispersivity exact round trip",
      (b.v_het.get(), b.v_alpha_l.get(), b.v_mol_diff.get()) ==
      ("Enter Your Own Value Below", "4.75", "4.20E-10"))
c = App(); c.v_units = V("feet")
for n in ("v_het", "v_alpha_l", "v_alpha_t", "v_alpha_v", "v_mol_diff", "v_pfaa1"):
    setattr(c, n, V())
sidecars.read_dispersivity(c, tmp3)
check("cross-unit m->ft conversion (4.75 m = 15.584 ft)",
      abs(float(c.v_alpha_l.get()) - 15.584) < 1e-3,
      f"got {c.v_alpha_l.get()}")
m = App(); m.v_mol_diff = V(); m.v_pfaa1 = V("PFOS")
stm = FakeState(); stm.set("E44", 3.5e-10)
r = sidecars.mol_diff_from_state_fallback(m, stm, tempfile.mkdtemp())
check("mol-diff E44 fallback when sidecar absent",
      r and m.v_mol_diff.get() == "3.50E-10" and m._mol_diff_user_edited,
      f"got {m.v_mol_diff.get()}")

# ════════════════════════════════════════════════════════════════════
print()
print(f"{PASS} passed, {FAIL} failed")
sys.exit(1 if FAIL else 0)
