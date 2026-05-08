# REMFluor-MD — building a standalone `.exe`

## Quick start

```cmd
conda activate remfluor
pip install pyinstaller
build_exe.bat
```

Output: **`dist\REMFluor-MD.exe`** (single file, ~100–200 MB).

## What gets bundled

PyInstaller follows the import graph automatically, so all `_Python/tkinter/**.py` files are pulled in. Non-Python data files have to be specified explicitly via `--add-data`. The script bundles:

| Path | Purpose |
|------|---------|
| `Figures/` | All section icons / illustrations loaded by `_load_figure()` |
| `template.inp` | Used by `generate_input_file` to build `input.inp` |
| `remfluor_v8a.exe` | Fortran solver, invoked as subprocess by Run Model |
| `Example/` | Paste-Example data (`1_Simple/input.inp`, etc.) |
| `docs/` | HTML help pages opened by the help links |

## What's excluded

Removed because the project no longer uses them (or never did):

- `win32com`, `pythoncom`, `win32api`, `pywin32_system32` — Excel COM was removed in v82–v86 (dashboard now reads `dashboard_state.json`)
- `xlwings` — replaced with `openpyxl` (deferred import)
- `matplotlib`, `scipy`, `sklearn` — not used
- `PyQt5`, `PyQt6`, `PySide2`, `PySide6` — UI is pure tkinter
- `IPython`, `jupyter`, `notebook`, `pytest`, `tornado`, `zmq`, `lxml`, `h5py`, `tables`

If PyInstaller still pulls something heavy in, add it to the `--exclude-module` list in `build_exe.bat`.

## Required Python deps (must be in the conda env when building)

```
pip install pyinstaller pandas numpy plotly dash openpyxl psutil
```

`tkinter` is part of Python's stdlib — no install needed.

## Single file vs folder layout

The script uses `--onefile` for simplicity. Trade-offs:

| Mode | Disk size | Startup | Distribution |
|------|-----------|---------|--------------|
| `--onefile`  (current) | smaller (compressed) | slow (unpacks to temp dir each run) | one file, easy |
| `--onedir` | larger | fast | folder of files |

To switch to `--onedir`, replace `--onefile` with `--onedir` and zip the resulting `dist\REMFluor-MD\` folder.

## The dashboard subprocess — important caveat

The Plotly-Dash dashboard runs as a **separate Python subprocess**:

```python
subprocess.Popen([sys.executable, "-u", script, workbook_path, sheet_name])
```

In a frozen `.exe`:
- `sys.executable` is the `.exe` itself, not `python.exe`
- Re-launching the `.exe` with extra argv re-runs `main()`, which is wrong

Two options for the dashboard to work in a frozen build:

### Option A — multi-mode dispatcher (recommended)

Add this at the very top of `_Python/tkinter/main.py` (before any other imports that might launch tkinter):

```python
import sys
if len(sys.argv) >= 2 and sys.argv[1] == "--mode=dashboard":
    # Re-entry from run_model.py launching itself as dashboard
    from functions import generate_dashboard
    sys.argv = [sys.argv[0]] + sys.argv[2:]   # strip the flag
    generate_dashboard.main()
    sys.exit(0)
```

Then update `_Python/tkinter/functions/run_model.py` `_launch_dashboard_async`:

```python
# Replace:
cmd = [py, "-u", script, workbook_path, sheet_name]
# With:
if getattr(sys, "frozen", False):
    cmd = [py, "--mode=dashboard", workbook_path, sheet_name]
else:
    cmd = [py, "-u", script, workbook_path, sheet_name]
```

This way the frozen `.exe` re-launches itself with `--mode=dashboard` and dispatches into the dashboard code. No external Python install needed on the target machine.

### Option B — leave the subprocess pattern

Ship the `.exe` AND require the user to have Python + the dashboard deps installed. The `subprocess.Popen([sys.executable, ...])` path will only work if `sys.executable` happens to point at a real Python (which it doesn't in a frozen build).

I recommend **Option A** for any user-facing distribution. The build script as it stands creates a working tkinter UI; the dashboard subprocess just needs the dispatcher patch above to work end-to-end.

## Size expectations

Typical `.exe` size for this build is **220–260 MB without UPX**, **120–160 MB with UPX**. The biggest contributors are:

| Library | Frozen | Why we need it |
|---|---|---|
| `pandas` | ~50 MB | Dashboard time-series + .out parsing |
| `numpy` | ~35 MB | pandas dependency, also used by cali_1.py |
| `plotly` (templates + colorscales + JS bundle) | ~50 MB | Dashboard plots |
| `dash` (+ Flask + Werkzeug) | ~25 MB | Dashboard server |
| `tkinter` (Tcl/Tk runtime) | ~15 MB | The whole UI |
| Python interpreter | ~15 MB | Required |
| `openpyxl`, `psutil`, stdlib | ~10 MB | §10 .xlsx import + port cleanup |
| Bundled data (`Figures/`, `Example/`, `docs/`, `template.inp`, `remfluor_v8a.exe`) | ~10–20 MB | Required at runtime |

→ raw 215–250 MB pre-compression is normal.

## UPX — cuts the size in half

UPX is a binary packer that compresses the bundled DLLs/PYDs at build time and decompresses on launch. **Drops a 250 MB build to ~140 MB** at the cost of ~5–10 s extra startup time.

### Install UPX once

1. Download from <https://upx.github.io/> (`upx-X.X.X-win64.zip`)
2. Unzip somewhere — e.g., `C:\Tools\upx-4.2.4-win64\`
3. Either add that folder to your PATH, **or** edit `build_exe.bat` and uncomment the `set UPX_DIR=...` line

### Then re-run

```cmd
build_exe.bat
```

The script auto-detects UPX and applies it if present. Without UPX, the build still works — just bigger.

## Slimming without UPX

If you can drop the dashboard entirely (e.g., ship a stripped build for users who never click Run Model), add these excludes:

```
--exclude-module "pandas"
--exclude-module "numpy"
--exclude-module "plotly"
--exclude-module "dash"
```

…which gets you down to ~30 MB. But Run Model will fail (the import pre-flight check will catch it). Not recommended for production distributions.

## What's already excluded

These are dropped from the bundle even though some might be in your conda env:

- `win32com`, `pythoncom`, `win32api`, `xlwings` — removed in v82–v86 (dashboard now reads `dashboard_state.json`)
- `matplotlib`, `scipy`, `sklearn` — never used
- `PyQt5/6`, `PySide2/6` — pure tkinter UI
- `IPython`, `jupyter`, `notebook`, `pytest` — dev/server-only
- `numpy.tests`, `pandas.tests`, `openpyxl.tests`, `numpy.distutils`, `numpy.f2py` — test data and distutils internals (~20 MB combined)
- `tornado`, `zmq`, `lxml`, `h5py`, `tables` — heavy libs sometimes pulled by transitive deps

## Sanity-check on a clean machine

1. Copy `dist\REMFluor-MD.exe` to a Windows machine WITHOUT Python or any of these libraries installed.
2. Double-click it.
3. Verify: app opens, Paste Example fills the form, Run Model produces `input.inp` + runs the Fortran solver.
4. Dashboard step requires the Option A patch above — without it, the dashboard subprocess will silently fail (you'll see "Dashboard NOT launched" in the runtime-clock window).

## Troubleshooting

**"Dashboard Dependencies Missing: pandas / numpy / plotly / dash"**
The pre-flight check in `run_model.py` runs `python -c "import pandas, ..."`. In a frozen build, `python -c` doesn't work — that check needs the same `sys.frozen` shim as the dashboard launch. Easiest fix: skip the pre-flight check entirely when frozen.

**".exe is 500 MB+"**
PyInstaller pulled in a heavy library you don't actually use. Run with `--log-level INFO` and look for surprising entries; add them to `--exclude-module`.

**"Run Model says input.inp not found"**
`template.inp` and/or `remfluor_v8a.exe` weren't bundled. Verify the `--add-data` lines in `build_exe.bat` and that the source files exist in the project root.

**"Help links open a 'file not found' dialog"**
`docs/_site/` is missing from the bundle. Make sure `docs/` exists at the project root with the rendered HTML.
