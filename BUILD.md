# REMFluor-MD — building a standalone `.exe`

## Quick start

```cmd
conda activate remfluor
pip install pyinstaller
build_exe.bat
```

Output: **`dist\REMFluor-MD\`** — a folder containing `REMFluor-MD.exe` plus its runtime dependencies and the sibling `docs/` + `Example/` folders.

To distribute: **zip the entire `dist\REMFluor-MD\` folder** and send/share that zip. Recipients unzip wherever they like and double-click `REMFluor-MD.exe`. Startup is ~1–2 s.

## Layout of the shipped folder

```
REMFluor-MD\
├── REMFluor-MD.exe        ← launcher
├── _internal\             ← PyInstaller runtime (Python DLLs, libs).  DO NOT delete or move.
│   ├── Figures\           ← bundled icons / illustrations
│   └── template.inp       ← bundled input.inp generator template
├── remfluor_v9a.exe       ← external Fortran solver (swappable)
├── docs\                  ← external help pages (HTML chicklets)
│   └── _site\…
└── Example\               ← external paste-example inputs
    ├── 1_Simple\…
    └── 2_Detailed\…
```

Three things to know:

1. `REMFluor-MD.exe` must stay next to `_internal\`. They ship together.
2. `docs/` and `Example/` live alongside the .exe (NOT inside `_internal\`). Users can:
   - Bookmark chicklet pages in their browser (paths are stable).
   - Copy / modify `Example/` site setups to use as templates.
   - You can hot-patch a docs typo by editing the file — no rebuild needed.
3. `Figures/` and `template.inp` are bundled inside `_internal\` because users should never edit them, and they're version-coupled to the GUI.
4. `remfluor_v9a.exe` is external (next to the launcher) so you can drop in a new build of the Fortran solver without rebuilding the GUI.

## What gets bundled vs external

| Path | Where | Why |
|------|-------|-----|
| `Figures/` | bundled (`_internal\`) | Runtime icons. Never user-edited. |
| `template.inp` | bundled (`_internal\`) | Internal generator. Version-coupled. |
| `remfluor_v9a.exe` | **external** (next to .exe) | Solver. Swap-in-place without rebuilding the GUI. |
| `docs/` | **external** (next to .exe) | Browsable HTML. Bookmarkable. Editable docs without rebuild. |
| `Example/` | **external** (next to .exe) | Users clone & modify these as templates for their own sites. |

`main.py` is already set up for this split:
- `_html()` resolves docs at `BASE_DIR/docs/_site/...` (next to .exe, never the bundle).
- `_bundle_path()` and `run_model._resolve_asset()` both check `BASE_DIR` first and fall back to `_MEIPASS` — so any asset works external OR bundled. The build script controls which one wins.
- `restore_from_example.run()` checks `BASE_DIR/Example` first, then `BUNDLE_DIR/Example`.

## What's excluded from the build

Removed because the project no longer uses them (or never did):

- `win32com`, `pythoncom`, `win32api`, `pywin32_system32` — Excel COM removed in v82–v86
- `xlwings` — replaced with `openpyxl` (deferred import)
- `matplotlib`, `scipy`, `sklearn` — not used
- `PyQt5`, `PyQt6`, `PySide2`, `PySide6` — UI is pure tkinter
- `IPython`, `jupyter`, `notebook`, `pytest`, `tornado`, `zmq`, `lxml`, `h5py`, `tables` — dev/server-only
- `numpy.tests`, `pandas.tests`, `openpyxl.tests`, `numpy.distutils`, `numpy.f2py` — test data and distutils internals (~20 MB combined)

If PyInstaller still pulls in something heavy, add it to the `--exclude-module` list in `build_exe.bat`.

## Required Python deps (must be in the conda env when building)

```
pip install pyinstaller pandas numpy plotly dash openpyxl psutil
```

`tkinter` is part of Python's stdlib — no install needed.

## Optimization choices (v104)

| Flag | Effect |
|------|--------|
| `--onedir` | Folder layout. Startup ~1–2 s vs ~10 s for `--onefile`. `--onefile` unpacks the entire bundle to a temp folder on every launch — wasted disk + slow start. |
| `--noupx` | Skip UPX compression. UPX adds 5–10 s to first-launch startup and compresses DLLs that Windows can mmap fast already. Not worth it for `--onedir`. |
| `--optimize 2` | Compile bytecode with `python -OO`: strips asserts + docstrings. ~5–8% smaller `.pyc`. No behavior change. |
| `--windowed` | No console window on launch. |
| `--noconfirm`, `--clean` | Overwrite `dist/` without prompting. Clear PyInstaller cache before build. |

## The dashboard subprocess

The Plotly-Dash dashboard runs as a **separate subprocess**. In a frozen build, `sys.executable` is the `.exe` itself, so the launcher uses the multi-mode dispatcher pattern (already wired in `main.py` lines ~25–60 — looks for `--mode=dashboard` argv):

```python
import sys
if len(sys.argv) >= 2 and sys.argv[1] == "--mode=dashboard":
    from functions import generate_dashboard
    sys.argv = [sys.argv[0]] + sys.argv[2:]
    generate_dashboard.main()
    sys.exit(0)
```

And `run_model.py` checks `sys.frozen` to decide between `[exe, --mode=dashboard, …]` and `[python, -u, script, …]`. No Python install on the target machine is required.

## Size expectations

Folder total **~250–300 MB without UPX**. The biggest contributors:

| Library | Frozen | Why we need it |
|---|---|---|
| `pandas` | ~50 MB | Dashboard time-series + .out parsing |
| `numpy` | ~35 MB | pandas dependency, also used by cali_1.py |
| `plotly` (templates + colorscales + JS bundle) | ~50 MB | Dashboard plots |
| `dash` (+ Flask + Werkzeug) | ~25 MB | Dashboard server |
| `tkinter` (Tcl/Tk runtime) | ~15 MB | The whole UI |
| Python interpreter | ~15 MB | Required |
| `openpyxl`, `psutil`, stdlib | ~10 MB | §10 .xlsx import + port cleanup |
| Bundled assets (`Figures/`, `template.inp`) | ~3 MB | Required at runtime |
| External `remfluor_v9a.exe` | ~1 MB | Fortran solver (sibling of .exe) |
| External `docs/`, `Example/` (not in .exe, but in the shipped folder) | ~5–20 MB | Side-loaded |

→ raw ~250–300 MB pre-compression is normal.

## Slimming further

If you can drop the dashboard entirely (e.g., ship a stripped build for users who never click Run Model), add these excludes:

```
--exclude-module "pandas"
--exclude-module "numpy"
--exclude-module "plotly"
--exclude-module "dash"
```

…which gets you down to ~30 MB. But Run Model will fail (the import pre-flight check will catch it). Not recommended for production distributions.

## Sanity-check on a clean machine

1. Zip `dist\REMFluor-MD\` and copy to a Windows machine WITHOUT Python or any of these libraries installed.
2. Unzip wherever.
3. Double-click `REMFluor-MD.exe`.
4. Verify:
   - App opens within ~2 s.
   - Click any "?" chicklet → opens HTML in the default browser at a path under the unzipped folder's `docs/_site/...`.
   - Paste Example → fills the form (reads from `Example/` next to the .exe).
   - Run Model → produces `input.inp` and runs the external `remfluor_v9a.exe`.
   - Run Model → dashboard launches in the browser (multi-mode dispatcher).

## Troubleshooting

**"Help links open nothing / wrong path"**
The shipped `docs/` folder is missing or got separated from the .exe. The chicklet flow resolves `BASE_DIR/docs/_site/...`, where `BASE_DIR` is the folder containing the .exe. Re-zip with `docs/` included.

**"Paste Example fails with 'Example folder not found'"**
Same issue with `Example/`. The code falls back to the bundled copy at `_MEIPASS/Example` if external is missing, but with `--onedir` builds nothing is bundled in `_MEIPASS` — it's all in `_internal/`. Keep `Example/` next to the .exe.

**".exe folder is 500 MB+"**
PyInstaller pulled in a heavy library you don't actually use. Run with `--log-level INFO` and look for surprising entries; add them to `--exclude-module`.

**"Run Model says input.inp not found" or "Could not start remfluor_v9a.exe"**
`template.inp` (bundled) and/or `remfluor_v9a.exe` (external) weren't shipped. Verify the `--add-data "template.inp;."` line in `build_exe.bat` and that `remfluor_v9a.exe` exists in the project root before building (it gets `copy`'d next to the launcher post-build).

**"Dashboard subprocess silently fails"**
The `--mode=dashboard` dispatcher at the top of `main.py` is missing or the `run_model.py` `sys.frozen` branch isn't there. Both should already be in v99+ codebases — see the dashboard subprocess section above.
