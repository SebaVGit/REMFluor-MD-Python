@echo off
REM ============================================================
REM REMFluor-MD — PyInstaller build script (v104, --onedir)
REM ============================================================
REM Run from THIS folder (project root, the one containing
REM template.inp + remfluor_v8a.exe + the _Python folder).
REM
REM Usage:
REM     conda activate remfluor
REM     pip install pyinstaller
REM     build_exe.bat
REM
REM Output:
REM     dist\REMFluor-MD\REMFluor-MD.exe   (launcher inside a folder)
REM     dist\REMFluor-MD\docs\             (copied alongside .exe)
REM     dist\REMFluor-MD\Example\          (copied alongside .exe)
REM     build\                             (intermediate, can be deleted)
REM     REMFluor-MD.spec                   (regenerated each run)
REM
REM Distribution: zip the entire dist\REMFluor-MD\ folder.
REM     Users unzip, double-click REMFluor-MD.exe.  Startup ~1-2 s.
REM ============================================================
REM
REM Optimization choices (v104):
REM   --onedir          Folder layout.  Startup ~1-2 s instead of ~10 s
REM                     (--onefile unpacks the whole bundle to a temp
REM                     folder on EVERY launch.)
REM   --noupx           UPX adds ~5-10 s to startup on first launch and
REM                     compresses DLLs that Windows can mmap fast anyway.
REM                     Not worth it for an onedir build.
REM   --optimize 2      Compile with `python -OO`: strip asserts AND
REM                     docstrings.  ~5-8% smaller .pyc files, no
REM                     behavior change (our code doesn't use docstrings
REM                     for runtime help text).
REM   docs/ + Example/  NOT bundled in the .exe.  Shipped as sibling
REM                     folders next to the .exe so:
REM                     - users can bookmark help pages
REM                     - users can clone & modify example inputs
REM                     - dev can hot-patch docs without rebuilding
REM   Figures/, template.inp        — STAY bundled.  Internal
REM                     runtime assets, never user-edited.
REM   remfluor_v8a.exe   — EXTERNAL (next to .exe).  Allows swapping
REM                     the Fortran solver without rebuilding the GUI.
REM ============================================================

setlocal

REM ---- Verify we're in the right folder ---------------------
if not exist "_Python\tkinter\main.py" (
    echo ERROR: run this script from the project root
    echo        ^(the folder containing _Python\tkinter\main.py^)
    exit /b 1
)

REM ---- Verify PyInstaller is installed ----------------------
where pyinstaller >nul 2>&1
if errorlevel 1 (
    echo.
    echo ============================================================
    echo ERROR: PyInstaller is not installed in the active environment.
    echo ============================================================
    echo.
    echo To fix:
    echo   1. Activate your conda env:
    echo        conda activate remfluor
    echo      ^(your prompt should show ^(remfluor^) at the start^)
    echo.
    echo   2. Install PyInstaller:
    echo        pip install pyinstaller
    echo.
    echo   3. Re-run this script:
    echo        build_exe.bat
    echo ============================================================
    exit /b 1
)

if not exist "remfluor_v8a.exe" (
    echo WARNING: remfluor_v8a.exe not found in this folder.
    echo          The Run Model button will not work in the .exe.
)
if not exist "template.inp" (
    echo WARNING: template.inp not found in this folder.
    echo          input.inp generation will fail in the .exe.
)
if not exist "docs\_site" (
    echo WARNING: docs\_site not found.  Help "?" chicklets won't open
    echo          anything in the shipped folder.  Rebuild docs first.
)
if not exist "Example" (
    echo WARNING: Example\ not found.  "Paste Example" will fail in
    echo          the shipped folder.
)

REM ---- Clean prior build artifacts --------------------------
if exist "build"           rmdir /s /q "build"
if exist "dist"            rmdir /s /q "dist"
if exist "REMFluor-MD.spec" del "REMFluor-MD.spec"

REM ---- PyInstaller invocation -------------------------------
REM  --onedir          folder output (fast startup, recommended)
REM  --windowed        no console window on launch
REM  --noconfirm       overwrite dist/ without prompting
REM  --clean           clear PyInstaller cache before build
REM  --noupx           skip UPX even if installed (fast startup)
REM  --optimize 2      strip asserts + docstrings from bytecode
REM  --add-data        bundled read-only assets
REM  --hidden-import   force-include modules PyInstaller misses
REM  --exclude-module  drop libs we know we don't use
echo.
echo ============================================================
echo Running PyInstaller (this takes 1-3 minutes)...
echo ============================================================
pyinstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onedir ^
  --noupx ^
  --optimize 2 ^
  --name "REMFluor-MD" ^
  --paths "_Python\tkinter" ^
  --add-data "Figures;Figures" ^
  --add-data "template.inp;." ^
  --hidden-import "openpyxl" ^
  --hidden-import "openpyxl.utils" ^
  --hidden-import "openpyxl.cell" ^
  --hidden-import "pandas" ^
  --hidden-import "numpy" ^
  --hidden-import "plotly" ^
  --hidden-import "plotly.express" ^
  --hidden-import "plotly.graph_objects" ^
  --hidden-import "dash" ^
  --hidden-import "dash.dcc" ^
  --hidden-import "dash.html" ^
  --hidden-import "psutil" ^
  --collect-submodules "plotly" ^
  --collect-submodules "dash" ^
  --exclude-module "win32com" ^
  --exclude-module "pythoncom" ^
  --exclude-module "win32api" ^
  --exclude-module "pywin32_system32" ^
  --exclude-module "xlwings" ^
  --exclude-module "matplotlib" ^
  --exclude-module "matplotlib.tests" ^
  --exclude-module "scipy" ^
  --exclude-module "scipy.tests" ^
  --exclude-module "sklearn" ^
  --exclude-module "PyQt5" ^
  --exclude-module "PyQt6" ^
  --exclude-module "PySide2" ^
  --exclude-module "PySide6" ^
  --exclude-module "IPython" ^
  --exclude-module "ipykernel" ^
  --exclude-module "jupyter" ^
  --exclude-module "notebook" ^
  --exclude-module "pytest" ^
  --exclude-module "numpy.tests" ^
  --exclude-module "numpy.distutils" ^
  --exclude-module "numpy.f2py" ^
  --exclude-module "pandas.tests" ^
  --exclude-module "pandas.io.tests" ^
  --exclude-module "openpyxl.tests" ^
  --exclude-module "test" ^
  --exclude-module "tests" ^
  --exclude-module "tornado" ^
  --exclude-module "zmq" ^
  --exclude-module "lxml" ^
  --exclude-module "h5py" ^
  --exclude-module "tables" ^
  _Python\tkinter\main.py

if errorlevel 1 (
    echo.
    echo BUILD FAILED — see PyInstaller output above.
    exit /b 1
)

REM ---- Copy external assets next to the .exe ----------------
REM  docs/ and Example/ live alongside the .exe (NOT bundled).
REM  See header comment for rationale.
echo.
echo ============================================================
echo Copying docs\ and Example\ next to the .exe...
echo ============================================================
if exist "docs" (
    xcopy "docs" "dist\REMFluor-MD\docs\" /E /I /Q /Y >nul
    if errorlevel 1 (
        echo WARNING: docs copy failed.
    ) else (
        echo   docs\     copied.
    )
)
if exist "Example" (
    xcopy "Example" "dist\REMFluor-MD\Example\" /E /I /Q /Y >nul
    if errorlevel 1 (
        echo WARNING: Example copy failed.
    ) else (
        echo   Example\  copied.
    )
)
if exist "remfluor_v8a.exe" (
    copy /Y "remfluor_v8a.exe" "dist\REMFluor-MD\" >nul
    if errorlevel 1 (
        echo WARNING: remfluor_v8a.exe copy failed.
    ) else (
        echo   remfluor_v8a.exe  copied.
    )
)

REM ---- Done -------------------------------------------------
echo.
echo ============================================================
echo BUILD SUCCEEDED
echo ============================================================
echo Output folder: dist\REMFluor-MD\
echo Launcher:      dist\REMFluor-MD\REMFluor-MD.exe
echo.
echo Folder contents:
dir /b "dist\REMFluor-MD" | findstr /v "_internal"
echo   _internal\   ^(PyInstaller runtime - DO NOT delete or move^)
echo.
echo Total size:
for /f "tokens=3" %%A in ('dir "dist\REMFluor-MD" /s /-c ^| findstr /C:"File(s)"') do echo   %%A bytes
echo ============================================================
echo.
echo TO SHIP:
echo   1. Zip the entire dist\REMFluor-MD\ folder.
echo   2. Recipient unzips wherever they want.
echo   3. Double-click REMFluor-MD.exe inside the unzipped folder.
echo.
echo TO TEST:
echo   1. Run dist\REMFluor-MD\REMFluor-MD.exe locally.
echo   2. Then test on a clean machine ^(no Python, no conda installed^).
echo   3. Click a "?" chicklet to verify external docs\ resolves.
echo   4. Click Paste Example to verify external Example\ resolves.
echo   5. Click Run Model to verify external remfluor_v8a.exe runs.

endlocal
