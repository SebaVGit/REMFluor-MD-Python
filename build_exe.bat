@echo off
REM ============================================================
REM REMFluor-MD — PyInstaller build script
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
REM     dist\REMFluor-MD.exe       (single-file .exe)
REM     build\                     (intermediate, can be deleted)
REM     REMFluor-MD.spec           (regenerated each run)
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
    echo.
    echo Notes:
    echo   - UPX compression is OPTIONAL.  Without it the .exe will
    echo     be ~250 MB instead of ~140 MB but works identically.
    echo   - If you also want the smaller build, download UPX from
    echo     https://upx.github.io/ and unzip somewhere ^(e.g.
    echo     C:\Tools\upx-4.2.4-win64\^), then either add that folder
    echo     to PATH or edit the top of this script to set UPX_DIR.
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

REM ---- Clean prior build artifacts --------------------------
if exist "build"          rmdir /s /q "build"
if exist "dist"           rmdir /s /q "dist"
if exist "REMFluor-MD.spec" del "REMFluor-MD.spec"

REM ---- Optional UPX compression -----------------------------
REM  UPX shrinks the bundled DLLs/PYDs ~40-50%% (250 MB -> ~140 MB).
REM  Trade-off: ~5-10 s extra startup time on first launch.
REM
REM  Install: download from https://upx.github.io/ and unzip.
REM  Then either:
REM    - put upx.exe somewhere on PATH, or
REM    - set UPX_DIR below to the folder containing upx.exe
REM
REM  Example:
REM    set UPX_DIR=C:\Tools\upx-4.2.4-win64
REM
REM  If UPX isn't installed the build still works, just bigger.

set UPX_FLAGS=
if not "%UPX_DIR%"=="" (
    if exist "%UPX_DIR%\upx.exe" (
        set UPX_FLAGS=--upx-dir "%UPX_DIR%"
        echo Using UPX at %UPX_DIR%
    )
) else (
    where upx >nul 2>&1
    if not errorlevel 1 (
        set UPX_FLAGS=
        echo Using UPX from PATH
    ) else (
        echo NOTE: UPX not found.  Build will skip compression.
        echo       For ~140 MB instead of ~250 MB, install UPX from:
        echo       https://upx.github.io/  and re-run.
    )
)

REM ---- PyInstaller invocation -------------------------------
REM  --onefile      single .exe (slower startup, simplest)
REM  --windowed     no console window on launch
REM  --noconfirm    overwrite dist/ without prompting
REM  --clean        clear PyInstaller cache before build
REM  --add-data     "src;dst" (Windows uses ;, Linux/Mac uses :)
REM  --hidden-import   force-include modules PyInstaller misses
REM  --exclude-module  drop libs we know we don't use
REM  --upx-exclude     skip these DLLs (UPX breaks some Win runtime DLLs)
REM
REM  All paths are relative to this script's folder.

REM NOTE: --strip is a Unix flag and Windows doesn't ship strip.exe.
REM       Including it produces ~150 harmless "WARNING: Failed to run
REM       strip" messages but doesn't break the build.  Removed.
pyinstaller ^
  --noconfirm ^
  --clean ^
  --windowed ^
  --onefile ^
  %UPX_FLAGS% ^
  --upx-exclude "vcruntime140.dll" ^
  --upx-exclude "python3*.dll" ^
  --upx-exclude "ucrtbase.dll" ^
  --upx-exclude "qwindows.dll" ^
  --upx-exclude "qwindowsvistastyle.dll" ^
  --name "REMFluor-MD" ^
  --paths "_Python\tkinter" ^
  --add-data "Figures;Figures" ^
  --add-data "template.inp;." ^
  --add-data "remfluor_v8a.exe;." ^
  --add-data "Example;Example" ^
  --add-data "docs;docs" ^
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

echo.
echo ============================================================
echo BUILD SUCCEEDED
echo Output: dist\REMFluor-MD.exe
echo Size:
dir /b /s dist\REMFluor-MD.exe
for %%I in ("dist\REMFluor-MD.exe") do echo   %%~zI bytes
echo ============================================================
echo.
echo NEXT STEPS:
echo   1. Test on this machine:   dist\REMFluor-MD.exe
echo   2. Test on a clean machine ^(no Python installed^) before
echo      shipping. The .exe is fully self-contained for the
echo      tkinter UI; the Plotly-Dash dashboard subprocess
echo      requires Python on the target machine ^(see BUILD.md^).

endlocal
