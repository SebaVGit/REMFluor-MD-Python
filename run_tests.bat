@echo off
REM ============================================================
REM REMFluor-MD smoke test - run before every commit / build.
REM v107: 18 checks (load round-trips, PSB, grid, units, sidecars).
REM v108: 16 checks (Ron review item 1 fcac percent, item 2 dt).
REM All must PASS before shipping.
REM ============================================================
cd /d "%~dp0"

REM stale bytecode can shadow edited source - always clear it
if exist "_Python\tkinter\functions\__pycache__" rd /s /q "_Python\tkinter\functions\__pycache__"
if exist "_Python\tkinter\__pycache__" rd /s /q "_Python\tkinter\__pycache__"

set _FAILED=0

python tests\test_regression_v107.py
if errorlevel 1 set _FAILED=1

python tests\test_regression_v108.py
if errorlevel 1 set _FAILED=1

if %_FAILED%==1 (
    echo.
    echo *** SMOKE TEST FAILED - do not build/commit. ***
) else (
    echo.
    echo All checks passed - safe to commit and build.
)
pause
