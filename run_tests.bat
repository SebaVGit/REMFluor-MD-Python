@echo off
REM ============================================================
REM REMFluor-MD smoke test — run before every commit / build.
REM 18 checks covering the v107 fixes (load round-trips, PSB,
REM grid, units, sidecars).  All must PASS before shipping.
REM ============================================================
cd /d "%~dp0"

REM stale bytecode can shadow edited source — always clear it
if exist "_Python\tkinter\functions\__pycache__" rd /s /q "_Python\tkinter\functions\__pycache__"
if exist "_Python\tkinter\__pycache__" rd /s /q "_Python\tkinter\__pycache__"

python tests\test_regression_v107.py
if errorlevel 1 (
    echo.
    echo *** SMOKE TEST FAILED — do not build/commit. ***
) else (
    echo.
    echo All checks passed — safe to commit and build.
)
pause
