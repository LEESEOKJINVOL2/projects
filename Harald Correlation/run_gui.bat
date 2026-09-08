@echo off
REM Double-click launcher for ReportBench -- run from this same folder.
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found on this PC.
    echo Install Python 3.9+ from https://www.python.org/downloads/ ^(check "Add python.exe to PATH" during setup^), then re-run this file.
    pause
    exit /b 1
)

REM Checks an actual CAPABILITY (table-cell merge), not just "is pptx
REM importable" -- a machine with some OLDER python-pptx already installed
REM from an unrelated project passed the plain-import check fine but was
REM missing _Cell.merge(), so every run crashed later with a bare
REM AttributeError deep inside a PPT-building script instead of ever
REM reaching this check's own upgrade step.
python -c "import openpyxl, matplotlib, numpy, win32com; from pptx.table import _Cell; assert hasattr(_Cell, 'merge')" 2>nul
if errorlevel 1 (
    echo Installing/upgrading required packages ^(first run, or an outdated package was found^)...
    python -m pip install --upgrade -r requirements.txt
)

python report_bench_gui.py
if errorlevel 1 pause
