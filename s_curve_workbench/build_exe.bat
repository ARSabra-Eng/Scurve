@echo off
REM Build script for Parametric S-Curve Workbench Windows executable
REM Requires: pip install pyinstaller

echo Building Parametric S-Curve Workbench v1.0...

REM Clean previous builds
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist

REM Run PyInstaller
pyinstaller --onefile ^
    --name "s_curve_workbench" ^
    --windowed ^
    --icon=NONE ^
    --add-data "requirements.txt;." ^
    --hidden-import PySide6 ^
    --hidden-import matplotlib ^
    --hidden-import openpyxl ^
    --hidden-import numpy ^
    s_curve_workbench.py

echo.
echo Build complete! Executable located in: dist\s_curve_workbench.exe
echo.
echo To distribute:
echo   1. Copy dist\s_curve_workbench.exe to target Windows 10/11 machine
echo   2. Run executable directly (no Python installation required)
pause
