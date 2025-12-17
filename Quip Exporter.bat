@echo off
REM Quip Exporter GUI Launcher
REM This batch file launches the Quip Exporter GUI

REM Change to the directory where this batch file is located
cd /d "%~dp0"

REM Check if the Python file exists
if not exist "quip_exporter_gui.py" (
    echo ERROR: quip_exporter_gui.py not found
    echo Please make sure this batch file is in the same folder as quip_exporter_gui.py
    pause
    exit /b 1
)

REM Try Method 1: pythonw.exe (windowless - best option)
pythonw.exe --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Launching Quip Exporter...
    start "" pythonw.exe quip_exporter_gui.py
    exit /b 0
)

REM Try Method 2: python.exe (with minimized window)
python.exe --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Launching Quip Exporter...
    start "" /min python.exe quip_exporter_gui.py
    exit /b 0
)

REM Try Method 3: py.exe launcher
py.exe --version >nul 2>&1
if %errorlevel% equ 0 (
    echo Launching Quip Exporter...
    start "" /min py.exe quip_exporter_gui.py
    exit /b 0
)

REM All methods failed
echo ERROR: Python not found
echo.
echo Please install Python from https://python.org
echo Make sure to check "Add Python to PATH" during installation
echo.
echo Alternative: Run "Launch Quip Exporter - Debug.bat" for detailed diagnostics
echo.
pause