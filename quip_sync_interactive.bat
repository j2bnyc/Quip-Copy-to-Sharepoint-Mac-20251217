@echo off
REM Quip Sync Interactive - Configure and run Quip folder sync (Windows)
REM Double-click or run from Command Prompt

setlocal EnableDelayedExpansion

echo ========================================
echo   Quip Folder Sync Tool (Windows)
echo ========================================
echo.

REM Get Quip folder URL
echo Enter Quip folder URL (root folder you want to copy)
echo   Example: https://quip-mycompany.com/ABC123XYZ/My-Team-Folder
set /p SOURCE_URL="> "
echo.

if "%SOURCE_URL%"=="" (
    echo ERROR: Quip folder URL is required
    echo Press any key to exit...
    pause >nul
    exit /b 1
)

REM Get destination path
echo Enter destination folder path (local folder synced to SharePoint)
echo   Example: C:\Users\username\OneDrive - mycompany.com\Team - Documents\Backup
set /p TARGET_PATH="> "
echo.

if "%TARGET_PATH%"=="" (
    echo ERROR: Target folder path is required
    echo Press any key to exit...
    pause >nul
    exit /b 1
)

REM Get PAT token
echo Enter your Quip API token
echo   Get from: https://quip-mycompany.com/dev/token
echo   (Note: input will be visible on Windows)
set /p QUIP_TOKEN="> "
echo.

if "%QUIP_TOKEN%"=="" (
    echo ERROR: Token is required
    echo Press any key to exit...
    pause >nul
    exit /b 1
)

REM Select sync mode
echo Select sync mode:
echo   1) Incremental - only sync modified files (default)
echo   2) Full - overwrite all files
echo.
set /p MODE_CHOICE="Enter choice [1/2]: "

if "%MODE_CHOICE%"=="2" (
    set "SYNC_MODE=full"
) else (
    set "SYNC_MODE=incremental"
)

echo.
echo ========================================
echo   Configuration Summary
echo ========================================
echo   Source: %SOURCE_URL%
echo   Target: %TARGET_PATH%
echo   Mode:   %SYNC_MODE%
echo ========================================
echo.
echo Press Enter to start sync (or Ctrl+C to cancel)...
pause >nul
echo.

REM Run sync
python quip_sync_cli.py --source "%SOURCE_URL%" --target "%TARGET_PATH%" --mode %SYNC_MODE% --token "%QUIP_TOKEN%"

echo.
echo ========================================
echo   Sync complete!
echo ========================================
echo Press any key to close...
pause >nul
