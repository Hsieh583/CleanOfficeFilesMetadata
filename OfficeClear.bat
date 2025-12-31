@echo off
REM OfficeClear.bat - Launcher for Office Metadata Cleaner Tool
REM Version 2.0

echo Starting Office Metadata Cleaner Tool...
echo.

powershell -ExecutionPolicy Bypass -File "%~dp0OfficeClear_Tool.ps1"

if errorlevel 1 (
    echo.
    echo An error occurred while running the tool.
    pause
)
