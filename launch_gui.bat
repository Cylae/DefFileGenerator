@echo off
REM ==============================================================================
REM WebdynSunPM DefFileGenerator - Windows 11 Desktop Launcher
REM ==============================================================================
setlocal enabledelayedexpansion

cd /d "%~dp0"

IF EXIST ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" -m DefFileGenerator.gui
    exit /b 0
)

IF EXIST ".venv\Scripts\python.exe" (
    start "" ".venv\Scripts\python.exe" -m DefFileGenerator.gui
    exit /b 0
)

where pythonw >nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    start "" pythonw -m DefFileGenerator.gui
    exit /b 0
)

where python >nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    start "" python -m DefFileGenerator.gui
    exit /b 0
)

echo [ERROR] Python was not found in PATH or in .venv.
echo Please run 'uv sync' or install Python.
pause
exit /b 1
