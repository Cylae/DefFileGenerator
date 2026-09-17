@echo off
REM ==============================================================================
REM WebdynSunPM DefFileGenerator - Standalone Windows Executable Builder (.bat)
REM ==============================================================================
setlocal enabledelayedexpansion

cd /d "%~dp0"

echo ====================================================================
echo  WebdynSunPM DefFileGenerator - Windows Executable (.exe) Compiler
echo ====================================================================

SET PYTHON_CMD=

IF EXIST ".venv\Scripts\python.exe" (
    SET PYTHON_CMD=".venv\Scripts\python.exe"
    goto :found_python
)

where python >nul 2>&1
IF %ERRORLEVEL% EQU 0 (
    SET PYTHON_CMD=python
    goto :found_python
)

echo [ERROR] Python was not found in PATH or in .venv.
echo Please install Python 3.10+ or configure your virtual environment.
pause
exit /b 1

:found_python
echo [INFO] Using Python: %PYTHON_CMD%

REM Ensure PyInstaller is installed
%PYTHON_CMD% -c "import PyInstaller" >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [INFO] Installing PyInstaller and build dependencies...
    %PYTHON_CMD% -m pip install -e ".[gui,build]"
    IF %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Failed to install build dependencies.
        pause
        exit /b 1
    )
)

echo [INFO] Running build_exe.py...
%PYTHON_CMD% build_exe.py %*

IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Build failed! Check the output messages above.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo ====================================================================
echo  [SUCCESS] Standalone executables compiled in dist/ folder:
echo    - dist\DefFileGenerator-GUI.exe (Windows Desktop UI)
echo    - dist\deffilegen.exe          (Command Line Interface)
echo ====================================================================
pause
exit /b 0
