@echo off
setlocal

:: Navigate to the project root directory
cd /d "%~dp0.."

:: Ensure both root and src/ are in PYTHONPATH
set "PYTHONPATH=%CD%;%CD%\src;%PYTHONPATH%"

echo ====================================================
echo Starting 3GPP Delegate Tools...
echo Root Directory : %CD%

:: Detect available Python interpreter (prefer py -3.14, then py -3, then python)
set "PY_CMD="
py -3.14 --version >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PY_CMD=py -3.14"
) else (
    py -3 --version >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        set "PY_CMD=py -3"
    ) else (
        where python >nul 2>nul
        if %ERRORLEVEL% EQU 0 (
            set "PY_CMD=python"
        )
    )
)

if "%PY_CMD%"=="" (
    echo [ERROR] No Python interpreter found!
    echo Please install 64-bit Python 3.10+ from https://www.python.org/
    echo ====================================================
    pause
    exit /b 1
)

%PY_CMD% --version
echo Launcher       : %PY_CMD%
echo ====================================================

:: Run application
%PY_CMD% -m main_tools

:: Only pause if the application crashed or closed with an error code
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Application exited with error code %ERRORLEVEL%.
    pause
)
