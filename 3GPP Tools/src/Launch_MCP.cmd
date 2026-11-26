@echo off
setlocal

:: Navigate to the project root directory
cd /d "%~dp0.."

:: Ensure both root and src/ are in PYTHONPATH
set "PYTHONPATH=%CD%;%CD%\src;%PYTHONPATH%"

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
    echo [ERROR] No Python interpreter found! >&2
    echo Please install 64-bit Python 3.10+ from https://www.python.org/ >&2
    exit /b 1
)

%PY_CMD% -m mcp_server %*
