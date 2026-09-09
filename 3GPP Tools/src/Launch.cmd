@echo off
setlocal

:: Ensure the script runs from the project root directory
cd /d "%~dp0.."

:: Inject both root and src into PYTHONPATH
set "PYTHONPATH=%CD%;%CD%\src;%PYTHONPATH%"

echo ====================================================
echo Starting 3GPP Delegate Tools...
echo Root Directory : %CD%
python --version
echo ====================================================

python -m main_tools

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Application exited with error code %ERRORLEVEL%.
    pause
)