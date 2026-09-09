@echo off
setlocal

:: Navigate to the project root directory
cd /d "%~dp0.."

:: Ensure both root and src/ are in PYTHONPATH
set "PYTHONPATH=%CD%;%CD%\src;%PYTHONPATH%"

echo ====================================================
echo Starting 3GPP Delegate Tools...
echo Root Directory : %CD%
py -3.11 --version
echo ====================================================

:: Run application
py -3.11 -m main_tools

:: Only pause if the application crashed or closed with an error code
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Application exited with error code %ERRORLEVEL%.
    pause
)