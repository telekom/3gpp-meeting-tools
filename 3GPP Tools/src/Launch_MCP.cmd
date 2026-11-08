@echo off
setlocal

:: Navigate to the project root directory
cd /d "%~dp0.."

:: Ensure both root and src/ are in PYTHONPATH
set "PYTHONPATH=%CD%;%CD%\src;%PYTHONPATH%"

:: Prefer py -3.14 when available, otherwise fall back to python
where py >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    py -3.14 -m mcp_server %*
) else (
    python -m mcp_server %*
)
