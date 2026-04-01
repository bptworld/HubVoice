@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PS_SCRIPT=%SCRIPT_DIR%restart-runtime.ps1"

if not exist "%PS_SCRIPT%" (
    echo ERROR: Could not find restart-runtime.ps1 in "%SCRIPT_DIR%"
    exit /b 1
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS_SCRIPT%"
set "EXIT_CODE=%ERRORLEVEL%"

if not "%EXIT_CODE%"=="0" (
    echo Runtime restart failed with exit code %EXIT_CODE%.
)

exit /b %EXIT_CODE%
