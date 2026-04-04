@echo off
setlocal
set "ROOT=%~dp0"
set "PYTHON=%ROOT%.envs\runtime\Scripts\python.exe"

if not exist "%PYTHON%" (
  echo Runtime Python is not installed in this repo at "%PYTHON%"
  exit /b 1
)

"%PYTHON%" -m esphome %*
exit /b %ERRORLEVEL%
