@echo off
setlocal
set "ROOT=%~dp0"
set "ESPHOME=%ROOT%.runtime-venv\Scripts\esphome.exe"

if not exist "%ESPHOME%" (
  echo ESPHome is not installed in this repo at "%ESPHOME%"
  exit /b 1
)

"%ESPHOME%" %*
exit /b %ERRORLEVEL%
