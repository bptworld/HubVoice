@echo off
setlocal

set "ROOT=%~dp0"
set "PYTHON=%ROOT%.envs\runtime\Scripts\python.exe"
set "PATH=%ROOT%.envs\runtime\Scripts;%ROOT%;%PATH%"
for /f "tokens=1,* delims=:" %%A in ('findstr /b /c:"  firmware_version:" "%ROOT%hubvoice-sat.yaml"') do if not defined FW_VERSION set "FW_VERSION=%%B"
for /f "tokens=1,* delims=:" %%A in ('findstr /b /c:"  device_name:" "%ROOT%hubvoice-sat.yaml"') do if not defined DEVICE_NAME set "DEVICE_NAME=%%B"
for /f "tokens=1,* delims=:" %%A in ('findstr /b /c:"  friendly_name:" "%ROOT%hubvoice-sat.yaml"') do if not defined FRIENDLY_NAME set "FRIENDLY_NAME=%%B"
if defined FW_VERSION set "FW_VERSION=%FW_VERSION: =%"
if defined FW_VERSION set "FW_VERSION=%FW_VERSION:"=%"
if defined DEVICE_NAME set "DEVICE_NAME=%DEVICE_NAME: =%"
if defined DEVICE_NAME set "DEVICE_NAME=%DEVICE_NAME:"=%"
if defined FRIENDLY_NAME set "FRIENDLY_NAME=%FRIENDLY_NAME: =%"
if defined FRIENDLY_NAME set "FRIENDLY_NAME=%FRIENDLY_NAME:"=%"
if not defined FW_VERSION set "FW_VERSION=unknown"
if not defined DEVICE_NAME set "DEVICE_NAME=hubvoice-sat"
if not defined FRIENDLY_NAME set "FRIENDLY_NAME=%DEVICE_NAME%"
set "SATELLITE_MAP=%ROOT%satellites.csv"

if not exist "%PYTHON%" (
  echo Runtime Python is not installed in this repo at "%PYTHON%"
  echo.
  echo Run this once from C:\HubVoice:
  echo   .\.envs\runtime\Scripts\python.exe -m pip install esphome==2026.2.4
  echo.
  pause
  exit /b 1
)

set "TARGET_DEVICE=%DEVICE_NAME%"
set "TARGET_FRIENDLY=%FRIENDLY_NAME%"
set "TARGET_IP="
set "TARGET_IP_FROM_MAP="

if "%~1"=="" (
  echo Known satellite names on this PC:
  if exist "%SATELLITE_MAP%" (
    for /f "tokens=1,2 delims=," %%A in (%SATELLITE_MAP%) do echo   %%A  [%%B]
  ) else if exist "%ROOT%.esphome\build" (
    for /d %%D in ("%ROOT%.esphome\build\*") do echo   %%~nxD
  ) else (
    echo   ^(none yet^)
  )
  echo.
  set /p TARGET_DEVICE=Satellite device name [%DEVICE_NAME%]: 
  if not defined TARGET_DEVICE set "TARGET_DEVICE=%DEVICE_NAME%"
  set "TARGET_FRIENDLY=%TARGET_DEVICE%"
)

if exist "%SATELLITE_MAP%" (
  for /f "tokens=1,2 delims=," %%A in ('findstr /i /b /c:"%TARGET_DEVICE%," "%SATELLITE_MAP%"') do (
    set "TARGET_IP=%%B"
    set "TARGET_IP_FROM_MAP=1"
  )
)

if "%~1"=="" (
  if defined TARGET_IP (
    echo Saved OTA IP for %TARGET_DEVICE%: %TARGET_IP%
  ) else (
    set /p TARGET_IP=Satellite IP address ^(leave blank to try .local^): 
  )
  echo.
)

if "%~1"=="" if defined TARGET_IP if not defined TARGET_IP_FROM_MAP call :save_ip

echo ========================================
echo HubVoiceSat Flash Helper
echo Device Name: %TARGET_DEVICE%
echo Firmware Version: %FW_VERSION%
if defined TARGET_IP (
  echo OTA Target: %TARGET_IP% ^(saved for %TARGET_DEVICE%^)
) else (
  echo OTA Target: %TARGET_DEVICE%.local
)
echo ========================================
echo.

if "%~1"=="" goto ota
if /I "%~1"=="config" goto config
if /I "%~1"=="compile" goto compile
if /I "%~1"=="ota" goto ota
if /I "%~1"=="usb" goto usb
if /I "%~1"=="run" goto run

echo Usage:
echo   flash.bat config
echo   flash.bat compile
echo   flash.bat ota
echo   flash.bat usb
echo   flash.bat run
exit /b 1

:config
"%PYTHON%" -m esphome -s device_name "%TARGET_DEVICE%" -s friendly_name "%TARGET_FRIENDLY%" config "%ROOT%hubvoice-sat.yaml"
goto done

:compile
"%PYTHON%" -m esphome -s device_name "%TARGET_DEVICE%" -s friendly_name "%TARGET_FRIENDLY%" compile "%ROOT%hubvoice-sat.yaml"
goto done

:usb
"%PYTHON%" -m esphome -s device_name "%TARGET_DEVICE%" -s friendly_name "%TARGET_FRIENDLY%" run "%ROOT%hubvoice-sat.yaml" --device COM3
goto done

:ota
if defined TARGET_IP (
  "%PYTHON%" -m esphome -s device_name "%TARGET_DEVICE%" -s friendly_name "%TARGET_FRIENDLY%" run "%ROOT%hubvoice-sat.yaml" --device "%TARGET_IP%"
) else (
  "%PYTHON%" -m esphome -s device_name "%TARGET_DEVICE%" -s friendly_name "%TARGET_FRIENDLY%" run "%ROOT%hubvoice-sat.yaml" --device "%TARGET_DEVICE%.local"
)
if not "%ERRORLEVEL%"=="0" goto ota_ip_retry
goto done

:run
if defined TARGET_IP (
  "%PYTHON%" -m esphome -s device_name "%TARGET_DEVICE%" -s friendly_name "%TARGET_FRIENDLY%" run "%ROOT%hubvoice-sat.yaml" --device "%TARGET_IP%"
) else (
  "%PYTHON%" -m esphome -s device_name "%TARGET_DEVICE%" -s friendly_name "%TARGET_FRIENDLY%" run "%ROOT%hubvoice-sat.yaml" --device "%TARGET_DEVICE%.local"
)
if not "%ERRORLEVEL%"=="0" goto ota_ip_retry
goto done

:ota_ip_retry
echo.
if defined TARGET_IP (
  echo Saved OTA IP did not respond.
) else (
  echo OTA hostname did not respond.
)
set /p TARGET_IP=Enter satellite IP address to retry, or press Enter to skip: 
if not defined TARGET_IP goto done
echo.
"%PYTHON%" -m esphome -s device_name "%TARGET_DEVICE%" -s friendly_name "%TARGET_FRIENDLY%" run "%ROOT%hubvoice-sat.yaml" --device "%TARGET_IP%"
if "%ERRORLEVEL%"=="0" call :save_ip
goto done

:save_ip
if exist "%SATELLITE_MAP%" (
  powershell -NoProfile -Command "$path = '%SATELLITE_MAP%'; $name = '%TARGET_DEVICE%'; $ip = '%TARGET_IP%'; $rows = @(); if (Test-Path $path) { $rows = Get-Content $path | Where-Object { $_ -and ($_ -notmatch ('^' + [regex]::Escape($name) + ',')) } }; $rows += ($name + ',' + $ip); Set-Content -Path $path -Value $rows"
) else (
  >"%SATELLITE_MAP%" echo %TARGET_DEVICE%,%TARGET_IP%
)
exit /b 0

:done
set "RESULT=%ERRORLEVEL%"
echo.
if "%RESULT%"=="0" (
  echo Finished successfully.
) else (
  echo Finished with error code %RESULT%.
)
pause
exit /b %RESULT%
