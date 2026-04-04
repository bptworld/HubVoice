@echo off
setlocal
set /p DEVICE_IP=Enter satellite IP address: 
if not defined DEVICE_IP exit /b 1
start "" "http://%DEVICE_IP%:8080/"
