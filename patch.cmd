@echo off
setlocal
if exist "%~dp0patch.exe" (
	"%~dp0patch.exe" %*
	exit /b %ERRORLEVEL%
)

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0patch-wrapper.ps1" %*
exit /b %ERRORLEVEL%
