@echo off
setlocal
cd /d "%~dp0"
if exist "dist\Engineering Job Assistant.exe" (
  start "" "dist\Engineering Job Assistant.exe"
  exit /b 0
)
where py >nul 2>nul || (
  echo Standalone application not found. Engineers should use the packaged EXE.
  echo Developers: install Python and run build_windows.bat first.
  pause
  exit /b 1
)
py job_assistant.py
