@echo off
setlocal
cd /d "%~dp0"
if /i "%~1"=="--source" (
  where py >nul 2>nul || (
    echo ERROR: Developer source mode requires Python's py launcher.
    exit /b 1
  )
  py job_assistant.py
  exit /b %errorlevel%
)
if exist "dist\Engineering Job Assistant.exe" (
  start "" "dist\Engineering Job Assistant.exe"
  exit /b 0
)
echo ERROR: dist\Engineering Job Assistant.exe is missing.
echo Build the complete package with build_windows.bat or reinstall it.
echo Developers may explicitly use: Launch Job Assistant.bat --source
exit /b 1
