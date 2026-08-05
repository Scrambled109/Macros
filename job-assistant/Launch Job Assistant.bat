@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul || (
  echo ERROR: Python was not found.
  echo Install Python and enable "Add Python to PATH".
  pause
  exit /b 1
)

py -c "import openpyxl" >nul 2>nul || (
  echo ERROR: Required Python libraries are not installed.
  echo From the repository root run:
  echo   py -m pip install -r requirements.txt
  pause
  exit /b 1
)

py job_assistant.py
if errorlevel 1 pause
exit /b %errorlevel%

