@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul || (
  echo ERROR: Python was not found.
  echo Install Python and enable "Add Python to PATH".
  pause
  exit /b 1
)

py -c "import ezdxf, win32com.client" >nul 2>nul || (
  echo Installing the required cut-file exporter libraries...
  py -m pip install -r requirements.txt || (
    echo ERROR: Library installation failed.
    pause
    exit /b 1
  )
)

py cutfile_exporter.py
if errorlevel 1 pause
exit /b %errorlevel%

