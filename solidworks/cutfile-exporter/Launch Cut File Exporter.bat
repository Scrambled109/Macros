@echo off
setlocal
cd /d "%~dp0"

where py >nul 2>nul || (
  echo ERROR: Python was not found.
  echo Install Python and enable "Add Python to PATH".
  pause
  exit /b 1
)

py -c "import importlib.metadata as m, ezdxf, win32com.client; assert m.version('ezdxf') == '1.4.2'; assert int(m.version('pywin32')) >= 312" >nul 2>nul || (
  echo Installing the required cut-file exporter libraries...
  py -m pip install -r requirements.txt || (
    echo ERROR: Library installation failed.
    pause
    exit /b 1
  )
)

py cutfile_exporter.py
set "export_status=%errorlevel%"
if not "%export_status%"=="0" pause
exit /b %export_status%
