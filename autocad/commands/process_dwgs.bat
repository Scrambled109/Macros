@echo off
echo Preparing safe script path...

set "ACCORE=C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe"
if not exist "%ACCORE%" (
    echo ERROR: AutoCAD Core Console was not found at:
    echo %ACCORE%
    echo Edit ACCORE near the top of this file if AutoCAD is installed elsewhere.
    pause
    exit /b 1
)

:: Copy script to Windows Temp folder to avoid AutoCAD choking on the "&" and "#" in your folder name
set "SAFE_SCRIPT=%TEMP%\zoom_save.scr"
copy /Y "%~dp0zoom_save.scr" "%SAFE_SCRIPT%" >nul

if not exist "%SAFE_SCRIPT%" (
    echo ERROR: Failed to copy script to Temp folder.
    pause
    exit /b
)

echo Starting batch processing...
echo.

if not exist "%~dp0*.dwg" (
    echo No DWG files were found in this folder. Nothing was changed.
    del "%SAFE_SCRIPT%" >nul 2>&1
    pause
    exit /b 0
)

:: Loop through all DWG files using the exact directory path
for %%f in ("%~dp0*.dwg") do (
    echo Processing %%~nxf...
    "%ACCORE%" /i "%%~f" /s "%SAFE_SCRIPT%"
    if errorlevel 1 echo ERROR: AutoCAD failed while processing %%~nxf.
)

:: Clean up the temporary script
del "%SAFE_SCRIPT%"

echo All files processed!
pause