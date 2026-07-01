@echo off
echo Preparing safe script path...

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

:: Loop through all DWG files using the exact directory path
for %%f in ("%~dp0*.dwg") do (
    echo Processing %%~nxf...
    "C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe" /i "%%~f" /s "%SAFE_SCRIPT%"
)

:: Clean up the temporary script
del "%SAFE_SCRIPT%"

echo All files processed!
pause