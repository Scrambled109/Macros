@echo off
echo Preparing safe script path...

set "ACCORE="
if defined ACAD_CONSOLE_PATH set "ACCORE=%ACAD_CONSOLE_PATH%"
if not defined ACCORE if exist "%ProgramFiles%\Autodesk\AutoCAD 2026\accoreconsole.exe" set "ACCORE=%ProgramFiles%\Autodesk\AutoCAD 2026\accoreconsole.exe"
if not defined ACCORE if exist "%ProgramFiles%\Autodesk\AutoCAD 2025\accoreconsole.exe" set "ACCORE=%ProgramFiles%\Autodesk\AutoCAD 2025\accoreconsole.exe"
for /f "delims=" %%A in ('dir /b /ad /o-n "%ProgramFiles%\Autodesk\AutoCAD *" 2^>nul') do if not defined ACCORE if exist "%ProgramFiles%\Autodesk\%%A\accoreconsole.exe" set "ACCORE=%ProgramFiles%\Autodesk\%%A\accoreconsole.exe"
if not exist "%ACCORE%" (
    echo ERROR: AutoCAD Core Console was not found.
    echo Install AutoCAD 2025/2026 or set ACAD_CONSOLE_PATH.
    pause
    exit /b 1
)
echo Using %ACCORE%

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