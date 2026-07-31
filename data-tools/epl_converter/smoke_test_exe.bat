@echo off
setlocal
cd /d "%~dp0"

if "%~2"=="" (
    echo Usage: smoke_test_exe.bat "C:\Path\To\LOAxxxxxx-EPL.xlsx" "C:\Path\To\BOP.xlsm"
    exit /b 2
)

if not exist "dist\EPLConverter.exe" (
    echo dist\EPLConverter.exe was not found. Run build_exe.bat first.
    exit /b 2
)

if not exist "%~1" (
    echo EPL file not found: %~1
    exit /b 2
)

if not exist "%~2" (
    echo BOP file not found: %~2
    exit /b 2
)

set "SMOKE_DIR=%TEMP%\EPLConverterSmoke_%RANDOM%"
mkdir "%SMOKE_DIR%"

"dist\EPLConverter.exe" --epl "%~1" --bop "%~2" --output-dir "%SMOKE_DIR%" --name "Smoke Test" --json > "%SMOKE_DIR%\result.json"
if errorlevel 1 (
    echo EXE CLI smoke test failed.
    echo Results folder: %SMOKE_DIR%
    exit /b 1
)

for %%F in (
    "Smoke Test - Plates.xlsx"
    "Smoke Test - Shapes.xlsx"
    "Smoke Test - Conversion Report.xlsx"
    "Smoke Test - Conversion Report.json"
) do (
    if not exist "%SMOKE_DIR%\%%~F" (
        echo Missing expected output: %%~F
        echo Results folder: %SMOKE_DIR%
        exit /b 1
    )
)

echo EXE and CLI smoke test passed without network access.
echo Results folder: %SMOKE_DIR%
type "%SMOKE_DIR%\result.json"
endlocal
