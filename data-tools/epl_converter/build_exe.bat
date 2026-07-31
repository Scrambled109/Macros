@echo off
setlocal
cd /d "%~dp0"

echo Building EPLConverter.exe...
pyinstaller --noconfirm --clean EPLConverter.spec
if errorlevel 1 (
    echo.
    echo Build failed. Confirm Python and requirements.txt are installed.
    exit /b 1
)

copy /Y "material_translations.json" "dist\material_translations.json" >nul
echo.
echo Build complete:
echo   %CD%\dist\EPLConverter.exe
echo   %CD%\dist\material_translations.json
endlocal
