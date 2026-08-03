@echo off
setlocal
cd /d "%~dp0"
set "DIST=%~dp0dist"
if exist "%DIST%" rmdir /s /q "%DIST%"
py -m pip install pyinstaller -r "..\data-tools\bom-converter\requirements.txt"
py -m PyInstaller --noconfirm --clean --onefile --windowed --distpath "%DIST%" --name "Engineering Job Assistant" job_assistant.py
if errorlevel 1 exit /b %errorlevel%
py -m PyInstaller --noconfirm --clean --onefile --console --distpath "%DIST%" --name "Engineering BOM Converter" "..\data-tools\bom-converter\bom_converter.py"
if errorlevel 1 exit /b %errorlevel%
py -m PyInstaller --noconfirm --clean --onefile --console --distpath "%DIST%" --name "Engineering Production Comparison" --add-data "..\data-tools\production-comparison\comparison_rules.json;." "..\data-tools\production-comparison\compare_production_parts.py"
if errorlevel 1 exit /b %errorlevel%
for %%E in ("Engineering Job Assistant.exe" "Engineering BOM Converter.exe" "Engineering Production Comparison.exe") do (
  if not exist "%DIST%\%%~E" (
    echo ERROR: Missing required build output: %DIST%\%%~E
    exit /b 1
  )
)
echo Built and verified the complete three-EXE distribution in: %DIST%
