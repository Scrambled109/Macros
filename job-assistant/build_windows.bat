@echo off
setlocal
cd /d "%~dp0"
py -m pip install pyinstaller -r "..\data-tools\bom-converter\requirements.txt"
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "Engineering Job Assistant" job_assistant.py
if errorlevel 1 exit /b %errorlevel%
py -m PyInstaller --noconfirm --clean --onefile --console --name "Engineering BOM Converter" "..\data-tools\bom-converter\bom_converter.py"
if errorlevel 1 exit /b %errorlevel%
py -m PyInstaller --noconfirm --clean --onefile --console --name "Engineering Production Comparison" --add-data "..\data-tools\production-comparison\comparison_rules.json;." "..\data-tools\production-comparison\compare_production_parts.py"
if errorlevel 1 exit /b %errorlevel%
echo Built the complete three-EXE distribution in: %CD%\dist
