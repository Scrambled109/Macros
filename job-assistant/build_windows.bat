@echo off
setlocal
cd /d "%~dp0"
set "DIST=%~dp0dist"
set "APPDIR=%DIST%\Engineering Job Assistant"
if exist "%DIST%" rmdir /s /q "%DIST%"
py -m pip install pyinstaller -r "..\data-tools\bom-converter\requirements.txt"
if errorlevel 1 exit /b %errorlevel%
rem Use onedir for the GUI. Unlike onefile, it does not unpack an unsigned GUI
rem executable into %%TEMP%% on every launch, a pattern endpoint protection can
rem mistake for malware. Distribute the whole APPDIR folder.
py -m PyInstaller --noconfirm --clean --onedir --windowed --distpath "%DIST%" --name "Engineering Job Assistant" job_assistant.py
if errorlevel 1 exit /b %errorlevel%
py -m PyInstaller --noconfirm --clean --onefile --console --distpath "%APPDIR%" --name "Engineering BOM Converter" "..\data-tools\bom-converter\bom_converter.py"
if errorlevel 1 exit /b %errorlevel%
py -m PyInstaller --noconfirm --clean --onefile --console --distpath "%APPDIR%" --name "Engineering Production Comparison" --add-data "..\data-tools\production-comparison\comparison_rules.json;." "..\data-tools\production-comparison\compare_production_parts.py"
if errorlevel 1 exit /b %errorlevel%
py -m PyInstaller --noconfirm --clean --onefile --console --distpath "%APPDIR%" --name "Engineering DXF Orchestrator" --add-data "..\autocad\dxf-orchestrator\ColorToLayer.lsp;." --add-data "..\autocad\dxf-orchestrator\SPC_Seed.dwg;." "..\autocad\dxf-orchestrator\Master_Orchestrator.py"
if errorlevel 1 exit /b %errorlevel%
for %%E in ("Engineering Job Assistant.exe" "Engineering BOM Converter.exe" "Engineering Production Comparison.exe" "Engineering DXF Orchestrator.exe") do (
  if not exist "%APPDIR%\%%~E" (
    echo ERROR: Missing required build output: %APPDIR%\%%~E
    exit /b 1
  )
)
echo Built and verified the application folder: %APPDIR%
echo Keep the entire folder together when copying it to another workstation.
