SIMPLE DXF-TO-DWG CONVERSION
============================

IMPORTANT: Work on copies of your DXF files the first time.

1. Copy Run_Conversion.bat and batch_convert.scr directly into the folder that
   contains the DXF files. Keep the two helper files together.
2. Double-click Run_Conversion.bat.
3. Wait until the window says "All files processed", then press a key.
4. Refresh File Explorer and inspect the new DWG files in AutoCAD.

The original DXF files are not intentionally deleted. Keep them until every
new DWG has been checked.

The batch expects AutoCAD 2026 in its normal installation location. If it says
that AutoCAD Core Console was not found, right-click Run_Conversion.bat, choose
Edit in Notepad, and change the ACCORE path near the top to the location of
accoreconsole.exe on this computer.

If no DXFs are present, the batch stops without changing anything. If AutoCAD
reports an error for a file, inspect that file manually; do not assume that its
DWG was created correctly.
