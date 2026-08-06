DXF ORCHESTRATOR — USER INSTRUCTIONS
====================================

Required files in this folder:
  Master_Orchestrator.py
  Master_Orchestrator.ps1
  ColortoLayer.lsp
  SPC_Seed.dwg

Required software: AutoCAD 2025 or 2026. The Python version uses only Python's
standard library; no extra Python package is needed.

1. Work on copies of the incoming DXF files.
2. Put the DXFs in numbered folders such as 100, 101, and 102 inside the job
   workspace.
3. Put "Parts List.csv" in the workspace root. It must contain PartNumber,
   Quantity, Thickness, and Material columns. Job Assistant creates this adapter
   after Parts List review.
4. Open Command Prompt or PowerShell in the workspace and run:

   py C:\Path\To\Macros\autocad\dxf-orchestrator\Master_Orchestrator.py

AutoCAD Core Console is detected automatically, preferring 2026 and falling
back to 2025 or another installed AutoCAD folder. Use --acad-console-path only
for a nonstandard installation. Use --workspace when the terminal is not open
in the job workspace.

The default is two parallel AutoCAD Core Console processes. Use --workers 1
through --workers 4 to change concurrency; use 1 if licensing or workstation
stability requires serial processing.

Every DXF is processed headlessly. Detected bevel drawings do not open graphical
AutoCAD; they receive the suffix "(B)" before ".dwg". Bevel callout text and
leader arrows are assigned to PLOT, while pin-stamp entities retain their
marking layers.

Original DXFs are archived only after a valid output exists. Failures remain
available for retry. Review _ORCHESTRATOR_LOGS, the final summary, and every
output before deleting or moving originals.
