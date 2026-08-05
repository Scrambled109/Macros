DXF ORCHESTRATOR — USER INSTRUCTIONS
====================================

Required files in this folder:
  Master_Orchestrator.py
  Master_Orchestrator.ps1
  ColortoLayer.lsp
  SPC_Seed.dwg

Required software: AutoCAD 2026. The Python version uses only Python's standard
library; no extra Python package is needed.

1. Work on copies of the incoming DXF files.
2. Put the DXFs in numbered folders such as 100, 101, and 102 inside the job
   workspace.
3. Put "Parts List.csv" in the workspace root. It must contain PartNumber,
   Quantity, Thickness, and Material columns. Job Assistant creates this adapter
   automatically after Parts List review.
4. Open Command Prompt or PowerShell in the workspace and run:

   py C:\Path\To\Macros\autocad\dxf-orchestrator\Master_Orchestrator.py

Use --workspace when the terminal is not opened in the job workspace. Use
--acad-console-path and --acad-gui-path for nonstandard AutoCAD installs. Run
the script with --help to see all options.

Clean parts run through AutoCAD Core Console. Possible bevels open for manual
review. After checking each review tab, type SPCFINISH in AutoCAD. Do not use
Save As. Original DXFs are moved to _PROCESSED_DXF_ARCHIVE only after a valid
output exists; failures remain available for retry.

Review _ORCHESTRATOR_LOGS and the final summary. Visually inspect every output
before deleting or moving any original file.

