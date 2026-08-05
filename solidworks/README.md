# SolidWorks Macros

The `.swp` files on this branch are the runnable SolidWorks macros. Run one from
SolidWorks using **Tools > Macro > Run**. Test on disposable or recoverable
documents before using production data.

Current workflow macros:

- `cad-batch-converter/Main.RunBatch.swp` — converts prepared DWGs into plate
  parts; normally launched as part of the Job Assistant workflow
- `auto-bom/AutoBOMProperties.swp` and `auto-bom/AUTOBOMACTUAL.swp` — update BOM
  properties and save models; validate which one is approved for the job
- `drawing-automation/drawing_auto.swp` — drawing automation
- `utilities/hide sketches.swp` — focused utility macro
- `drawing_auto.swp` and `flip_dwg.swp` — additional runnable uploads retained
  at their existing paths

Python/SolidWorks workflow tools:

- `cutfile-exporter/Launch Cut File Exporter.bat` — exports flat SolidWorks parts
  to validated layered DXFs with separate cut, line-marking, and text layers

The two macros under `legacy/bounding-box-test` are retained because this branch
contains all `.swp` files, but they are experimental and should not be used for
production work without validation.

Editable `.bas` source is intentionally hidden from this user branch. Developer
changes, tests, and rebuild work belong on `main`.

