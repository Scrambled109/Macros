# Engineering Macros and CAD Automation

This repository is a toolbox for repetitive **AutoCAD**, **SolidWorks**, BOM,
and production-data tasks. You do not need to know programming to use it, but
you do need the matching engineering software and you should always work on
**copies** of production files.

> **The three rules for a first run**
>
> 1. Copy a few test drawings into a new local folder (for example,
>    `C:\Macro Test`). Never make the first run on the only copy.
> 2. Read the section for the tool you plan to use. The tools are independent;
>    you do **not** run everything in this repository.
> 3. Inspect the result before processing a real job. CAD versions, templates,
>    layer names, drive letters, and company exports vary.

## What the computer terms mean

- A **file** is one document, drawing, spreadsheet, or program.
- A **folder** contains files. A path such as `C:\Macro Test\part.dwg` tells the
  computer which drive, folder, and file to use.
- A **macro/script** is a saved set of instructions for another program.
- **AutoLISP** (`.lsp`) and AutoCAD scripts (`.scr`) run inside AutoCAD.
- A **SolidWorks macro** (`.swp`) runs inside SolidWorks.
- **PowerShell** (`.ps1`) and batch files (`.bat`) run in Windows.
- **Python** (`.py`) is a separate free program used by the two data tools.
- **CSV** is a plain-text spreadsheet export. Do not rename an Excel workbook
  from `.xlsx` to `.csv`; use **File > Save As > CSV** in Excel.
- A **terminal** is a text window where commands are typed. On Windows, open
  one by right-clicking Start and selecting **Terminal** or **PowerShell**.

## Before you download or run anything

1. Use a **Windows** computer. The CAD automation targets AutoCAD 2026 and
   SolidWorks 2025; other versions may require path/reference changes.
2. On this GitHub page, select **Code > Download ZIP**, open Downloads,
   right-click the ZIP, select **Extract All**, and choose a local folder.
3. If Windows shows an **Unblock** checkbox under right-click **Properties** for
   the ZIP, check it before extracting.
4. Keep this original extracted folder unchanged. Copy only the tool you need
   into a test folder unless its instructions say otherwise.
5. Enable file-name extensions in File Explorer (**View > Show > File name
   extensions**) so `.dwg`, `.dxf`, `.csv`, and `.scr` are visible.

## Choose the right tool

| I want to… | Use | Required software |
|---|---|---|
| Compare SolidWorks, Parts List, and nest exports | `compare_production_parts.py` | Python 3 only |
| Copy DS rows from an A-BOM into a parts-list workbook | `bom_converter.py` | Python 3 + packages |
| Convert and sort numbered folders of DXFs, including bevel review | `CLBO_ps1_DIRECTORY` | AutoCAD 2026 + PowerShell |
| Convert every DXF in one folder to DWG | `BATCH_CONVERT` | AutoCAD 2026 |
| Zoom Extents and save every DWG in one folder | `process_dwgs.bat` | AutoCAD 2026 |
| Move AutoCAD objects to layers according to color | `ColortoLayer.lsp` | AutoCAD |
| Toggle temporary labels showing each object's layer | `LayerX.lsp` | AutoCAD |
| Turn every `#` in an open drawing into `-` | `H2D.scr` | AutoCAD |
| Filter DWGs and create extruded SolidWorks parts | `Main.RunBatch.swp` / `CAD_Batch_Converter` | AutoCAD 2026 + SolidWorks 2025 |
| Run one of the older SolidWorks utilities | `macros/*.swp` | SolidWorks; test copy required |

## 1. Production Part Reconciliation (recommended data-checking tool)

This is the safest place to start because it only **reads** the selected input
exports and writes reports into a new output folder. It does not edit CAD data.
Full input-column and report details are in
[`production_part_comparison_README.md`](production_part_comparison_README.md).

### Easiest run

1. Install Python from <https://www.python.org/downloads/windows/>. During
   setup, check **Add python.exe to PATH**.
2. Prepare the Parts List CSV, SolidWorks Assembly Visualization CSV, and a
   folder containing all linear-nesting CSV/TXT files.
3. Double-click `compare_production_parts.py`. If Windows asks which program,
   choose Python. File-selection windows ask for the inputs and output folder.
4. Read `production_part_comparison.xlsx` first. Begin with the **Errors
   Requiring Action** sheet. The HTML and CSV files contain the same audit in
   other formats.

### Terminal run (best when asking someone for support)

```powershell
cd "C:\path\to\Macros"
py compare_production_parts.py --nests "C:\Job\Nests" --parts "C:\Job\parts.csv" --solidworks "C:\Job\solidworks.csv" --output "C:\Job\Reports"
```

Edit `comparison_rules.json` only when you deliberately need different aliases,
keywords, checks, or tolerances. Keep a copy of the old rules.

## 2. A-BOM to Parts List converter

`bom_converter.py` reads the `Lofting` sheet, keeps part numbers beginning with
`DS`, removes repeated `ENG MAT ID` routing rows (keeping the last), and appends
the mapped data to an existing output workbook or a blank template. Existing
output rows are not erased, so do not run it twice unless you want another copy
of the rows.

### One-time installation

```powershell
cd "C:\path\to\Macros"
py -m pip install -r requirements.txt
```

### Every run

```powershell
py bom_converter.py "C:\Job\A-BOM.xlsm" "C:\Job\TEMP.xlsx" "C:\Job\MASTER_PARTS_LIST.xlsx"
```

The source must contain a `Lofting` sheet with `ENG MAT ID`, `Description`,
`QTY`, `Width`, `Length`, and `MTL Type` headings. Close the workbooks in Excel
before running. The program now stops with a readable error when a file or
required heading is missing rather than failing partway through.

## 3. Full DXF conversion and sorting orchestrator

This workflow changes layer assignments, creates DWGs, sorts output by stock,
and archives successfully processed source DXFs. Read
[`CLBO_ps1_DIRECTORY/README.txt`](CLBO_ps1_DIRECTORY/README.txt) before use.

1. Copy the entire `CLBO_ps1_DIRECTORY` folder to the computer.
2. Put raw DXFs inside numbered subfolders such as `100` and `101`.
3. Optionally put `parts.csv` beside `Master_Orchestrator.ps1`. Its headings
   must be `PartNumber`, `Quantity` (the misspelling `Quanity` is also accepted),
   `Thickness`, and `Material`.
4. `ColorToLayer.lsp` and `SPC_Seed.dwg` are already beside the script. The
   script now finds these files relative to itself, so another person's user
   name or mapped `U:` drive is no longer required.
5. Confirm AutoCAD is installed at
   `C:\Program Files\Autodesk\AutoCAD 2026`. If not, edit the two AutoCAD paths
   at the top of `Master_Orchestrator.ps1` in Notepad.
6. Open PowerShell in the folder containing the numbered folders and run:

```powershell
powershell -ExecutionPolicy Bypass -File ".\CLBO_ps1_DIRECTORY\Master_Orchestrator.ps1"
```

For a bevel drawing, AutoCAD opens. Review it, type `FINISH`, and press Enter.
Do not use Save As. Originals move to `_PROCESSED_DXF_ARCHIVE` **only after** a
valid, stable output exists. Failures remain in place; inspect
`_ORCHESTRATOR_LOGS`. `FINISH` closes the reviewed drawing but leaves AutoCAD
running, avoiding a full application restart before the next beveled part. If
`FINISH` is forgotten, the review times out after one hour by default and keeps
the original DXF for retry.

## 4. Simple DXF-to-DWG folder conversion

1. Make a backup folder containing the DXFs.
2. Copy all three `BATCH_CONVERT` files directly beside those DXFs.
3. Double-click `Run_Conversion.bat` and wait for **All files processed**.
4. Check every DWG before deleting any DXF.

This batch assumes AutoCAD 2026 is in its default installation folder. It now
stops cleanly when no DXFs exist and displays an error for an AutoCAD failure.

## 5. Zoom and save a folder of DWGs

Copy `process_dwgs.bat` and `zoom_save.scr` beside **copies** of the DWGs, then
double-click `process_dwgs.bat`. Each drawing is opened headlessly, zoomed to
Extents, and saved. This intentionally modifies each DWG. It assumes the
default AutoCAD 2026 path and now does nothing safely when the folder has no
DWGs.

## 6. Individual AutoCAD commands

For `.lsp` files, type `APPLOAD` in AutoCAD, select the file, choose **Load**,
then run its command. Add it to **Startup Suite** only after testing. More CUI
button instructions are in `setting up macros or scripts in AUTOCAD.txt`.

- `ColortoLayer.lsp` → command `ColorToLayer`: maps explicit colors 1, 2, 3,
  5, 6, and 7 to company layer names and changes the objects to ByLayer. The
  similarly named file in `CLBO_ps1_DIRECTORY` also moves ByLayer objects from
  layer 0 to `PIN STAMP TEXT`. Use the correct variant for the workflow.
- `LayerX.lsp` → command `LayerToggle`: first run adds non-plotting yellow layer
  labels; second run deletes them. Save only if you intentionally want to keep
  that temporary layer.
- `H2D.scr`: type `SCRIPT`, select this file, and it drives AutoCAD's Find
  command to replace `#` with `-` in the open drawing.
- `SPC_IMPORT.las` is an AutoCAD saved layer-state file;
  `SPC_IMPORT.cuix` is a custom user-interface package. `macros/macros.txt`
  contains the corresponding CUI macro text.

## 7. CAD Batch Converter (AutoCAD to SolidWorks)

`Main.RunBatch.swp` is the packaged macro. The editable source lives in
`CAD_Batch_Converter/*.bas`, and its detailed setup, architecture, recovery
behavior, output, and log instructions are in
[`CAD_Batch_Converter/README.md`](CAD_Batch_Converter/README.md).

This is an advanced, job-specific tool. Before running, edit the three folder
constants at the top of `CAD_Batch_Converter/Config.bas`; they currently point
to a specific company job. Also verify target/text layer names, drawing units,
and extrusion depth. In SolidWorks use **Tools > Macro > Run**, choose
`Main.RunBatch.swp`, and run `Main_RunBatch1.RunBatch`. Run the separate
`TextStamp1.RunTextStamp` pass only after validating the generated parts.

Developers rebuilding the `.swp` must import all eight `.bas` modules and add
the AutoCAD 2026, SolidWorks 2025, and SolidWorks constants type-library
references described by the component README.

## 8. Older SolidWorks macros

The files below are compiled VBA projects. Their VBA modules have now been
recovered into [`macros/source`](macros/source/README.md) for static review.
Those recovered files make the logic auditable, but they do not replace an
official SolidWorks VBA-editor export or a compile/run test:

- `drawing_auto.swp`
- `macros/(MOD)2(SECONDARY).swp`
- `macros/AUTOBOMACTUAL.swp`
- `macros/AutoBOMProperties.swp`
- `macros/Bounding box.swp`
- `macros/MaterialSpec(MOD).swp`
- `macros/Test_test.swp`
- `macros/hide sketches.swp`

Read the extracted-source inventory and safety notes first. Then open a
throwaway part/assembly, use **Tools > Macro > Edit** to confirm the entry
procedure, references, and recovered code before selecting **Tools > Macro >
Run**. Do not use `Test_test.swp` for production.

## Other reference files

- `Automation_Reference_CURRENT.docx` is a human reference document.
- `RENAME_SIGN_CONVERT_POWERSHELL.txt` contains a PowerShell rename routine as
  reference text; it renames files and has no preview/dry-run mode. Copy it to
  a `.ps1` only after testing on disposable data.
- `SPC_Seed.dwg` is the seed drawing used by the orchestrator.

## Troubleshooting

- **“py is not recognized”**: reinstall Python and select **Add Python to
  PATH**, or use `python` instead of `py`.
- **PowerShell will not run the script**: use the exact `-ExecutionPolicy
  Bypass` command above. This changes policy only for that process.
- **AutoCAD executable not found**: search `C:\Program Files\Autodesk` for
  `accoreconsole.exe`, then update the path near the top of the `.bat` or `.ps1`.
- **AutoCAD blocks a LISP**: copy it to a local trusted folder or add its folder
  under AutoCAD **Options > Files > Trusted Locations**. Do not globally lower
  security for files you did not inspect.
- **Wrong/missing layers**: layer names must match exactly. Undo immediately,
  close without saving, and adjust the mapping on a copy.
- **Nothing happens on a network drive**: copy the test job locally. AutoCAD
  and scripts can reject mapped drives, permissions, or special characters.
- **A run partially fails**: keep originals, capture the terminal message and
  logs, and retry only failed parts. Never assume “completed” means geometrically
  correct—visually inspect output.

## Audit status and maintenance notes

The text source was reviewed for obvious syntax, missing-input, path, and data
safety problems. The Python programs can be syntax-checked on any platform;
actual AutoCAD/SolidWorks automation can only be integration-tested on a
licensed Windows workstation with the specified versions and company files.
Binary `.swp`, `.dwg`, `.cuix`, and `.docx` contents require their native apps.

When changing a tool, update this README and its component README together.
Prefer adding new job paths as configuration rather than committing another
person's user folder. Keep generated reports, job drawings, workbooks, logs,
and processed archives outside this repository.
