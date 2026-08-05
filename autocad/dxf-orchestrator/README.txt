1. Prerequisites
To run this pipeline, you must have the following files prepared and accessible on your machine:

Master_Orchestrator.py: The canonical implementation launched by the Engineering
Job Assistant. It uses only the Python standard library.

Master_Orchestrator.ps1: A sequential PowerShell entry point with the same
automatic bevel-marking behavior.

Parts List.csv: A required, orchestrator-specific CSV containing the exact headers PartNumber, Quantity, Thickness, and Material. This is not the standard Parts List workbook's native column layout. When the operator completes the reviewed Parts List stage, the Job Assistant exports this CSV automatically, includes only rows whose DESCRIPTION is PL or PLATE, and copies it into each DXF workspace.

ColortoLayer.lsp: The AutoLISP routine responsible for mapping colors and moving bevel annotations/arrows to PLOT.

SPC_Seed.dwg: A blank AutoCAD drawing containing the master layer definitions. The script temporarily injects this into every part to ensure the layer database exists before running the LISP routine.

2. Directory Structure
The script scans for raw DXF files stored inside numbered subdirectories. Your workspace should look like this before running the script. parts of this are option such as moving COolorToLayer.lsp and SPC_Seed; both should be connected through the shared drive natively and do not need to be moved. if they are moved file path must be updated in the orchestrator.:


WORKSPACE_FOLDER/
│
├── Master_Orchestrator.ps1
├── Parts List.csv
├── ColortoLayer.lsp
├── SPC_Seed.dwg
│
├── 100/                       <-- Numbered folder containing DXFs
│   ├── CleanPart_A.dxf
│   └── BeveledPart_B.dxf
│
└── 101/                       <-- Another numbered folder
    └── CleanPart_C.dxf
3. First-Time Configuration
By default, the Python script looks for Parts List.csv in the workspace and for
ColortoLayer.lsp and SPC_Seed.dwg beside the script. Use --parts-list to supply
a different CSV path.

Open the .ps1 file in a text editor (like Notepad) and locate the # --- CONFIGURATION PATHS --- section. Update the absolute paths to match where the required files are stored on your specific machine:

PowerShell
$CsvPath         = Join-Path $PSScriptRoot "parts.csv"
$LspPath         = Join-Path $PSScriptRoot "ColortoLayer.lsp"
$SeedPath        = Join-Path $PSScriptRoot "SPC_Seed.dwg"
$AcadConsolePath = "C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe"
(Verify that the AutoCAD version year in the paths matches the version installed on your machine).

4. How to Run the Pipeline
Windows restricts the execution of unsigned scripts by default. To run the Orchestrator, you must bypass the execution policy for your current session.

Open Windows PowerShell.

Use the cd command to navigate to the folder containing your script.

PowerShell
cd C:\Path\To\Your\Workspace
Execute the script with the bypass flag:

PowerShell
powershell -ExecutionPolicy Bypass -File .\Master_Orchestrator.ps1
The AutoCAD Core Console path defaults to AutoCAD 2026. The Engineering Job
Assistant can override it safely; paths containing spaces are supported.

Alternatively, run the Python implementation from the workspace that contains
the numbered input folders:

Command Prompt
py C:\Path\To\Macros\autocad\dxf-orchestrator\Master_Orchestrator.py

It also defaults to AutoCAD 2026. Override installed locations when needed:

Command Prompt
py C:\Path\To\Macros\autocad\dxf-orchestrator\Master_Orchestrator.py --acad-console-path "C:\Path\To\accoreconsole.exe"

Use `--workspace C:\Path\To\Workspace` if the command is not launched from the
folder containing the numbered input directories. Run `py ... --help` for the
timeout options. The Python file does not require third-party packages. The Job
Assistant launches this implementation instead of the PowerShell script.

All drawings use two AutoCAD Core Console workers by default:

Command Prompt
py C:\Path\To\Macros\autocad\dxf-orchestrator\Master_Orchestrator.py --workers 2

Choose 1 through 4. Two is recommended because it improves throughput without
starting an unbounded number of AutoCAD processes. Use `--workers 1` if the
workstation or AutoCAD licensing is unstable. Each job uses a unique script and
log. Bevel drawings use the same headless path and never open graphical AutoCAD.

5. The Operator Workflow
Once the script starts, it requires minimal input. Watch the PowerShell terminal for color-coded status updates.

The Background Auto-Pilot
Every part runs through accoreconsole.exe. The orchestrator injects the standard
layers, swaps hashtags to dashes, saves the file as a DWG, and moves it to a
sorted folder such as 250-DH-36. No AutoCAD interaction is required.

After the normal color mapping, bevel-related TEXT/MTEXT and every
LEADER/MLEADER/MULTILEADER arrow are forced onto the `PLOT` layer. Ordinary
pin-stamp line marking and pin-stamp text remain on their existing layers.

If the DXF contains a bevel flag (K, V, BEVEL/BVL, SNIPE, CHAMFER, an angle
callout such as V22.5, V-22.5, RV9, or K30, or an arrow/leader annotation), the
output filename receives the exact suffix `(B)` before `.dwg`. Example:
`PartName_250-DH-36_5(B).dwg`. Detection deliberately favors one extra marker
over missing a potentially beveled part.

Before parallel work starts, duplicate inputs that resolve to the same output
DWG are rejected with a readable message and both originals are retained. This
prevents two AutoCAD workers from racing to overwrite one output.

If a headless job fails or hangs, the script logs it (see _ORCHESTRATOR_LOGS), keeps the original DXF, and moves on to the next part instead of stalling the whole batch. A per-job timeout (default 180s, set via $ConsoleTimeoutSec) guards against a wedged accoreconsole.

Bevel Identification (Magenta)
Detected bevel parts print a magenta status line, receive `(B)` in the output
filename, and otherwise process exactly like standard parts. The orchestrator
does not launch acad.exe, wait for an operator, or use SPCFINISH.

6. Output and Data Safety
Sorted Output: Processed parts will appear in newly generated folders at the root of your workspace named according to the spreadsheet data (e.g., 500-A-36, 250-DH-36).

Appended Naming: The final .dwg files append the target folder name and quantity
(e.g., PartName_250-DH-36_5.dwg). Detected bevel parts additionally end in
`(B).dwg`.

The Archive Failsafe: The script will never delete your original .dxf files. Once a file is successfully processed and the DWG size is validated (and AutoCAD has fully released its lock on the new DWG), the original raw DXF is moved into a _PROCESSED_DXF_ARCHIVE folder. If a batch fails or requires re-processing, you can always retrieve the pristine original files from the archive.

Diagnostics: Each headless conversion writes its console output to _ORCHESTRATOR_LOGS\<part>.log (and .err.log for errors), so a failed part can be diagnosed instead of just reporting "DWG missing." The run ends with a summary line counting Standard / Bevel-marked / Failed parts.

Pre-flight checks: Before doing any work, the script verifies that the seed DWG, ColortoLayer.lsp, and accoreconsole.exe exist, and aborts with a clear message if any are missing. The required Parts List.csv is validated before processing starts.
