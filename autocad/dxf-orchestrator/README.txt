1. Prerequisites
To run this pipeline, you must have the following files prepared and accessible on your machine:

Master_Orchestrator.ps1: The main PowerShell script.

parts.csv: A spreadsheet containing PartNumber, Quantity, Thickness, and Material columns. (Note: not actually required for the script to run just wont organize or append the file names. EVERY OTHER COMPONENT IS REQUIRED FOR OPERATION).

ColorToLayer.lsp: The AutoLISP routine responsible for mapping specific object colors to their correct layers.

SPC_Seed.dwg: A blank AutoCAD drawing containing the master layer definitions. The script temporarily injects this into every part to ensure the layer database exists before running the LISP routine.

2. Directory Structure
The script scans for raw DXF files stored inside numbered subdirectories. Your workspace should look like this before running the script. parts of this are option such as moving COolorToLayer.lsp and SPC_Seed; both should be connected through the shared drive natively and do not need to be moved. if they are moved file path must be updated in the orchestrator.:


WORKSPACE_FOLDER/
│
├── Master_Orchestrator.ps1
├── parts.csv
├── ColorToLayer.lsp
├── SPC_Seed.dwg
│
├── 100/                       <-- Numbered folder containing DXFs
│   ├── CleanPart_A.dxf
│   └── BeveledPart_B.dxf
│
└── 101/                       <-- Another numbered folder
    └── CleanPart_C.dxf
3. First-Time Configuration
By default, the script looks for parts.csv, ColorToLayer.lsp, and SPC_Seed.dwg
in the same folder as Master_Orchestrator.ps1. The CSV is optional. You only
need to edit these paths if you deliberately store those files elsewhere.

Open the .ps1 file in a text editor (like Notepad) and locate the # --- CONFIGURATION PATHS --- section. Update the absolute paths to match where the required files are stored on your specific machine:

PowerShell
$CsvPath         = Join-Path $PSScriptRoot "parts.csv"
$LspPath         = Join-Path $PSScriptRoot "ColorToLayer.lsp"
$SeedPath        = Join-Path $PSScriptRoot "SPC_Seed.dwg"
$AcadConsolePath = "C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe"
$AcadGuiPath     = "C:\Program Files\Autodesk\AutoCAD 2026\acad.exe"
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
The two AutoCAD executable paths default to AutoCAD 2026. The Engineering Job
Assistant can override them safely with the `-AcadConsolePath` and
`-AcadGuiPath` PowerShell parameters; paths containing spaces are supported.

5. The Operator Workflow
Once the script starts, it requires minimal input. Watch the PowerShell terminal for color-coded status updates.

The Background Auto-Pilot (Green)
If a part does not contain a bevel flag (K, V, or BEVEL) in its TEXT/MTEXT notes, the terminal will display [+] Clean Part. The script will use accoreconsole.exe to invisibly open the part, inject the layers, swap hashtags to dashes, save the file as a DWG, and move it to a sorted folder (e.g., 250-DH-36). No action is required. (Bevel detection now reads only the drawing's text entities, so stray K/V characters elsewhere in the file no longer force an unnecessary manual gate.)

If a headless job fails or hangs, the script logs it (see _ORCHESTRATOR_LOGS), keeps the original DXF, and moves on to the next part instead of stalling the whole batch. A per-job timeout (default 180s, set via $ConsoleTimeoutSec) guards against a wedged accoreconsole.

The Manual Review Gate (Magenta/Cyan)
If the script detects a bevel flag in the DXF, it will pause the background processing and launch the full graphical AutoCAD application.

The drawing will automatically load with all standard corrections (layer mapping, hashtag replacement) already applied.

Review the bevel notes on the drawing.

CRITICAL: Do not use "Save" or "Save As". When you are finished reviewing the part, simply type FINISH into the AutoCAD command line and press Enter.

The custom FINISH command automatically saves the file to the correct sorted directory and closes only the current drawing tab. AutoCAD itself remains open, so the next beveled part can load without paying the full application startup cost again. The orchestrator watches for the saved DWG and then moves on.

If FINISH is not used, the orchestrator cannot use AutoCAD exiting as its completion signal because AutoCAD is intentionally kept open and may reuse an existing process. It waits up to $BevelReviewTimeoutSec (one hour by default), then marks the part as failed, keeps the original DXF for retry, and continues. This setting is near the top of Master_Orchestrator.ps1 and can be changed before a run.

6. Output and Data Safety
Sorted Output: Processed parts will appear in newly generated folders at the root of your workspace named according to the spreadsheet data (e.g., 500-A-36, 250-DH-36).

Appended Naming: The final .dwg files will automatically append the target folder name and the spreadsheet quantity to the filename (e.g., PartName_250-DH-36_5.dwg).

The Archive Failsafe: The script will never delete your original .dxf files. Once a file is successfully processed and the DWG size is validated (and AutoCAD has fully released its lock on the new DWG), the original raw DXF is moved into a _PROCESSED_DXF_ARCHIVE folder. If a batch fails or requires re-processing, you can always retrieve the pristine original files from the archive.

Diagnostics: Each headless conversion writes its console output to _ORCHESTRATOR_LOGS\<part>.log (and .err.log for errors), so a failed clean part can be diagnosed instead of just reporting "DWG missing." The run ends with a summary line counting Clean / Bevel / Failed parts.

Pre-flight checks: Before doing any work, the script verifies that the seed DWG, ColorToLayer.lsp, accoreconsole.exe, and acad.exe all exist at the configured paths, and aborts with a clear message if any are missing. The parts.csv remains optional -- without it, parts still convert but land in an "Unsorted" folder with a quantity of 1.
