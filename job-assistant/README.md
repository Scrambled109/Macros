# Engineering Job Assistant

The Engineering Job Assistant is a Windows dashboard for the work this
repository can actually perform or launch. It keeps job-specific paths and an
audit record, prepares controlled copies, and requires a positive review before
an automated step becomes complete. Starting a macro is never recorded as
engineering success.

Manual shape/specialized modeling, assembly review, nesting, discrepancy
resolution, and final approval remain in [`../WORKFLOW.md`](../WORKFLOW.md), but
they are deliberately not dashboard steps because the assistant does not do
that work.

## The five assistant steps

1. Create a Parts List from the BOP/BOM.
2. Review and prepare cut files with the DXF orchestrator.
3. Guide automatic SolidWorks plate-model creation.
4. Apply modified Parts List properties and record the review.
5. Compare Parts List, SolidWorks, and nesting production data.

Existing manifests are migrated without losing history. Former manual stages
are retained under `legacy_stages` in `job_manifest.json`; they are simply no
longer shown as assistant-owned work.

## Start the assistant

Install the repository dependencies once, then use the launcher:

```powershell
cd "X:\path with spaces\Macros"
py -m pip install -r requirements.txt
job-assistant\Launch Job Assistant.bat
```

The launcher and Python source are the supported shared-drive workflow. A
packaged build remains available for controlled testing, but it is not required
for normal use.

## Set up or attach a job

Choose **Job > Set Up / Attach Job**, then select three existing folders:

1. Engineering Process/root folder (often `SA ENGINEERING PROCESS`)
2. 3D Model folder
3. Cut Files folder

Confirm the authoritative job number and revision. Part Checking, Nesting, and
Forming are optional job-folder references and can be selected later with
**Job Folders…**. If the selected root already contains a manifest, the
assistant opens it rather than replacing its audit history.

## Dashboard design

The dark dashboard uses Steel America's blue as its action color and includes
the Steel America banner logo. The default view intentionally shows only the job, recommended next action,
overall progress, five workflow statuses, and one selected-step action. Use the
selected step's **More** menu for readiness, folders, file recording,
completion, reopening, technical details, and **Move Completed Outputs**.
**View > Job Details** contains production paths, assistant workspace paths,
warnings, and recent files. **Tools** contains folder/report shortcuts.

Background-process text is hidden while nothing is running and appears beside
the job heading only while AutoCAD, SolidWorks, or the comparison tool is
active.

Background completion is posted in the green or amber dashboard banner with an
**Open Result** button. The button targets the generated output—not the process
log. Successful required-review steps open their result automatically and ask
whether the operator wants to mark the step complete; no typed note is required
for this completion path.

Statuses distinguish Not Started, Ready, In Progress, Needs Review, Complete,
and Warning. The manual **Mark Complete** action remains available when review
is performed later. **Reopen Step** preserves event history. Events and warning
overrides include the Windows username and UTC time.

## Assistant-owned data

The assistant creates an isolated area below the selected Engineering Process
folder:

```text
_JOB_ASSISTANT/
  Source Copies/   received inputs copied for automation
  Working/         controlled DXF workspaces
  Staging/         generated SolidWorks plate models
  Logs/            process and move reports
  Backups/         production files replaced by a move
  Reports/         comparison reports
  job_manifest.json
```

Manifest writes are atomic. Received DXFs are copied into a timestamped working
run; the received originals are not changed.

## Parts List conversion

Select the received BOP/BOM, the standard Parts List template on first use, and
the output workbook. The template path is remembered per user. The converter
discovers template destinations by normalized header name, so inserted or
reordered columns do not shift data into the wrong field. Ambiguous duplicate
headers stop with a readable error.

The converter also extends the template's existing Excel table through the
last generated record. This preserves its alternating-row table style beyond
the original blank-row limit (row 187 in the current standard template) while
retaining any additional template columns.

After workbook review, completing the step exports
`_JOB_ASSISTANT/Source Copies/Parts List.csv` with the DXF orchestrator's exact
`PartNumber,Quantity,Thickness,Material` columns and plate rows only.

## Cut-file preparation

The assistant proposes DXFs and initially excludes DWGs as likely manual shape
sketches. Confirmed files are copied into
`Working/DXF Orchestrator/<run>/001`, together with `Parts List.csv`.

AutoCAD Core Console is detected automatically, preferring AutoCAD 2026 and
falling back to AutoCAD 2025 or another installed version. An explicit Settings
path still overrides detection.

All drawings use two AutoCAD Core Console processes by default. Set **Parallel
AutoCAD processes** from 1–4 in Settings; use 1 if licensing or workstation
stability requires serial processing. Every process receives a unique AutoCAD
script and log. Inputs that resolve to the same output DWG are rejected before
AutoCAD runs instead of racing or silently overwriting one another.

Detected bevel drawings use the same headless process and receive the exact
suffix `(B)` before `.dwg`. Bevel callout text and all leader arrows are moved
to `PLOT`, while ordinary pin-stamp entities retain their normal layers.
Graphical AutoCAD no longer opens for bevel review.

## SolidWorks plate models and modified Parts List properties

For one reviewed material/thickness folder, the assistant asks for the
thickness and configures the source, filtered, output, and extrusion-depth
settings used by `Main.RunBatch.swp`. The folder picker starts in the selected
production **Cut Files** folder, so moved orchestrator outputs remain easy to
find.

The assistant then launches `run_macro.py`. The runner checks for an active
SolidWorks COM session first and reuses it; only when none exists does it start
one instance. It waits until the VBA host can inspect the compiled `.swp`, runs
the macro in the background, restores the same SolidWorks window for review,
and never opens or rewrites the batch-converter source folder.

SolidWorks plate work remains serial because competing controllers in one
session can activate or close the wrong document. When one thickness folder
finishes, its output opens and the assistant asks whether to select the next
folder. A new thickness only changes the registry/environment value read when
the next macro begins; it does not open another SolidWorks instance.

The compiled `.swp` remains the production macro. Loose `.bas` changes, such as
the added pin-stamp sketch display suppression, require importing the updated
modules and rebuilding the `.swp` on a SolidWorks workstation before those VBA
source improvements take effect.

Plate parts are staged below `Staging/SolidWorks Parts/<group>`. The property
step launches `solidworks/utilities/(MOD)2(SECONDARY).swp`. It uses the active
Excel workbook or asks for one, then maps spreadsheet columns to SolidWorks
properties. This is a high-impact step because it saves part files. Work from a
recoverable model copy and account for every unmatched, skipped, or failed file.

## Move completed outputs

After positively completing the cut-file or plate-model step, choose **More >
Move Completed Outputs** beside **Start Step**. One review table shows every
automatic destination:

- sorted cut DWGs are merged into `Cut Files/<material-thickness>/`;
- staged plate `.SLDPRT` files move into the selected `3D Model` folder.

Same-name production conflicts are backed up automatically before replacement.
Identical production files are recognized without creating a needless backup.
Each source is removed only after a verified temporary copy is atomically
placed at the destination. Every run writes a detailed
`completed-output-move-*.json` recovery report and records results in the
manifest.

## Production comparison

The assistant launches comparison asynchronously and remains usable. Parts List
`#` separators are normalized to the `-` used by nesting/SolidWorks, and known
SolidWorks configuration suffixes such as `(Default<As Machined>)` are removed
before matching. The process runs unbuffered so progress reaches the log
immediately, and `--no-open` never pauses for console input.

Results are written below Assistant Reports. A nonzero exit, missing summary,
or actionable discrepancy is never shown as a pass.

## Per-user settings

Settings are stored outside the repository, normally at
`%LOCALAPPDATA%\EngineeringJobAssistant\settings.json`:

- shared Macros repository;
- standard Parts List template;
- default jobs parent;
- AutoCAD Core Console and SolidWorks executable paths;
- parallel AutoCAD processes (default 2, maximum 4).

## Tests and Windows validation

Platform-independent tests:

```powershell
py -m unittest discover -s job-assistant\tests -v
py -m unittest discover -s data-tools\bom-converter\tests -v
py -m unittest discover -s data-tools\production-comparison\tests -v
py autocad\dxf-orchestrator\test_master_orchestrator.py
py -m unittest discover -s solidworks\cad-batch-converter\tests -v
```

Run [`WINDOWS_TEST_CHECKLIST.md`](WINDOWS_TEST_CHECKLIST.md) on a licensed
AutoCAD/SolidWorks workstation before production rollout. The overnight,
checkpointed automation concept is intentionally a separate second phase.
