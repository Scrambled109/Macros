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
4. Launch AutoBOM and record the property review.
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

Choose **Set Up / Attach Job**, then select three existing folders:

1. Engineering Process/root folder (often `SA ENGINEERING PROCESS`)
2. 3D Model folder
3. Cut Files folder

Confirm the authoritative job number and revision. Part Checking, Nesting, and
Forming are optional job-folder references and can be selected later with
**Job Folders…**. If the selected root already contains a manifest, the
assistant opens it rather than replacing its audit history.

## Dashboard design

The default view is intentionally concise: current status, what is needed,
what to do, the review requirement, and readiness checks. **Technical Details**
opens tool paths, changes, logs, warnings, and corrective guidance for an
experienced operator. The toolbar tracks background processes by job number.

Background AutoCAD/comparison completion is posted in the green or amber
dashboard banner with an **Open Result** button. It does not use a modal
completion dialog, which prevents a hidden Windows dialog from disabling the
entire Tk dashboard and making it appear frozen.

Statuses distinguish Not Started, Ready, In Progress, Needs Review, Complete,
and Warning. **Mark Complete** requires a review note. **Reopen Step** preserves
event history. Events and warning overrides include the Windows username and
UTC time.

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

Clean drawings use two AutoCAD Core Console workers by default. Set **AutoCAD
clean-file workers** from 1–4 in Settings; use 1 if licensing or workstation
stability requires serial processing. Every worker receives a unique AutoCAD
script and log. Inputs that resolve to the same output DWG are rejected before
AutoCAD runs instead of racing or silently overwriting one another.

Bevel drawings remain in one interactive AutoCAD session. Review each tab and
type `SPCFINISH`; it saves the sorted DWG and closes that drawing, not AutoCAD.

## SolidWorks plate models and AutoBOM

For one reviewed material/thickness folder, the assistant asks for the
thickness, configures source/filtered/output paths and extrusion depth, and
opens exact **Tools > Macro > Run** instructions. The operator still runs and
reviews `Main.RunBatch.swp`; loose `.bas` source does not update a compiled
`.swp`.

Plate parts are staged below `Staging/SolidWorks Parts/<group>`. AutoBOM is a
high-impact step because it updates properties and saves models. Work from a
recoverable model copy and account for every skipped or failed file.

## Move completed outputs

After positively completing the cut-file or plate-model step, choose **Move
Completed Outputs**. One review table shows every automatic destination:

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
- AutoCAD GUI/console and SolidWorks executable paths;
- AutoCAD clean-file workers (default 2, maximum 4).

## Tests and Windows validation

Platform-independent tests:

```powershell
py -m unittest discover -s job-assistant\tests -v
py -m unittest discover -s data-tools\bom-converter\tests -v
py -m unittest discover -s data-tools\production-comparison\tests -v
py autocad\dxf-orchestrator\test_master_orchestrator.py
```

Run [`WINDOWS_TEST_CHECKLIST.md`](WINDOWS_TEST_CHECKLIST.md) on a licensed
AutoCAD/SolidWorks workstation before production rollout. The overnight,
checkpointed automation concept is intentionally a separate second phase.
