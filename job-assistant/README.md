# Engineering Job Assistant

The Job Assistant is the recommended starting point for a complete production
job. It keeps job-specific paths, controlled working copies, staging files,
logs, reports, review notes, and backups together. Starting an external tool is
never treated as proof that the engineering stage succeeded.

## Run it

Install the branch requirements once from the repository root:

```powershell
py -m pip install -r requirements.txt
```

Then double-click `Launch Job Assistant.bat`, or run:

```powershell
py job-assistant/job_assistant.py
```

This branch intentionally has no packaged EXE. Keep `job_assistant.py` and
`job_core.py` together, and keep the surrounding repository folder structure
unchanged so the assistant can find the other tools.

## First setup

Choose **Set Up / Attach Job** and select the existing:

1. Engineering Process/root folder
2. 3D Model folder
3. Cut Files folder

Confirm the job number and revision. Optional Part Checking, Nesting, and
Forming folders can be selected later. Machine-specific paths are stored in
your local user settings, not in the repository.

The assistant creates `_JOB_ASSISTANT` inside the selected Engineering Process
folder with Source Copies, Working, Staging, Logs, Backups, Reports, and a job
manifest. Received originals are copied into controlled work areas; they are
not silently deleted.

## CAD stages

- The DXF stage runs the Python orchestrator against controlled copies. AutoCAD
  Core Console is detected automatically, preferring 2026 and falling back to
  2025 or another installed version. Bevel drawings stay headless and receive a
  `(B)` filename suffix.
- For plate models, the assistant prepares source, filtered-DWG, output, and
  extrusion-depth settings, starts SolidWorks when configured, and opens the
  macro folder. Wait for SolidWorks to finish loading, then manually run
  `solidworks/cad-batch-converter/Main.RunBatch.swp`.
- The modified Parts List property step launches
  `solidworks/auto-bom/AutoBOMProperties.swp`. It uses the active Excel
  workbook or asks for one, shows mapping dropdowns, and processes selected
  components—or the whole assembly when none are selected.
- Production comparison runs in the background so the Job Assistant remains
  responsive. Review the Excel or HTML report before completing the stage.
- **Open Results** opens the selected step's actual output: the current DXF run,
  the current SolidWorks staging folder, the modified model folder, or the
  comparison report folder. Logs remain separately available from **Tools >
  Open Logs**.
- The AutoCAD worker setting accepts any positive whole number. Two remains the
  recommended default, but the assistant does not impose an upper limit.

## Finishing a stage

Use **Complete After Review** only after opening the output and reading its log.
Use **Reopen** when a correction is required; the event history is retained.
When promoting approved staged files, same-name conflicts default to **Do Not
Replace**. Choosing replacement creates a timestamped backup first.

If a run is interrupted, inspect `_JOB_ASSISTANT/Logs`, `Working`, and `Staging`
before retrying. Never assume a launch means every file completed.

