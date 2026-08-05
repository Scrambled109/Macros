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

- The DXF stage launches the Python orchestrator and leaves results and logs in
  the selected job's controlled workspace.
- For the CAD Batch Converter, the assistant supplies the selected DWG folder,
  filtered-DWG folder, output folder, and confirmed plate thickness at run
  time. In SolidWorks choose **Tools > Macro > Run** and select
  `solidworks/cad-batch-converter/Main.RunBatch.swp`.
- AutoBOM changes properties and saves models. Use a recoverable CAD copy and
  review skipped/save results.
- Production comparison runs separately so the Job Assistant window stays
  responsive. Review the Excel or HTML report before completing the stage.

## Finishing a stage

Use **Complete After Review** only after opening the output and reading its log.
Use **Reopen** when a correction is required; the event history is retained.
When promoting approved staged files, same-name conflicts default to **Do Not
Replace**. Choosing replacement creates a timestamped backup first.

If a run is interrupted, inspect `_JOB_ASSISTANT/Logs`, `Working`, and `Staging`
before retrying. Never assume a launch means every file completed.

