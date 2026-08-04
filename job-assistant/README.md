# Engineering Job Assistant — beta

The Engineering Job Assistant is a Windows workflow command center for the normal BOP/BOM → Parts List → prepared cut files → SolidWorks models and assembly → AutoBOM → nesting → reconciliation sequence. It tells an engineer what to select, works on controlled copies where practical, launches repository tools with job-specific paths, and requires a positive review before a stage becomes complete. **Starting a macro is never recorded as engineering success.**

## Beta scope and limitations

This beta provides selected-path job setup, a persistent dashboard and stage guide, BOP conversion, reviewable DXF preparation, existing-tool launches, audit history, comparison summaries, and conflict-safe staged-file promotion. AutoCAD, SolidWorks, the compiled VBA macros, and the PowerShell orchestrator still require the documented Windows/CAD versions and local validation. Manual modeling, assembly, nesting, and CAD review remain engineering work. The assistant does not infer that every DWG is a shape or every DXF is a plate.

The packaged EXE has not been built or CAD-tested on Linux. Build and acceptance-test it on a representative Windows engineering workstation before distribution. The editable CAD batch source now reads the registry/environment values written by the assistant, but editing `.bas` source does not update the compiled `.swp`. Rebuild and test the macro before relying on those values; until then, treat the launch as a guided manual checkpoint.

## Set up or attach a job

Choose **Set Up / Attach Job**. File Explorer asks for three existing folders; names are not assumed:

1. Engineering Process/root folder (often `SA ENGINEERING PROCESS`)
2. 3D Model folder
3. Cut Files folder

Enter and confirm the authoritative job number and revision. The assistant may suggest leading digits from a surrounding folder, but it never overrides what the engineer types. The same setup flow attaches to an existing job or initializes a new, already-created job layout. It does **not** create a fixed production tree and does not require Drawings, LOA, or forming folders.

If the selected root already contains an assistant manifest, setup opens it rather than offering to replace it. Audit history is never discarded by the setup workflow.

Part Checking, Nesting, and Forming are optional. Select them later with **Set Optional Folder** only when relevant. Every selected path is restored from `_JOB_ASSISTANT/job_manifest.json`.

## Dashboard and stage guide

The dashboard shows job/revision, required and optional paths, overall progress, recommended next action, outstanding warnings, color-coded stage states, artifact counts, and the five most recently recorded files. Double-click a recent file to open it, or use the direct Staging and Logs buttons. Select a stage to see:

- required input and what to select;
- what the assistant copies or changes;
- which external tool opens;
- the post-run review;
- staging and log locations; and
- how an ordinary warning can be overridden.

Statuses distinguish Not Started, Ready, In Progress, Needs Review, Complete, and warnings. **Complete After Review** requires a review note. **Reopen** preserves the event history. Events and warning overrides contain the Windows username and UTC time.

The toolbar also shows active external automation and the job number that
launched it. You may switch to another job while DXF automation runs: its final
status remains tied to the launching job, and a completion notification names
that job. If you close the assistant while automation is active, it warns that
AutoCAD will continue but status monitoring and notifications will stop.

## Assistant-owned data and recovery

The application creates this isolated area under the selected Engineering Process folder:

```text
_JOB_ASSISTANT/
  Source Copies/   received inputs copied for automation
  Working/         disposable controlled workspaces
  Staging/         generated files awaiting engineering review
  Logs/            process logs
  Backups/         replaced production files, never deleted
  Reports/         comparison reports
  job_manifest.json
```

Manifest saves use a same-directory temporary file, flush it, and atomically replace the manifest. Older version-1 manifests are migrated to selected-path schema; a manifest newer than the application gives a compatibility message instead of being guessed at. If a run is interrupted, inspect Logs, Working, and Staging. Received originals are never moved or deleted by DXF preparation.

## BOP/BOM to Parts List

Select **Create Parts List from BOP/BOM** and **Start This Step**:

1. Select the received BOP/BOM workbook.
2. On first use, select the shared standard Parts List template. Its UNC/network path is remembered only in per-user settings.
3. Confirm or change the recommended output `<JOB_NUMBER>_PARTS_LIST.xlsx` in the Engineering Process folder.
4. The assistant copies the source into `Source Copies/BOP-BOM` and runs the existing BOM converter on that copy.
5. Open and review the result; only then complete the stage.

The converter's non-interactive API uses its existing conservative mapping suggestions (`DS,DV` prefixes). If the source needs custom column mapping, use the converter GUI and record its result rather than accepting a questionable automatic output.

## DXF and shape-sketch review

Select the incoming single folder. The assistant reports DXF and DWG counts, initially selects `.dxf` files, and initially excludes `.dwg` files as likely manual shape sketches. Choose **Review individual files** to include/exclude every candidate. Classification is a proposal, not a claim.

Confirmed files are copied—not moved—to a timestamped `Working/DXF Orchestrator/.../001` folder. If two preparations start within the same second, a numeric suffix prevents one run from colliding with another. The `001` adapter satisfies the orchestrator's numbered-folder input without changing received files. The orchestrator runs with that timestamped workspace as its working directory, so its archive, sorted DWGs, and `_ORCHESTRATOR_LOGS` remain controlled assistant data. Configured AutoCAD GUI/console paths are passed as quoted PowerShell arguments; blank settings retain the orchestrator's documented AutoCAD 2026 defaults. For Cylance Script Control failures, follow [`CYLANCE_TROUBLESHOOTING.md`](CYLANCE_TROUBLESHOOTING.md). For bevel reviews follow the orchestrator instruction to type `FINISH`; never assume a launch completed. Record and review outputs/logs before completion.

## SolidWorks and AutoBOM

The plate stage asks for the reviewed folder that directly contains the prepared DWGs. It sets `MACROS_SOURCE_FOLDER`, `MACROS_FILTERED_FOLDER`, and `MACROS_OUTPUT_FOLDER` in the process environment and the per-user registry keys under `HKCU\Software\VB and VBA Program Settings\EngineeringMacros\CadBatch`. Updated `Config.bas` source reads environment first, registry second, normalizes trailing slashes, and fails clearly when no job folders are configured. It directs intended output to assistant staging. A compiled `.swp` cannot be proven to contain updated `.bas` source merely because it opened; rebuild/test it in the documented CAD environment and inspect `BatchLog.txt`.

AutoBOM is high impact: it updates properties and saves models. Make a recoverable CAD copy and review every skipped/save result. The assistant records only “launch initiated” and leaves the stage at Needs Review.

## Comparison summary

The comparison wizard selects the Parts List CSV and SolidWorks/model CSV, reuses the selected Nesting folder, writes beneath Assistant Reports, and invokes the existing comparison CLI in headless mode (no second process-owned completion dialog). The comparison tool writes a versioned `comparison_summary.json` beside its Excel and HTML outputs, including input paths, detailed counts, outcome, and report locations. After a successful exit, the assistant requires and reads that stable contract—even when the tool creates its normal timestamped run folder—and summarizes:

- No discrepancies found
- Review recommended
- Action required
- error and warning counts

Excel and HTML report paths are retained in the manifest. Open the reports folder or record the reports as artifacts. A started process, missing output, or nonzero exit is never shown as a pass.

## Copy approved files to production

**Copy Approved Files to Production** is optional; engineers may move files manually. The assistant accepts promotion sources only from its Staging folder and previews every source and destination. New files default to Copy. Same-name conflicts default to **Do Not Replace**. For each conflict an engineer may choose **Back Up Existing and Replace**. Existing files are copied to a timestamped, collision-safe Backups folder before the staged file is copied. The event history records copied, skipped, replaced, backup, and failed operations honestly; a partial failure does not roll back successful independent copies.

Promotion uses one review table rather than a sequence of conflict dialogs. Select one or more rows and assign Copy New File, Do Not Replace, or Back Up + Replace. Results remain visible per row after execution. Every run writes a `promotion-*.json` recovery report in Assistant Logs with user, job, revision, summary counts, paths, actions, backups, errors, and outcomes.

## Per-user settings

**Settings** stores machine-specific values outside the repository (normally `%LOCALAPPDATA%\EngineeringJobAssistant\settings.json`):

- shared Macros repository, including UNC paths and spaces;
- standard Parts List template;
- default jobs parent;
- AutoCAD GUI/console and SolidWorks executable paths.

Use **Browse** rather than typing paths. Nothing machine-specific is committed. Executable detection is intentionally conservative in this beta; browse to nonstandard installs.

## Run from source (developers only)

End users should not install Python. For development:

```powershell
cd "X:\path with spaces\Macros"
py -m pip install -r data-tools\bom-converter\requirements.txt
py job-assistant\job_assistant.py
```

Core tests are platform independent:

```powershell
py -m unittest discover -s job-assistant\tests -v
```

A GUI display and Windows CAD applications are required for end-to-end GUI/CAD acceptance.

## Build the standalone Windows beta

On a clean Windows build machine, run `job-assistant\build_windows.bat`. It installs PyInstaller for that build environment and creates a complete three-EXE distribution in `job-assistant\dist`: the Job Assistant plus console companions for BOM conversion and production comparison. Keep all three EXEs together; packaged users do not need Python, and the assistant reports a missing companion rather than attempting to run a `.py` file. Validate the distribution, network-drive/UNC access, Office dependencies, and all CAD versions/macros before sharing it. `job-assistant\Launch Job Assistant.bat` requires the packaged assistant. Developers can explicitly request source mode with `Launch Job Assistant.bat --source`. This is a portable beta distribution, not an enterprise installer or auto-updater.

See [`../WORKFLOW.md`](../WORKFLOW.md) for the domain checkpoints and the component documentation for each tool's recovery procedure.

For a structured Windows pilot and an exact failure-report template, use [`WINDOWS_TEST_CHECKLIST.md`](WINDOWS_TEST_CHECKLIST.md).
