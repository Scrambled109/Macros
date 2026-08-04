# Engineering Job Assistant — Windows beta test checklist

Thank you for testing. You do not need to provide a complete production job.
Use a disposable copy with fake or previously completed data. Do not send
controlled, customer-sensitive, or export-controlled files.

## Before testing

Record these facts in your reply:

- Windows version:
- AutoCAD version and installed GUI path:
- `accoreconsole.exe` path:
- SolidWorks version:
- Excel version:
- Macros repository path (mapped drive or UNC; redact server names if needed):
- Job path type (local, mapped drive, or UNC):
- Did you rebuild `Main.RunBatch.swp` from the updated `.bas` modules? Yes/No
- EXE build commit shown by your test package:

Create a disposable job with:

- an Engineering Process/root folder;
- separate 3D Model and Cut Files folders, with at least one folder name that
  differs from the usual spelling;
- a small BOP/BOM copy and a blank Parts List template;
- two ordinary DXFs, one bevel DXF if available, one DWG shape sketch, and one
  unrelated file;
- small Parts List, SolidWorks, and nesting comparison exports if available.

## 1. Package and startup

1. Run `job-assistant\build_windows.bat` on the build workstation.
2. Confirm all three EXEs and the `_internal` runtime directory are together in
   `job-assistant\dist\Engineering Job Assistant`. Copy the entire directory,
   not only the GUI EXE.
3. Run `Launch Job Assistant.bat` without a developer Python terminal.
4. Open Settings and browse to the shared repository, Parts List template,
   AutoCAD GUI, and AutoCAD console.
5. Close and reopen Settings. Confirm all paths remain correct.

Report:

- PASS/FAIL for each step;
- exact error text or a screenshot for a failure;
- whether any console window appeared unexpectedly;
- whether Windows or endpoint security blocked, removed, or quarantined a file
  (follow [`EXE_TROUBLESHOOTING.md`](EXE_TROUBLESHOOTING.md));
- `%LOCALAPPDATA%\EngineeringJobAssistant\settings.json` with private server
  names redacted, if settings did not persist.

## 2. Job setup and dashboard

1. Attach the disposable job and select the three required folders.
2. Confirm the suggested job number, then deliberately edit it before saving.
3. Confirm the dashboard shows job/revision, all paths, progress, next action,
   warnings, stage colors, and recently recorded files.
4. Close and reopen `_JOB_ASSISTANT\job_manifest.json` through Open Job.
5. Run Set Up / Attach against the same root again. It must open the existing
   manifest and retain history; it must not offer to replace it.
6. Double-click a recent file and confirm Windows opens it.
7. Test Open Staging and Open Logs.

Report any clipped text, confusing status/color, wrong next action, path that
is hard to read, or control that requires unnecessary typing.

## 3. BOP/BOM and Parts List

1. Select the disposable BOP/BOM.
2. Confirm the source copy appears under `_JOB_ASSISTANT\Source Copies\BOP-BOM`.
3. Confirm the recommended output is `<JOB_NUMBER>_PARTS_LIST.xlsx` in the
   Engineering Process folder; change the destination once to test browsing.
4. Run conversion and inspect `_JOB_ASSISTANT\Logs\bom_converter.log`.
5. Confirm a failed conversion remains In Progress/Needs Review and records an
   event rather than claiming completion.
6. Inspect and record the workbook, then complete the stage with a review note.

Report the source workbook's sheet/header layout in words if conversion fails;
do not send sensitive workbook contents.

## 4. DXF review and orchestrator

For PowerShell blocks, use [`CYLANCE_TROUBLESHOOTING.md`](CYLANCE_TROUBLESHOOTING.md); do not disable or evade endpoint security.

1. Select the mixed incoming folder.
2. Confirm DXFs start selected, the DWG starts excluded, and the unrelated file
   is not proposed.
3. Open individual review, exclude one DXF, and confirm the selection.
4. Confirm originals are unchanged and only selected files appear in `001`.
5. Prepare twice quickly; confirm the second run gets a numeric suffix.
6. Launch the orchestrator. Verify it uses the configured AutoCAD paths.
7. For a bevel drawing, use `FINISH` as documented.
8. Inspect the timestamped working folder, `_ORCHESTRATOR_LOGS`, assistant DXF
   log, sorted outputs, and archive behavior.

Report the exact last 30 lines of the relevant log for a failure and whether
AutoCAD opened, timed out, or produced a DWG.

## 5. SolidWorks plate macro and AutoBOM

Only use disposable CAD copies.

1. Rebuild the SWP from the updated VBA modules if this has not been done.
2. Select the reviewed folder that directly contains prepared DWGs.
3. Confirm filtered DWGs go to Assistant Working and SLDPRTs go to Assistant
   Staging, never an old compiled-in job path.
4. Inspect `BatchLog.txt`, model geometry, thickness, markings, and failures.
5. Confirm launching leaves the stage at Needs Review.
6. Run AutoBOM on disposable parts and confirm it also requires review.

Report the VBA compile error with module and line if rebuilding fails. For a
run failure, provide `BatchLog.txt` with customer/part names redacted.

## 6. Comparison

1. Select the three disposable comparison inputs.
2. Confirm no second comparison completion dialog blocks the assistant.
3. Confirm the timestamped report folder contains `comparison_summary.json`,
   Excel, and HTML reports.
4. Confirm the dashboard summary agrees with the reports.
5. Test Open Comparison Report.

Provide `comparison_summary.json` and the last 30 lines of `comparison.log` if
the dashboard result is wrong. These should be much smaller than full exports.

## 7. Promotion table and recovery

1. Put three harmless files in Assistant Staging: one new name, one that
   conflicts with production, and one second new name.
2. Open Copy Approved Files to Production and verify every source, destination,
   conflict, action, and result is visible in one table.
3. Set one new file to Copy, the conflict to Do Not Replace, and the other new
   file to Do Not Replace. Execute and verify results.
4. Reopen the table, set the conflict to Back Up Existing + Replace, execute,
   and verify both production and timestamped backup contents.
5. Attempt to select a source outside Staging; it must fail safely.
6. Open Assistant Logs and locate `promotion-*.json`.

Return the promotion JSON report for any unexpected result. It records source,
destination, chosen action, backup, status, error, Windows user, revision, and
summary counts.

## What to send back

For each failure, send one block:

```text
Test section and step:
Expected:
Actual:
Exact error text:
Did the original/source file change? Yes/No/Unknown
Relevant log path:
Last 30 log lines (redacted):
Screenshot filename, if useful:
Can you reproduce it? Always/Sometimes/Once
```

Also send `_JOB_ASSISTANT\job_manifest.json` and the relevant small JSON/log
files when permitted, with customer names, server names, and part numbers
redacted consistently. Do not edit timestamps, statuses, event types, exit
codes, or folder structure because those fields are needed for diagnosis.
