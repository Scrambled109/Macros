# Super Speed Boost experiment

This branch stages SolidWorks plate batches on the workstation's local drive
instead of running imports and saves directly inside OneDrive.

## Test

1. Download or switch to `agent/super-speed-boost`.
2. Move aside any contaminated plate staging folder from an earlier failed run.
3. Start the Engineering Job Assistant normally and run one plate thickness.
   Choose the number of parallel SolidWorks instances when prompted. Four is
   the aggressive default; use two first if workstation memory is limited.
4. Confirm the log shows:
   - a local speed workspace below `%LOCALAPPDATA%\EngineeringJobAssistant\PlateBatches`;
   - one unique SolidWorks process ID for every requested worker;
   - `Verified N current-batch part(s)`;
   - `Macro time: ... s`;
   - `Published N verified part(s)`;
   - removal of the local speed workspace.
5. Review the published parts in the job's normal
   `_JOB_ASSISTANT\Staging\SolidWorks Parts\<group>` folder.

The stable `user` branch is unchanged. This experiment starts a dedicated,
hidden SolidWorks COM session for each worker. A start barrier checks that every
worker has a unique SolidWorks process ID before any macro is allowed to run.
Each worker receives its own source, filtered, output, ready, and log files.
Dedicated workers are launched from the SolidWorks executable configured in
Job Assistant Settings and attached by their exact Windows process ID. COM-only
creation requests can resolve back to the already-running SolidWorks session,
so the controller deliberately does not use that approach.
Workers are opened one at a time and held at the start barrier. This avoids
concurrent SolidWorks startup requests collapsing into the same session; the
actual plate macros still run together after every unique session is ready.

The compiled VBA currently attempts to reuse an active AutoCAD session. That is
the biggest experimental risk: SolidWorks sessions are isolated, but their DWG
imports may still contend for AutoCAD. Exact output-set validation prevents a
bad or mixed batch from being published. If the trial hangs, fails, or is not
faster, return to the stable `user` branch.

If the macro fails or publication cannot be verified, the local workspace is
preserved for recovery and the runner returns a nonzero exit code.
