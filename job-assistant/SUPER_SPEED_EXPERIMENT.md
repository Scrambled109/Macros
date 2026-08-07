# Local SolidWorks plate pipeline

This branch runs each plate-thickness batch through one persistent SolidWorks
session. It no longer attempts to create or automate several SolidWorks
instances under one Windows desktop session.

## Execution model

1. The Job Assistant copies the reviewed DWGs to a unique workspace below
   `%LOCALAPPDATA%\EngineeringJobAssistant\PlateBatches`.
2. One background Python runner reuses the active SolidWorks session, or starts
   one configured SolidWorks executable when none is active.
3. The compiled batch macro processes the entire folder through that single
   application owner. Source, filtered, and output folders remain on the local
   drive during the run.
4. The runner accepts only parts created or changed by this batch, normalizes
   their filenames, and compares the complete result set with the selected DWGs.
5. Verified parts and `BatchLog.txt` are copied atomically to the job's normal
   staging folder.
6. Only after verified publication does the runner remove the marked local
   workspace.

A failed macro, incomplete output set, unrelated staging file, rename conflict,
or publication error preserves the local workspace and returns a nonzero exit
code. Nothing partial is intentionally accepted as a completed plate batch.

## Why the multi-instance experiment was removed

SolidWorks' COM ProgID and Running Object Table behavior did not provide a
reliable one-client-to-one-process mapping under the same Windows user session.
Separate Python workers repeatedly resolved to the same SolidWorks process, and
the compiled VBA also connects through the active application registration.
AutoCAD was another shared COM dependency. Opening several windows therefore
did not establish several safely isolated automation owners.

Real multi-instance SolidWorks automation requires isolation outside this
desktop session, such as separate Windows sessions, virtual machines, or
separate workstations. That complexity is outside this operator tool.

## Test

1. Download or switch to `agent/super-speed-boost`.
2. Close abandoned SolidWorks processes from earlier parallel trials.
3. Run one reviewed plate-thickness folder.
4. Confirm the assistant does not ask for a worker count.
5. Confirm the log contains:
   - `SolidWorks mode: one persistent local session`;
   - the local source/output workspace paths;
   - `Verified N current-batch part(s)`;
   - `Published N verified part(s)`;
   - removal of the verified local workspace.
6. Review the published parts in
   `_JOB_ASSISTANT\Staging\SolidWorks Parts\<group>`.

The stable `user` branch remains unchanged while this local pipeline is tested.
