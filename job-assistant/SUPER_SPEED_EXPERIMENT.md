# Super Speed Boost experiment

This branch stages SolidWorks plate batches on the workstation's local drive
instead of running imports and saves directly inside OneDrive.

## Test

1. Download or switch to `agent/super-speed-boost`.
2. Move aside any contaminated plate staging folder from an earlier failed run.
3. Start the Engineering Job Assistant normally and run one plate thickness.
4. Confirm the log shows:
   - a local speed workspace below `%LOCALAPPDATA%\EngineeringJobAssistant\PlateBatches`;
   - `Verified N current-batch part(s)`;
   - `Macro time: ... s`;
   - `Published N verified part(s)`;
   - removal of the local speed workspace.
5. Review the published parts in the job's normal
   `_JOB_ASSISTANT\Staging\SolidWorks Parts\<group>` folder.

The stable `user` branch is unchanged. This experiment keeps one visible,
reused SolidWorks instance. It does not use multiple instances or hide the
SolidWorks application.

If the macro fails or publication cannot be verified, the local workspace is
preserved for recovery and the runner returns a nonzero exit code.
