# Engineering Job Assistant

Run `Launch Job Assistant.bat` from the repository root. The assistant creates a
standard job workspace and records progress in `job_manifest.json`.

The assistant does not replace engineering judgment. A launched stage remains
`in_progress` until an operator records a review note and marks it complete.
Warnings require acknowledgement; missing or invalid required inputs block a
launch.

Copy `local_config.example.json` to `local_config.json` and update installed
software paths. The local file is ignored by Git because paths vary by machine.

