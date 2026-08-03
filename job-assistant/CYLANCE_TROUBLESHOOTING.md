# Cylance Script Control troubleshooting for the DXF stage

## What is happening

The Engineering Job Assistant starts PowerShell to execute
`autocad/dxf-orchestrator/Master_Orchestrator.ps1`. If Cylance Script Control
blocks PowerShell, the block happens before the orchestrator can contact
AutoCAD. Changing AutoCAD executable or macro paths therefore cannot solve this
pre-execution security decision.

PowerShell Execution Policy and Cylance Script Control are separate controls.
The command's `-ExecutionPolicy Bypass` option cannot override Cylance or any
other endpoint-security decision. Do **not** disable, bypass, or attempt to evade
endpoint security.

## Logs and evidence

The assistant's `dxf-*.log` and job audit history record the exact requested
command and argument list, workspace, assistant log path, stage, launched PID,
and actual exit code. Provide IT with:

- the full `Master_Orchestrator.ps1` path;
- the exact command at the top of `dxf-*.log`;
- the Cylance event time and identifier;
- the workstation name and username; and
- the tested repository commit.

## IT-approved resolution options

Ask IT/security to choose and validate an organizationally approved option,
such as trusted-path or hash allow-listing, signing with an organizational
certificate, managed deployment, or an approved signed wrapper.

Until approval is in place, engineers may prepare controlled source copies, but
must keep the DXF stage incomplete and use the existing approved manual CAD
workflow.
