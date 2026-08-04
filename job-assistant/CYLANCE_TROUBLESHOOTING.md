# Cylance Script Control troubleshooting for the DXF stage

## What is happening

The packaged Engineering Job Assistant launches the separate
`Engineering DXF Orchestrator.exe` companion. Source-mode development launches
`autocad/dxf-orchestrator/Master_Orchestrator.py` with the active Python
interpreter. Neither path uses PowerShell.

If Cylance stops only `Engineering Job Assistant.exe`, changing the DXF
orchestrator cannot make the assistant itself start: give IT the Cylance event
for that executable and the tested build. If the assistant starts but the DXF
stage does not, identify whether the event names the assistant, the DXF
companion, Python, or AutoCAD before changing configuration.

PowerShell Execution Policy and Cylance are separate controls, and switching the
orchestrator to Python does not override any endpoint-security decision. Do
**not** disable, bypass, or attempt to evade endpoint security.

## Logs and evidence

The assistant's `dxf-*.log` and job audit history record the exact requested
command and argument list, workspace, assistant log path, stage, launched PID,
and actual exit code. Provide IT with:

- the full path of the executable named in the Cylance event;
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
