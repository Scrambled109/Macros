# SolidWorks macros

## Folder policy

- `drawing-automation/`, `auto-bom/`, `cad-batch-converter/`, and `utilities/`
  contain the current runnable `.swp` artifacts.
- `cutfile-exporter/` is the runnable Python/SolidWorks tool for verified layered
  DXF output; its launcher installs only its two third-party packages.
- A `reference/` directory contains readable `.bas` source for review. These
  files are not executed by SolidWorks and do not patch the `.swp` binaries.
- `reference-extracted-source/` is a reproducible snapshot recovered from the
  compiled projects.
- `legacy/` contains files that are explicitly experimental, unsafe, or
  superseded/version-named. Nothing was deleted during the reorganization.

Always test compiled macros on disposable documents. To carry a source fix into
production, open the matching `.swp` in the SolidWorks VBA editor, apply the
change, compile the project, and save a newly validated binary.
