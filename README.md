# Engineering Macros — User Branch

This branch contains only the files needed to run the engineering tools. It has
no Windows EXE distributions, build files, tests, editable VBA `.bas` source, or
developer-only reference material. Development stays on `main`.

Always test on copies of job files first. AutoCAD and SolidWorks automation can
modify or save files without asking again.

## One-time Windows setup

1. On GitHub, switch to the **user** branch, select **Code > Download ZIP**, and
   extract the ZIP to a normal local folder.
2. Install 64-bit Python 3.11 or newer. During installation, enable **Add Python
   to PATH**.
3. Open PowerShell in the extracted folder and run:

```powershell
py -m pip install -r requirements.txt
```

That installs every third-party Python package used anywhere in this branch:

- `openpyxl` — reads and writes Excel workbooks
- `pandas` — used by the BOP/BOM converter
- `customtkinter` — provides the EPL converter desktop interface

Everything else used by the Python scripts is included with Python. On a normal
Windows Python installation, `tkinter` is included too.

Python alone is enough for the spreadsheet comparison logic, but CAD workflows
still require the matching desktop software: AutoCAD 2026 and/or SolidWorks
2025. Excel is useful for reviewing the generated workbooks.

## Start here

| Task | What to run | Extra software |
|---|---|---|
| Guided full-job workflow | Double-click `job-assistant/Launch Job Assistant.bat` | AutoCAD/SolidWorks for CAD stages |
| Compare production data | Double-click `data-tools/production-comparison/compare_production_parts.py` | None beyond Python |
| Convert a BOP/BOM to a Parts List | Double-click `data-tools/bom-converter/bom_converter.py` | Python packages above |
| Convert EPLs using required BOP scope | Double-click `data-tools/epl_converter/main.py` | Python packages above |
| Prepare and sort DXF files | Run `autocad/dxf-orchestrator/Master_Orchestrator.py` | AutoCAD 2026 |
| Run a SolidWorks macro | SolidWorks **Tools > Macro > Run**, then select a `.swp` | SolidWorks 2025 |

If you are processing a complete job, use the Job Assistant. It supplies the
job-specific folders and extrusion thickness to the CAD Batch Converter and
keeps working files, logs, reports, and backups together.

## Important file rules

- Keep each tool's supporting JSON, LISP, DWG, and Python module files in their
  existing folders.
- Do not copy only `main.py` from the EPL converter; its `epl_converter` folder
  and `material_translations.json` are required.
- Do not move `SPC_Seed.dwg` or `ColortoLayer.lsp` away from the DXF
  Orchestrator.
- The `.swp` files are the runnable SolidWorks macros. Loose `.bas` files are
  developer source and are intentionally not present on this branch.
- Generated reports, job drawings, workbooks, logs, and processed files should
  remain outside this repository.

For the normal order of operations, read [WORKFLOW.md](WORKFLOW.md). Each major
tool also has a README in its own folder.

## Common problems

- **`py` is not recognized:** reinstall Python and enable **Add Python to PATH**.
- **A Python import is missing:** from the repository root, rerun
  `py -m pip install -r requirements.txt`.
- **PowerShell blocks a script:** use the exact command in the tool's README;
  do not permanently lower Windows security settings.
- **AutoCAD or SolidWorks is not found:** open Job Assistant settings and browse
  to the installed executable.
- **A CAD run fails partway through:** keep the originals, read the generated
  log, and retry only failed parts after correcting the cause.

