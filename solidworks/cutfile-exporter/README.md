# SolidWorks Layered Cut-File Exporter

Creates production-review DXFs from flat `.SLDPRT` files by driving the installed
SolidWorks application. It exports the largest planar face using SolidWorks'
native face exporter, verifies every cut entity belongs to a closed loop, and
assigns:

- outer perimeter → `CUT - OUTSIDE STRAIGHT`
- holes/internal loops → `CUT - INSIDE STRAIGHT`
- every non-construction entity and rendered text edge in the sketch named
  `CUTFILE MARKING` → `PIN STAMP TEXT`

The tool fails a part rather than releasing a DXF when SolidWorks and the DXF
disagree about loop counts, the cut geometry is open/branching, units cannot be
verified, the part is multibody, or the marking sketch is off the exported face.

## One-time setup

Use Windows with SolidWorks installed. Double-click `Launch Cut File
Exporter.bat`; it installs `ezdxf` and `pywin32` if needed. Or install manually:

```powershell
py -m pip install -r requirements.txt
```

## Prepare each SolidWorks part

1. The part must contain exactly one visible solid body.
2. The intended cut shape must be represented by its largest planar face.
3. Create an ordinary 2D sketch on that face named exactly `CUTFILE MARKING`.
4. Put every pin-stamp line and SolidWorks sketch-text object in that sketch.
5. Use a single-line/stick font when the downstream marking process requires
   single-stroke geometry. Ordinary fonts export as their rendered outlines.
6. Construction geometry is intentionally ignored.

The marking sketch is optional. A part without it still receives correct inside
and outside cut layers and reports zero marking paths.

## Run

Double-click `Launch Cut File Exporter.bat`, select the folder containing the
parts, select an output folder, and choose **Create Layered DXFs**. SolidWorks
opens or becomes visible while the batch runs.

Command-line equivalent:

```powershell
py cutfile_exporter.py --input "C:\Job\Parts" --output "C:\Job\Cut Files"
```

Add `--recursive` for subfolders or `--overwrite` to replace same-name outputs.
Use `--sketch-name "ANOTHER NAME"` only when the job uses a controlled alternate
marking-sketch name.

## Review

Open `cutfile_export_report.csv` and every failed part. Visually inspect the DXF
layers, scale, holes, outside profile, and representative marking text before
releasing files to nesting or cutting. Existing DXFs are skipped unless
overwrite is explicitly enabled.

The first Windows/SolidWorks run is an integration acceptance test because the
SolidWorks COM application is not available in Linux CI.

