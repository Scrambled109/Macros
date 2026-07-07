# CAD Batch Converter (VBA)

Automated batch converter that, for **every DWG** in a source folder:

1. Opens the drawing in **AutoCAD 2026**,
2. deletes every entity **not** on the `CUT-OUTSIDE STRAIGHT` layer,
3. runs **AUDIT** + **PURGE** and saves a filtered DWG,
4. imports that DWG into **SolidWorks 2025** as a **new-part 2D sketch**,
5. finds the imported sketch, detects open contours, best-effort repairs it,
6. **blind-extrudes** it 0.00635 m (6.35 mm / 0.25 in), merge = true,
7. rebuilds and saves `filename.SLDPRT` into the staging folder (overwrite),
8. closes the document and moves on.

The batch is **fully unattended** and **never stops for one bad file** — every
stage has structured error handling, and every file is logged.

---

## Modules

| File                    | Responsibility |
|-------------------------|----------------|
| `Config.bas`            | All paths, layer name, depth, tolerances, ProgIDs, switches, the `TFileResult` type. **Every tunable lives here.** |
| `Utilities.bas`         | Logging, folder creation, timing, result serialisation, small helpers. Host-independent. |
| `AutoCAD_Filter.bas`    | Connect to AutoCAD; open → strip non-target layers → AUDIT → PURGE → SaveAs → close. |
| `SolidWorks_Import.bas` | Connect to SolidWorks; import DWG → find sketch → detect open contour → extrude → save → close. |
| `TextMarking.bas`       | Harvest the words from the text layer (AutoCAD) and recreate them as native sketch text on the part (SolidWorks). |
| `Main.bas`              | `RunBatch()` entry point, orchestration, progress, summary, cleanup. |

---

## Setup

1. Open the VBA IDE in your host application (AutoCAD: `VBAIDE`; or SolidWorks:
   **Tools ▸ Macro ▸ Edit**). The project runs from **either** host.
2. **Import all six `.bas` files** (File ▸ Import File…).
3. Add **both** references (Tools ▸ References…):
   - **AutoCAD 2026 Type Library**
   - **SOLIDWORKS 2025 Type Library** (`SldWorks.tlb`)
4. Run **`Main.RunBatch`** (F5, or Macros dialog).

The `FilteredDWGs` and `staging` folders are created automatically if missing.

---

## Configuration (`Config.bas`)

The three folder paths, `TARGET_LAYER`, and `EXTRUDE_DEPTH_METERS` are already
set to the values from the job spec.

**Before your first run, set `TEXT_LAYER`** to the exact layer that holds the
words (leave it `""` to skip text marking). Other useful switches:

| Constant                 | Default | Effect |
|--------------------------|---------|--------|
| `APP_VISIBLE`            | `False` | Keep both apps hidden for speed. Set `True` to watch/debug. |
| `UNLOCK_LOCKED_LAYERS`   | `True`  | Unlock locked non-target layers so their entities can be deleted. Locked layers are always reported in the log. |
| `QUIT_APPS_ON_FINISH`    | `False` | Leave the apps open (safe default) or shut them down. |
| `SHOW_SUMMARY_DIALOG`    | `False` | Show one message box at the end. Off = fully unattended. |
| `IMPORT_MERGE_METERS`    | `0.000254` | Gap that import will snap closed (helps close tiny open contours). |
| `TEXT_LAYER`             | `""` | Layer holding the words. **Set this.** |
| `DWG_UNITS_TO_METERS`    | `0.0254` | DWG unit → meter scale for placing the words. `0.0254` = inches, `0.001` = mm. |

---

## The words (part marking)

Trying to import DWG text as extrudable geometry and fuse it with the outline in
one sketch is exactly the approach that fights you — `TEXT`/`MTEXT` doesn't come
in as clean closed loops, and letters with holes (A, O, R…) become nested
islands. So this project **doesn't** do that.

Instead (`TextMarking.bas`):

1. **AutoCAD** reads every `TEXT`/`MTEXT` on `TEXT_LAYER` *before* the strip and
   captures each word's string, position, height and rotation as plain data
   (MTEXT formatting codes are cleaned off).
2. **SolidWorks** recreates each word as **native sketch text** on the part's
   **Front plane** (coincident with the imported outline, so the words line up
   in plan), and **leaves the sketch un-extruded** — the words are just there
   for modeling, per your spec (depth 0, not cut, not embossed).

Each file's log line shows how many words were placed:

```
Text marks        : 3 word(s) - OK
```

**If the words land in the wrong place**, `DWG_UNITS_TO_METERS` is the single
knob — it must match the drawing's units (inches by default).

**One version note:** the actual `InsertSketchText` call is late-bound and
isolated in `TextMarking.InsertOneText`, and its exact argument list has varied
between SolidWorks releases. If a word fails to place, record a one-line macro
of "insert sketch text" on your install and match that argument order in that
one helper — nothing else needs to change. Sketch text uses the document font
size (absolute height from the DWG is not enforced); adjust the document font or
that helper if you need the DWG height honoured. Rotation is captured but text
is placed horizontal by default.

---

## Log

Written to `staging\BatchLog.txt` (appended across runs). Each file gets a block:

```
-------------------------------------------------------------------
Timestamp         : 2026-07-07 14:03:22
File              : PART_1234.dwg
AutoCAD filter    : OK
SolidWorks import : OK
Extrusion         : OK
Save SLDPRT       : OK
Locked layers     : DIMENSIONS, NOTES
Processing time   : 3.41 s
```

Failures record the stage that failed plus the exception message, then the
batch continues.

---

## How open contours are handled

Small gaps are merged on import (`IMPORT_MERGE_METERS`). Open contours are then
detected **geometrically** by pairing sketch-segment end points — any unpaired
end means the profile is open, which is flagged in the log. If the profile is
genuinely open the blind extrude will fail; the reason is logged and the batch
moves to the next file.

---

## Version-specific notes (please verify against your installed SP)

This project is written against the documented AutoCAD 2026 / SolidWorks 2025
APIs and uses **defensive, guarded** calls so a member that shifted between
service packs cannot abort a run. Two spots are worth a sanity check on first
use — both are isolated and clearly commented:

- **ProgIDs** (`Config.bas`): `AutoCAD.Application.25` and
  `SldWorks.Application.33`. If attach/launch ever falls back, the generic
  ProgIDs (`AutoCAD.Application`, `SldWorks.Application`) are tried next, so the
  batch still runs; adjust the versioned strings only if you want to pin a
  specific install.
- **Import options** (`SolidWorks_Import.ConfigureImportData`): each property
  (`ImportMethod`, `MergePoint`, `MergePointDistance`, `ImportDimensions`,
  `ImportSketchAsConstruction`, `EntitiesToImport`) is set under
  `On Error Resume Next`. Any property your SP names differently is simply
  skipped — SolidWorks then applies its default (new-part 2D sketch), which is
  exactly the behaviour we want. Use the Object Browser (`F2`) on
  `ImportDwgDxfData` to confirm names if you want every option enforced.

The `FeatureExtrusion3` and `IModelDocExtension.SaveAs` signatures used are the
long-stable ones and require no adjustment.
