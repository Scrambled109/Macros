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
| `TextMarking.bas`       | Harvest the words from the text layer (AutoCAD) and recreate them as native sketch text on the part's top face (SolidWorks). |
| `TextStamp.bas`         | `RunTextStamp()` entry point — the separate, re-runnable pass that stamps the words onto finished parts. |
| `Main.bas`              | `RunBatch()` entry point, orchestration, progress, summary, cleanup. |

---

## Setup

1. Open the VBA IDE in your host application (AutoCAD: `VBAIDE`; or SolidWorks:
   **Tools ▸ Macro ▸ Edit**). The project runs from **either** host.
2. **Import all seven `.bas` files** (File ▸ Import File…).
3. Add **all three** references (Tools ▸ References…):
   - **AutoCAD 2026 Type Library**
   - **SOLIDWORKS 2025 Type Library** (`sldworks.tlb`)
   - **SOLIDWORKS 2025 Constant Type Library** (`swconst.tlb`) — **required**:
     the import method is set with the real `swImportDxfDwg_ImportToPartSketch`
     enum so a wrong numeric guess can never silently import the DWG as a
     drawing. If this reference is missing the project will not compile
     (`Variable not defined` on that name), which is intentional.
4. Run the two stages **in order**:
   1. **`Main.RunBatch`** — converts every DWG to an extruded `SLDPRT`.
   2. **`TextStamp.RunTextStamp`** — stamps the words onto the finished parts.

The `FilteredDWGs` and `staging` folders are created automatically if missing.

The two stages are deliberately **separate** for this first version: the extrude
batch never touches text, and the text pass can be re-run at any time against
the `staging` folder without rebuilding parts.

---

## Configuration (`Config.bas`)

The three folder paths, `TARGET_LAYER`, `EXTRUDE_DEPTH_METERS`, `TEXT_LAYER`
(`PIN STAMP TEXT`) and `DWG_UNITS_TO_METERS` (`0.0254`, inches) are already set
to the values for this job. Other useful switches:

| Constant                 | Default | Effect |
|--------------------------|---------|--------|
| `APP_VISIBLE`            | `False` | Keep both apps hidden for speed. Set `True` to watch/debug. |
| `UNLOCK_LOCKED_LAYERS`   | `True`  | Unlock locked non-target layers so their entities can be deleted. Locked layers are always reported in the log. |
| `QUIT_APPS_ON_FINISH`    | `False` | Leave the apps open (safe default) or shut them down. |
| `SHOW_SUMMARY_DIALOG`    | `False` | Show one message box at the end. Off = fully unattended. |
| `IMPORT_MERGE_METERS`    | `0.000254` | Gap that import will snap closed (helps close tiny open contours). |
| `TEXT_LAYER`             | `"PIN STAMP TEXT"` | Layer holding the words (blank = skip text). |
| `DWG_UNITS_TO_METERS`    | `0.0254` | DWG unit → meter scale for placing the words. `0.0254` = inches, `0.001` = mm. |

---

## The words (part marking)

Trying to import DWG text as extrudable geometry and fuse it with the outline in
one sketch is exactly the approach that fights you — `TEXT`/`MTEXT` doesn't come
in as clean closed loops, and letters with holes (A, O, R…) become nested
islands. So this project **doesn't** do that.

This is done as a **separate pass** (`TextStamp.RunTextStamp`) that runs after
`RunBatch`, so text can never damage the base parts and can be re-run on its
own. For each `SLDPRT` in `staging`:

1. **AutoCAD** opens the matching **source DWG read-only** (its `PIN STAMP TEXT`
   layer is still intact) and captures each word's string, position, height and
   rotation as plain data (MTEXT formatting codes are cleaned off).
2. **SolidWorks** opens the part and recreates each word as **native sketch
   text** on the **top face of the extruded part** (auto-detected — the planar
   face whose normal is +Z, at the greatest Z; it falls back to the Front plane
   if the top face can't be found), **leaves the sketch un-extruded** (words are
   just there for modeling — depth 0, not cut, not embossed), rebuilds and saves
   the part in place.

Each part's log line shows how many words were placed:

```
Words found       : 3
Text stamped      : OK
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

## Why the extrude used to fail right after import

The DWG part-sketch import leaves the imported sketch **open in edit mode**,
and an active sketch cannot be selected as a feature — so the extrude failed
before it started. The converter now exits sketch edit mode immediately after
the import (`SketchManager.InsertSketch True`, the API equivalent of clicking
*Exit Sketch*) and retries once inside the extrude helper if selection still
fails. If an extrude fails now, the log states the exact reason: the sketch
could not be selected, `FeatureExtrusion3` created no feature (profile open /
self-intersecting / empty), or a runtime error with its description.

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
- **Import method** (`SolidWorks_Import.SetPartImportMethod`): the DWG
  **default** import method is *create new drawing*, so this setting is the
  one thing that is **never** guarded or defaulted. `ImportMethod` is an
  **indexed** property — `ImportMethod(sheetName)` — and is set to
  `swImportDxfDwg_ImportToPartSketch` (from `swconst.tlb`) with sheet index
  `""` first and the file path as fallback. If neither form is accepted the
  file **fails loudly** in the log instead of quietly becoming a drawing, and
  after `LoadFile4` the document type is verified to be a part (`swDocPART`).
- **Optional import options** (`SolidWorks_Import.ConfigureImportOptions`):
  `SetMergePoints` (merge coincident end points within `IMPORT_MERGE_METERS`)
  and `ImportDimensions = False` are set under `On Error Resume Next` — these
  are nice-to-haves, and a member your SP names differently is simply skipped.
  Use the Object Browser (`F2`) on `ImportDxfDwgData` to confirm names if you
  want every option enforced.

The `FeatureExtrusion3` and `IModelDocExtension.SaveAs` signatures used are the
long-stable ones and require no adjustment.
