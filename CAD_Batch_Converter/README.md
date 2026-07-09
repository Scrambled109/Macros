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
| `NativeSketch.bas`      | The DWG-import **workaround**: harvest the outline coordinates from AutoCAD and redraw them as a **native** SolidWorks sketch, then extrude. Runs automatically whenever the import path fails. |
| `TextMarking.bas`       | Harvest the words from the text layer (AutoCAD) and recreate them as native sketch text on the part's top face (SolidWorks). |
| `TextStamp.bas`         | `RunTextStamp()` entry point — the separate, re-runnable pass that stamps the words onto finished parts. |
| `Main.bas`              | `RunBatch()` entry point, orchestration, progress, summary, cleanup. |

---

## Setup

1. Open the VBA IDE in your host application (AutoCAD: `VBAIDE`; or SolidWorks:
   **Tools ▸ Macro ▸ Edit**). The project runs from **either** host.
2. **Import all eight `.bas` files** (File ▸ Import File…).
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
| `TEXT_LAYER`             | `"PIN STAMP TEXT"` | Layer(s) holding the words — one name, or several separated by commas, e.g. `"PIN STAMP TEXT, PART MARKING, ETCH"` (blank = skip text). |
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

1. **AutoCAD** opens the matching **source DWG read-only** (its text layer(s)
   are still intact) and captures each word's string, position, height and
   rotation as plain data (MTEXT formatting codes are cleaned off).
2. **SolidWorks** opens the part and **draws each word as single-stroke line
   geometry** (a built-in stick font rendered through
   `SketchManager.CreateLine` — the same proven call the outline workaround
   uses) in a sketch on the **top face of the extruded part** (auto-detected —
   the planar face whose normal is +Z, at the greatest Z; it falls back to the
   Front plane if the top face can't be found). Position, **height and
   rotation** all come from the DWG. The sketch is **left un-extruded** (words
   are just there for modeling — depth 0, not cut, not embossed); the part is
   rebuilt and saved in place.

Each part's log line shows how many words were placed:

```
Words found       : 3
Text stamped      : OK
```

**If the words land in the wrong place**, `DWG_UNITS_TO_METERS` is the single
knob — it must match the drawing's units (inches by default).

**Engraving (readability):** bare sketch lines are hard to read in shaded
views (hairline strokes + endpoint dots), so by default the pass follows the
drawing step with a **shallow thin-slot cut** along the strokes —
`TEXT_ENGRAVE_DEPTH_M` deep (0.2 mm default), stroke width =
`TEXT_STROKE_WIDTH_FRAC` × cap height. The words become crisp shaded grooves,
like the real pin-stamped part. Both cut directions and two thin-cut API forms
are attempted, verified by the returned feature; if none succeeds the words
stay as sketch lines and the log says so. Set `TEXT_ENGRAVE = False` for
sketch-lines-only behaviour.

**Why the words are drawn instead of using SolidWorks sketch text:**
`IModelDoc2::InsertSketchText` proved unusable unattended on this install — it
returned `Nothing` for every word, with the app visible or hidden, on both
known call signatures. So the pass no longer uses any text API at all: each
word is rendered from a **built-in single-stroke font**
(`TextMarking.StrokeFor`) as plain `CreateLine` segments — deterministic, works
hidden, and a stick font is what pin-stamp / dot-peen equipment engraves
anyway. Unknown characters draw as a small box so a missing glyph is visible
rather than silent (add glyphs in `StrokeFor` if your drawings use exotic
symbols). Every placement is counted; the log reports `Placed X of Y words` on
any shortfall, so a silent no-op is impossible.

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

Two stacked causes, both fixed:

1. The DWG part-sketch import leaves the imported sketch **open in edit mode**,
   and an active sketch cannot be selected as a feature. The converter now
   exits sketch edit mode right after the import (`SketchManager.InsertSketch
   True`, the API equivalent of clicking *Exit Sketch*) and re-activates the
   imported part (`ActivateDoc3`), since selection acts on the active document.
2. Selecting the imported sketch **by name** (`SelectByID2("Model", "SKETCH",
   …)`) proved unreliable — the log showed `could not select sketch 'Model'` —
   notably with SolidWorks hidden (`APP_VISIBLE = False`). The sketch is now
   selected as a **feature object** (`IFeature::Select2`), which has no
   name/visibility dependency; `SelectByID2` remains only as a backup.

If an extrude fails now, the log states the exact reason: the sketch could not
be selected (and the contour fallback also failed), `FeatureExtrusion3` created
no feature (profile open / self-intersecting / empty), or a runtime error with
its description.

If the whole-sketch extrude is rejected, the converter automatically retries by
selecting only the sketch's **closed contours** and extruding those — the API
equivalent of dropping the profile into the **Selected Contours** box of the
Boss-Extrude PropertyManager (e.g. `Model-Contour<1>`), which is exactly what
works when the extrude is done by hand on these imports. Open fragments and
stray segments are simply left unselected, so they can't block the extrude.

---

## The native-sketch workaround (no DWG import at all)

The unattended DWG import (`LoadFile4`) has proven flaky: it drops the part
into the interactive **2D-to-3D** / sketch-edit state, and has produced an
unselectable or even **blank** `Model` sketch — while the same profiles import
and extrude fine by hand. So the converter now carries a second, import-free
route (`NativeSketch.bas`), used automatically whenever the import path fails
to deliver a saved part:

1. **AutoCAD** reopens the filtered DWG read-only and reads the outline out as
   plain coordinates: lines, polylines (bulges converted to true arcs), arcs
   and circles.
2. **SolidWorks** creates a fresh part from the default part template and
   **redraws** that geometry as a native sketch on the front plane (scaled by
   `DWG_UNITS_TO_METERS`), then extrudes and saves as usual. Native sketches
   have none of the import quirks — they extrude exactly like one drawn by
   hand.

When this route rescues a file, its log block ends with
`Recovered via native-sketch workaround (outline redrawn from AutoCAD
geometry).` Splines/ellipses are not redrawn; if a drawing contains them on
the cut layer, they are reported by name in the log and the file fails loudly
rather than producing a wrong outline.

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
