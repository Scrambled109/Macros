# CAD Batch Converter

`Main.RunBatch.swp` is the runnable SolidWorks macro. It supports AutoCAD 2025
or 2026 and processes prepared DWG
profiles, creates SolidWorks plate parts, and recreates pin-stamp text/line
marking as an unconsumed sketch on each part.

## Recommended run

Use the Engineering Job Assistant. Select one reviewed material/thickness DWG
folder and confirm the plate thickness. The assistant supplies the source,
filtered-DWG, output, and extrusion-depth settings at run time, starts
SolidWorks when configured, and opens the macro folder. Wait until SolidWorks
has fully loaded before running the `.swp`.

Then in SolidWorks:

1. Choose **Tools > Macro > Run**.
2. Select `Main.RunBatch.swp`.
3. Let the macro's `main()` entry point run.
4. Read `BatchLog.txt` in the output folder.
5. Confirm the logged source folder, output folder, requested extrusion depth,
   saved parts, and marking counts before processing the next thickness group.

If the folders were not configured by Job Assistant, the macro stops before it
opens drawings. This user branch contains the compiled `.swp`, not editable
`.bas` modules.

## Expected layers

- `CUT - OUTSIDE STRAIGHT` — required exterior profile
- `CUT - INSIDE STRAIGHT` — holes and cutouts
- `PIN STAMP LINE MARKING` and `PIN STAMP TEXT` — marking geometry/text
- `PLOT` — discarded

Layer spelling must match the prepared DWG contract. The log reports actual
model-space layers when the required exterior is missing.

## Safety and recovery

The batch continues after an individual file fails. It writes one log block per
file and uses a native-sketch fallback when the normal DWG import cannot produce
a usable SolidWorks sketch. Unsupported profile geometry fails loudly instead
of silently producing a wrong outline.

The marking pass reads TEXT, MTEXT, block attributes, and supported marking
geometry from the source DWG and draws deterministic single-stroke sketch
geometry on the part's top face. It does not cut or emboss the words. Review the
reported words/segments placed and visually inspect representative parts.

Keep the source DWGs and filtered copies until every generated part has been
checked.

