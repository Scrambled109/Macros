# EPL-to-Parts-List Converter

This offline Windows tool converts selected EPL workbooks into Plates and Shapes
Parts Lists. At least one BOP/BOM workbook is required: it defines which LOA and
part-number combinations are in scope.

## Run it

From the repository root, install requirements once:

```powershell
py -m pip install -r requirements.txt
```

Then double-click `data-tools/epl_converter/main.py`, or run:

```powershell
py data-tools/epl_converter/main.py
```

Keep `main.py`, the `epl_converter` package folder, and
`material_translations.json` together. This user branch does not include an EXE
or EXE build files.

## Desktop workflow

1. Select every EPL workbook for the run.
2. Select the required BOP/BOM workbook(s) that define scope.
3. Optionally select a Plates template and metadata JSON/CSV.
4. Select an output folder and enter the shared output name.
5. Select **Convert**.
6. Review the Conversion Report for missing metadata, unmapped materials,
   duplicates, BOP/EPL disagreements, and unclassified items.

The converter creates Plates, Shapes, Conversion Report XLSX, and Conversion
Report JSON outputs. Unknown materials are preserved and reported; they are not
silently substituted. Edit `material_translations.json` only after a mapping has
been verified.

For command-line options, run:

```powershell
py data-tools/epl_converter/main.py --help
```

