# EPL-to-Parts-List Converter

This is a standalone, offline Windows converter for local Engineering Parts List
(EPL) and Build/Breakdown of Parts (BOP) workbooks. A BOP is required and is the
scope authority: an EPL row is exported only when the same LOA and Part No.
appear in a selected BOP. One run creates:

- `{name} - Plates.xlsx`
- `{name} - Shapes.xlsx`
- `{name} - Conversion Report.xlsx`
- `{name} - Conversion Report.json`

Each selected EPL gets one tab in both output workbooks. Tabs follow the existing
`LOA - drawing identifier` convention, such as `LOA012095 - R5610102`.

The Plates workbook uses the supplied blue, two-row header layout with Hull No.,
the eight PROCESS columns, and Remarks. If a Plates template is selected, the
converter copies the first visible sheet's header styles, merged cells, widths,
row heights, borders, page settings, and data-row style without copying its old
parts or row-specific process marks.

## Put it in the requested Windows folder

Copy this entire project folder to:

`C:\Users\pbowen\Downloads\Macros\epl_converter`

Do not copy only `main.py`; the package, JSON settings, and build files are also
required.

## Run from Python

Open Command Prompt in the project folder, then run:

```bat
py -m pip install -r requirements.txt
py main.py
```

Running with no arguments opens the desktop application.

## Build the Windows EXE

The build must be run on Windows because PyInstaller does not cross-compile a
Windows EXE from Linux or macOS.

```bat
build_exe.bat
```

The finished files are:

```text
dist\EPLConverter.exe
dist\material_translations.json
```

Keep those two files together. The JSON file is intentionally editable.

With the computer offline, smoke-test the built EXE and its CLI against local
copies of an EPL and BOP:

```bat
smoke_test_exe.bat "C:\Path\To\LOA012095_B-EPL.xlsx" "C:\Path\To\BOP.xlsm"
```

The batch file checks the process exit code and all four expected outputs. It
writes only to a new temporary folder and prints that folder when finished.

## Desktop workflow

The current window identifies itself as **Version 1.1.1**. If that version is
not shown, an older EXE is being opened and must be rebuilt from this project.

1. Under **1. EPL WORKBOOKS**, click **Choose EPL File(s)** and select every EPL
   for this run.
2. Under the orange **2. BOP / BOM WORKBOOKS — REQUIRED** panel, click
   **Choose Required BOP/BOM File(s)** and select the workbook(s) defining scope.
3. Optionally select the Plates template and JSON/CSV enrichment.
4. Choose the output folder.
5. Enter the shared output name.
6. Click **Convert**.
7. Review the concise completion message, then open the Conversion Report for
   missing metadata, unmapped materials, duplicates, and unclassified items.

The application never scans a folder. It opens only the files explicitly
selected or supplied on the command line.

## CLI

```bat
EPLConverter.exe --epl "LOA012095_B-EPL.xlsx" "LOA013501_C.1-EPL.xlsx" ^
  --bop "828-BHD155-BOPB_JBG_WS.xlsm" ^
  --plate-template "Parts List (Plates Only).xlsx" ^
  --output-dir "C:\Output" ^
  --name "063 - HII - CLB BHD 155 Parts List" ^
  --metadata "metadata.json"
```

Add `--json` for a machine-readable response.

## Orchestrator JSON

```bat
EPLConverter.exe --request request.json
```

The request format is:

```json
{
  "epl_files": ["C:\\Input\\LOA012095_B-EPL.xlsx"],
  "bop_files": ["C:\\Input\\828-BHD155-BOPB_JBG_WS.xlsm"],
  "plate_template": "C:\\Input\\Parts List (Plates Only).xlsx",
  "output_dir": "C:\\Output",
  "output_base": "063 - HII - CLB BHD 155 Parts List",
  "metadata_by_part": {
    "R5610102-557": {
      "mdlprt": "MDLPRT003964722",
      "width": "48\"",
      "length": "8' - 0\"",
      "beveled": true,
      "other_info": "FIELD VERIFY"
    }
  }
}
```

Use `--request -` to read the JSON request from standard input.

## Python API

```python
from epl_converter import convert_epls

result = convert_epls(
    epl_paths=["LOA012095_B-EPL.xlsx", "LOA013501_C.1-EPL.xlsx"],
    output_dir=r"C:\Output",
    output_base="063 - HII - CLB BHD 155 Parts List",
    metadata_by_part={
        "R5610102-557": {
            "mdlprt": "MDLPRT003964722",
            "width": "48\"",
            "length": "8' - 0\"",
            "beveled": True,
            "other_info": "",
        }
    },
    bop_paths=[r"C:\Input\828-BHD155-BOPB_JBG_WS.xlsm"],
    plate_template_path=r"C:\Input\Parts List (Plates Only).xlsx",
)
```

## BOP scope and data precedence

- Matching key: normalized BOP LOA plus exact EPL Part No.
- BOP supplies scope, Hull No., and MDLPRT.
- EPL supplies quantity, MF Part No., description, material, and thickness.
- Duplicate BOP occurrences include the part once; EPL order is preserved.
- A BOP/EPL MF Part No. disagreement is reported and the EPL value is retained.
- BOP LOAs without a selected EPL and BOP parts missing from their selected EPL
  are reported.
- EPL rows absent from the BOP are counted as out of scope.

## Metadata

JSON can be keyed directly by `part_no`, or can contain a
`metadata_by_part` object. CSV headers are:

```text
part_no,mdlprt,width,length,beveled,other_info
```

Only a real boolean `true` in JSON, or `true/yes/1/y` in CSV, adds `BEVELED`.
The converter does not infer bevels or machining.

## Material rules and classification

Edit `material_translations.json` beside the EXE to add verified mappings or
stock keywords. Regular expressions are matched against the original material
specification, description, and type. Unknown materials are preserved in the
output and reported. They are never substituted.

The seeded rules cover:

- T9074 HY-80 -> `HY-80`
- MIL-S-22698 EH-36T -> `MIL-S-22698_EH-36T`
- applicable 316L text/specifications -> `CRES 316L`

The conservative default plate keywords are `PLATE`, `SHEET`, `SHIM`, and
`WEDGE`. Items such as `PIG` are reported as unclassified until someone adds a
verified rule.

## Tests

```bat
py -m unittest discover -s tests -v
```

The automated tests cover required BOP input and scope gating, assemblies,
plate/shape classification, seeded and
unknown materials, explicit dimensions, metadata, bevel handling, duplicate
parts, duplicate LOAs, malformed workbooks, output formatting, and the
31-character Excel tab limit.
