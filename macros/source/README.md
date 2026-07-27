# Extracted SolidWorks Macro Source

This directory contains best-effort text extraction from the legacy `.swp`
files. It makes their VBA logic searchable and reviewable without changing the
runnable macro binaries. `MANIFEST.sha256` records the exact binary and
extracted-source revisions used for this audit.

## Important limitation

The files were recovered directly from the OLE/VBA containers with
`tools/extract_swp_source.py`; they were **not** exported by SolidWorks and have
not been compiled in the SolidWorks VBA editor. Empty `ThisLibrary.bas` files
contain project metadata only. References, forms, digital signatures, and VBA
project settings are not reconstructed. A SolidWorks **Tools > Macro > Edit >
File > Export File** export remains authoritative and should replace these
copies if one is available.

## Inventory and initial audit

| Runnable macro | Extracted primary module | What it does | Initial safety notes |
|---|---|---|---|
| `drawing_auto.swp` | `drawing_auto/Module1.bas` | Adds drawing-view names, automatic dimensions, and/or an individual-part table. | Validates that a drawing is active and asks which actions to run. Automatic dimensioning and table insertion modify the open drawing; test on a copy. |
| `(MOD)2(SECONDARY).swp` | `MOD_2_SECONDARY/Module_MOD_2_SECONDARY.bas` | Copies Description and Raw_Material from columns C/I of the active Excel workbook's `Parts_List` sheet to selected components, matching the full part filename in column B. | Uses whichever Excel workbook is active. It modifies global and active-configuration properties in memory but does not explicitly save each part. Verify the active workbook and save results deliberately. |
| `AUTOBOMACTUAL.swp` | `AUTOBOMACTUAL/autobomtest1.bas` | Calculates bounding-box LENGTH/SHAPE properties for unique assembly parts. | High impact: inserts bounding-box features when missing, forcibly replaces LENGTH/SHAPE in every configuration, and silently saves every processed part. Select components first to limit scope; no selection means the entire assembly. |
| `AutoBOMProperties.swp` | `AutoBOMProperties/AutoBOMProperties1.bas` | Same behavior as `AUTOBOMACTUAL.swp`. | The recovered code is identical except for its VBA module name. Treat it as the same high-impact macro until SolidWorks testing establishes why both binaries exist. |
| `Bounding box.swp` | `Bounding_box/Bounding_box1.bas` | Finds or creates a global bounding box and converts its segments into a 3D sketch. | Does not verify document type, does not use `Option Explicit`, and has no error handler. Run only on a disposable part until compiled and tested. |
| `MaterialSpec(MOD).swp` | `MaterialSpec_MOD/MaterialSpec_MOD_1.bas` | Copies Description and Raw_Material from Excel to selected parts, matching the trailing numeric item number to column B. | Uses the active Excel workbook and only the active SolidWorks configuration plus global properties. It does not explicitly save each modified part. |
| `Test_test.swp` | `Test_test/Test_test1.bas` | Experimental bounding-box property writer for one part. | Explicitly a test macro. It overwrites `Box_Length`/`Box_Shape` and deletes the selected bounding-box feature afterward. Do not use on production data. |
| `hide sketches.swp` | `hide_sketches/hide_sketches1.bas` | Hides every visible sketch in the active part/assembly (`HIDE_ALL_SKETCHES = True`). | Changes visibility but not geometry. It restores the FeatureManager visibility preference in its error/finally path. Toolbar+/Batch+ is disabled by the compile constant. |

## How to reproduce the extraction

From the repository root:

```powershell
py tools/extract_swp_source.py drawing_auto.swp macros\*.swp
```

PowerShell may not expand `macros\*.swp` for native programs on every version.
If it does not, list each `.swp` path explicitly. On Bash:

```bash
python tools/extract_swp_source.py drawing_auto.swp macros/*.swp
```

The extractor uses only the Python standard library. It reads both normal and
mini OLE streams, decompresses MS-OVBA module containers, and writes a separate
UTF-8 `.bas` file for each source-bearing module. It never modifies an `.swp`.

## What still requires a Windows workstation

Before treating any recovered module as a supported production tool:

1. Open its `.swp` using SolidWorks **Tools > Macro > Edit**.
2. Record every checked item under **Tools > References**.
3. Export all standard modules, class modules, and forms.
4. Compare those official exports with this directory.
5. Use **Debug > Compile VBAProject**.
6. Run against disposable parts/assemblies and verify every property, feature,
   drawing annotation, and save operation.

Static extraction proves what code is stored in the binary; it cannot prove
that a particular SolidWorks release will compile or execute the API calls.
