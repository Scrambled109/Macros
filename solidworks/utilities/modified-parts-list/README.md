# Modified Parts List Properties

This is the maintained source for `(MOD)2(SECONDARY).swp`. It copies mapped
spreadsheet values into SolidWorks part properties.

## Runtime behavior

1. Run from an open SolidWorks assembly.
2. If Excel already has an active workbook, that workbook is used. Otherwise,
   the macro opens a workbook picker and loads the selected workbook read-only.
3. The `Parts_List` worksheet is preferred. If it is absent, the active sheet
   (or first worksheet) is used.
4. A mapping form appears every run. Choose the part-number/filename column and
   the columns that should populate `Description` and `Raw_Material`.
5. If components are selected, only those components are processed. With no
   component selection, every unique part/configuration in the assembly is
   processed.
6. Parts are matched case-insensitively by SolidWorks filename without the
   `.SLDPRT` extension. Properties are written globally and to the component's
   referenced configuration, then the part is saved silently.

The summary reports updated, unmatched, skipped, and save-failed parts.
Individual names are written to the VBA Immediate window (`Ctrl+G`). Duplicate
spreadsheet part numbers retain the first row and are reported.

## Rebuild the SWP

Source edits do not change the compiled `.swp` automatically.

1. In SolidWorks, choose **Tools > Macro > Edit** and open
   `solidworks\utilities\(MOD)2(SECONDARY).swp`.
2. Remove the old `Module_MOD_2_SECONDARY_` module.
3. Import `ModifiedPartsListProperties.bas` and
   `ModifiedPartsListMapper.frm` from this folder.
4. Choose **Debug > Compile VBAProject**.
5. Save the SWP.
6. Test on a copied assembly and copied parts before using production files.

The module uses the existing SolidWorks type-library references and late-bound
Excel automation. The form contains no images or binary assets, so it does not
require a companion `.frx` file.
