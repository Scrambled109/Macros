# Modified Parts List Properties

This is the maintained source for `(MOD)2(SECONDARY).swp`. It copies mapped
spreadsheet values into SolidWorks part properties.

## Runtime behavior

1. Run from an open SolidWorks assembly.
2. If Excel already has an active workbook, that workbook is used. Otherwise,
   the macro opens a workbook picker and loads the selected workbook read-only.
3. The `Parts_List` worksheet is preferred. If it is absent, the active sheet
   (or first worksheet) is used.
4. A temporary Excel mapping sheet appears every run. Use its three dropdowns
   to choose the part-number/filename column and the columns that should
   populate `Description` and `Raw_Material`, then return to SolidWorks and
   click **OK**. The temporary mapping workbook closes without saving.
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
2. If the broken mapper was previously imported, remove
   `ModifiedPartsListMapper`. Raw lines beginning with `VERSION 5.00` or
   `Begin VB.UserForm` must not remain in any code module.
3. Remove the old `Module_MOD_2_SECONDARY_` or
   `ModifiedPartsListProperties` module.
4. Choose **File > Import File**, then import only
   `ModifiedPartsListProperties.bas` from this folder.
5. Choose **Debug > Compile VBAProject**.
6. Save the SWP.
7. Test on a copied assembly and copied parts before using production files.

The macro uses the existing SolidWorks type-library references and late-bound
Excel automation. The mapping UI is generated in a temporary Excel workbook,
so no VBA UserForm or binary `.frx` dependency is required.
