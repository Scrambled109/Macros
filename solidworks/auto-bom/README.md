# Modified Parts List Property Macro

`AutoBOMProperties.swp` is the current runnable modified Parts List property
macro on this branch.

Run it from an open SolidWorks assembly using **Tools > Macro > Run**. It:

- uses the active Excel workbook, or opens a file picker when Excel has no active
  workbook;
- creates a temporary Excel mapping sheet with dropdowns for part number,
  `Description`, and `Raw_Material`;
- processes selected components, or every unique assembly part/configuration
  when no components are selected;
- updates global and referenced-configuration properties and silently saves each
  affected part; and
- reports unmatched, skipped, duplicate, and save-failed items.

This is a high-impact macro. Work from a recoverable model copy and review the
summary and VBA Immediate window before using the results downstream.

`AUTOBOMACTUAL.swp` remains available as a separate legacy/alternate AutoBOM
workflow; confirm approval before using it.
