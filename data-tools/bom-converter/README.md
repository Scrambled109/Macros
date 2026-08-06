# BOP/BOM to Parts List Converter

This tool detects a source table, keeps the configured production part rows,
maps them into a Parts List, and reports missing required headings before it
writes output.

Install requirements once from the repository root:

```powershell
py -m pip install -r requirements.txt
```

For the guided interface, double-click `bom_converter.py`. For a command-line
run:

```powershell
py data-tools/bom-converter/bom_converter.py "C:\Job\A-BOM.xlsm" "C:\Job\TEMP.xlsx" "C:\Job\MASTER_PARTS_LIST.xlsx"
```

Close all input and output workbooks in Excel first. Keep
`bom_converter_mapping.json` beside the script. The converter appends mapped
rows to an existing output workbook, so do not repeat a run unless duplicate
rows are intended. Review quantities, material, thickness, and descriptions
before using the Parts List downstream.

Template columns are resolved by normalized header name, not fixed Excel column
letters. Engineers may insert or reorder additional information columns without
shifting mapped data. Duplicate normalized headings are rejected because they
would make the destination ambiguous.

If the template uses an Excel table, the converter extends that table through
the final written row while preserving its existing columns. Alternating row
banding therefore continues beyond the template's original blank-row limit.

