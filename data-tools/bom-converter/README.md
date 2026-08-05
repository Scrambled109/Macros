# BOP/BOM to Parts List Converter

This tool reads the `Lofting` sheet, keeps the configured production part rows,
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

