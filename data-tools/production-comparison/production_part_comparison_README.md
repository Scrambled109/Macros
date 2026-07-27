# Production Part Reconciliation

## Files

- `compare_production_parts.py` — comparison and report program
- `comparison_rules.json` — editable rules, keywords, aliases, checks, and tolerances

Keep both files in the same folder.

The program uses only the Python standard library. No additional Python packages are required.

## Input exports

### Parts List CSV

Include:

- PART NUMBER
- DESCRIPTION
- TOTAL QUANTITY
- THICKNESS/SHAPE
- LENGTH
- MATERIAL
- WIDTH is optional

The script uses `DESCRIPTION` containing `PLATE` to identify plates.

### SolidWorks Assembly Visualization CSV

Include:

- FILE NAME
- QUANTITY
- SHAPE
- LENGTH
- MATERIAL
- DESCRIPTION is optional
- WIDTH is optional

SolidWorks is the baseline for required quantity and geometry.

Known SolidWorks configuration suffixes such as `Default` and
`Default Machined` are removed from the end of the file name. Revision suffixes
are preserved.

### Linear nesting files

Put every nest CSV/TXT in one folder. They may appear in one Excel column.
The script reads the semicolon-delimited `Parts;Length;Quantity` section
directly.

## Run

Double-click `compare_production_parts.py`, or run:

```powershell
py compare_production_parts.py
```

The program asks you to select:

1. The linear nesting folder
2. The Parts List CSV
3. The SolidWorks CSV
4. A base output folder

It creates a timestamped report folder.

## Primary outputs

### `production_part_comparison.xlsx`

This is the formatted Excel report.

Visible sheets:

- **Errors Requiring Action** — comparison errors plus parts missing from SolidWorks or the Parts List
- **All Comparisons** — every part in the compact report format

Both sheets freeze the following columns so the part identification remains visible while scrolling:

- Part Number
- Review Status
- Category
- Description

The workbook also contains hidden **Technical Details** and **Source Data Issues** sheets. They can be unhidden in Excel when the complete audit data is needed.

### `errors_requiring_action.csv`

Compact actionable output. It includes:

- Part Number
- Review Status
- Category
- Description
- Problem
- only the SolidWorks, Parts List, or Nest values involved in the problem
- Required Action

Parts missing from SolidWorks or the Parts List are included here so they are not overlooked.

### `all_comparisons.csv`

The same compact layout for every part. Exact matches do not repeat every source value; they simply report that all required checks match.

### `technical_details.csv`

The complete prior detailed output, including normalized keys, deltas, source rows, and every comparison field.

## Other outputs

- `comparison_report.html` — browser report with frozen Part Number, Review Status, Category, and Description columns
- `missing_from_solidworks_or_parts_list.csv`
- `not_checked.csv`
- `exact_matches.csv`
- `source_data_issues.csv`
- `source_rows.csv`
- `run_summary.txt`

Plates are compared between SolidWorks and the Parts List, but their nesting
status remains `NOT CHECKED` until the plate-nesting parser is added.

## Quantity-column handling

Parts List exports may contain both `QUANTITY`/`QTY` and `TOTAL QUANTITY`.

- The program counts usable numeric values in every recognized quantity column and selects the most populated column.
- When columns are equally populated, `TOTAL QUANTITY` is preferred.
- This prevents a blank `TOTAL QUANTITY` column from overriding a populated `QTY` column.
- `TOTAL QUANTITY` is already the complete part total, so duplicate rows are not summed again.
- `QUANTITY` or `QTY` values are treated as per-row quantities and duplicate rows are summed.
- The selected column and the numeric-row count for every available quantity column are recorded in the report diagnostics.

## Updating rules

Edit `comparison_rules.json`.

Common updates:

- Add another plate spelling to `plate_description_keywords`
- Add hardware terms to `excluded_hardware_description_keywords`
- Add known SolidWorks configuration names
- Add shape aliases when two systems describe the same profile differently
- Add material aliases
- Add alternate column headings
- Change length or width tolerances
- Set `output.create_excel_workbook` to `false` to disable the formatted workbook
