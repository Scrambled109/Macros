# Typical Production Workflow

This guide describes the normal order for taking an engineering BOM through cut
file preparation, part modeling, assembly, nesting, and the final production
comparison. It answers **what to use for each stage**; individual tool setup and
troubleshooting remain in the linked component documentation.

> **Production rule:** work from controlled copies, keep the job number and part
> number consistent at every handoff, and do not advance past a checkpoint with
> unexplained missing or duplicate parts.

The Engineering Job Assistant shows only steps it performs or directly
launches: Parts List, cut files, automatic plate models, AutoBOM, and production
comparison. Manual modeling, assembly, Parts List reconciliation, nesting, and
final approval remain documented here as external engineering work.

## Workflow at a glance

1. Create the initial Parts List from the A-BOM.
2. Prepare DXF cut files and sort them from the Parts List.
3. Automatically create the plate models from the cut files.
4. Manually model shapes and beveled parts.
5. Manually assemble the complete SolidWorks 3D model.
6. Run AutoBOM on the completed model to calculate `LENGTH` and `SHAPE`.
7. Compare those model values and update the Parts List.
8. Create linear nests in 1D Cutting Optimizer from the updated Parts List.
9. Export the final SolidWorks, Parts List, and nest data.
10. Run the production comparison and correct every actionable discrepancy.

## Tool map

| Stage | Primary tool | Result |
|---|---|---|
| A-BOM conversion | `data-tools/bom-converter/bom_converter.py` | Initial production Parts List workbook |
| DXF preparation and sorting | `autocad/dxf-orchestrator/Master_Orchestrator.py` with AutoCAD | Organized and converted cut files |
| Automatic plate modeling | `solidworks/cad-batch-converter/Main.RunBatch.swp` | Extruded SolidWorks plate parts |
| Shape and bevel modeling | Manual SolidWorks workflow | Modeled linear shapes and beveled parts |
| 3D model | Manual SolidWorks assembly | Completed production assembly |
| Model length/shape properties | `solidworks/auto-bom/AutoBOMProperties.swp` or the currently approved AutoBOM `.swp` | Model-derived `LENGTH` and `SHAPE` properties |
| Linear nesting | 1D Cutting Optimizer | Linear nest CSV/TXT exports |
| Final reconciliation | `data-tools/production-comparison/compare_production_parts.py` | XLSX, HTML, and CSV comparison reports |

The `.swp` files are the runnable SolidWorks macros. The `.bas` files in
`reference/` are review copies and are not used directly in production.

---

## 1. Create the initial Parts List from the A-BOM

### Inputs

- The engineering A-BOM workbook.
- An existing Parts List workbook or approved blank Parts List template.

### Use

Install the converter requirements once:

```powershell
py -m pip install -r data-tools/bom-converter/requirements.txt
```

Run the converter from the repository root:

```powershell
py data-tools/bom-converter/bom_converter.py "C:\Job\A-BOM.xlsm" "C:\Job\TEMP.xlsx" "C:\Job\MASTER_PARTS_LIST.xlsx"
```

The converter reads the A-BOM `Lofting` sheet, retains the applicable `DS`
parts, removes repeated routing rows according to its documented rules, and
appends the mapped data to the output workbook.

### Checkpoint before cut files

- Confirm the job is correct and the expected part numbers are present.
- Review quantities, material, thickness, and descriptions.
- Resolve duplicate or missing part numbers before using this list for sorting.
- Save a dated or revision-controlled copy of this **initial** Parts List. Its
  length and shape values may change after the completed model is measured.

See [`data-tools/bom-converter/bom_converter.py`](data-tools/bom-converter/bom_converter.py)
for the current converter and the root README for its required headings.

## 2. Prepare and organize the DXF cut files

### Inputs

- Raw DXF cut files in numbered job folders.
- A CSV named `Parts List.csv` exported from the Parts List when Parts List sorting
  is required. Its headings are `PartNumber`, `Quantity`, `Thickness`, and
  `Material`; the legacy spelling `Quanity` is also accepted.

### Use

Run the master orchestrator from the folder containing the numbered folders:

```powershell
py .\autocad\dxf-orchestrator\Master_Orchestrator.py --workspace "C:\Job\DXF Workspace" --parts-list "C:\Job\DXF Workspace\Parts List.csv" --workers 2
```

The AutoCAD workflow converts and organizes the cut files and uses the Parts
List data for material, thickness, and quantity-based sorting. Every drawing
runs headlessly through AutoCAD Core Console. Detected bevel drawings receive
the exact filename suffix `(B)` before `.dwg`; graphical AutoCAD does not open.
Bevel callout text and leader arrows are placed on `PLOT`, not on either
pin-stamp layer.

### Checkpoint before plate modeling

- Review `_ORCHESTRATOR_LOGS` and resolve all failed files.
- Confirm every expected plate DXF produced a valid output.
- Confirm material and thickness folders agree with the Parts List.
- Use the `(B)` filenames to identify parts that still require manual bevel
  modeling; automatic plate modeling does not create bevel geometry.
- Keep the source archive until the model and final comparison are accepted.

Full operating instructions are in
[`autocad/dxf-orchestrator/README.txt`](autocad/dxf-orchestrator/README.txt).

## 3. Automatically create plate models

### Inputs

- The prepared DWG/DXF-derived cut geometry from the AutoCAD stage.
- The assistant-configured source, filtered, output, and extrusion-depth
  runtime values used by the rebuilt production macro.

### Use

Process one material/thickness folder at a time. The Engineering Job Assistant
infers the thickness from the folder name, asks the operator to confirm it, and
supplies `EXTRUDE_DEPTH_METERS` to the macro at run time. It also reports the
number of sibling folders containing DWGs so the operator can account for every
required run. The production `.swp` must first be rebuilt once with the updated
runtime-aware `Config.bas`; loose `.bas` files cannot alter a compiled `.swp`.
The assistant opens guided instructions and does not treat opening an `.swp`
file as proof that its entry point executed.

In SolidWorks, choose **Tools > Macro > Run**, select:

```text
solidworks/cad-batch-converter/Main.RunBatch.swp
```

Use the `Main_RunBatch1.RunBatch` entry point. This stage filters the drawing
geometry, imports or reconstructs the outline, extrudes the plate, and saves a
SolidWorks part. Run the separate text-stamp pass only when required and only
after the base plate output has been checked.

### Checkpoint before manual modeling

- Read the batch log and account for every input file.
- Open representative parts from each material and thickness group.
- Verify outline geometry, units, extrusion thickness, orientation, file name,
  and marking/text placement.
- Rework any failed or incorrect plates before assembly.

See
[`solidworks/cad-batch-converter/README.md`](solidworks/cad-batch-converter/README.md)
for configuration and recovery details.

## 4. Manually model shapes and beveled parts

The automated batch is for ordinary plate creation. Model the following in
SolidWorks using the approved manual modeling standards:

- Linear shapes such as angles, channels, beams, tubes, bars, and other stock
  represented by a cross-section plus cut length.
- Beveled parts and bevel features not represented by the automatic plate
  extrusion.
- Any plate rejected during automatic validation that requires deliberate
  manual repair.

Keep the SolidWorks filename and production part number aligned with the Parts
List. Assign the correct material and verify the modeled cut length because the
completed model will later become the source for the AutoBOM `LENGTH` and
`SHAPE` comparison.

### Checkpoint before assembly

- Every Parts List item that requires a model has one validated part file.
- Shapes have correct section, material, and end conditions.
- Bevel direction, side, angle, and extent match the drawing.
- No temporary or test part is mixed into the production model folder.

## 5. Manually assemble the complete 3D model

Build the production assembly in SolidWorks from the validated automatic plate
parts and manually modeled shapes/bevels.

Before treating the assembly as complete:

- Confirm component part numbers and quantities against the Parts List.
- Resolve suppressed, lightweight, missing, or broken-reference components.
- Check orientation, location, fit, and obvious interference conditions.
- Confirm that alternate configurations are intentional.
- Save the assembly and all approved component changes.

This completed assembly is the input to AutoBOM. Do not run the final
length/shape reconciliation against a knowingly incomplete model.

## 6. Run AutoBOM on the completed model

Open the completed SolidWorks assembly and run the currently approved compiled
AutoBOM macro from `solidworks/auto-bom/`. The two retained binary names are
`AutoBOMProperties.swp` and `AUTOBOMACTUAL.swp`; use the one approved for the
job until the duplicate production name is formally resolved.

AutoBOM calculates bounding-box dimensions for the assembly's unique parts and
writes model-derived `LENGTH` and `SHAPE` custom properties. It is a
high-impact step because it overwrites those properties and saves processed
part files.

### Safe operating sequence

1. Save and back up the completed model.
2. If testing a subset, select only the intended components. With no selection,
   the macro processes the assembly.
3. Choose the required output unit consistently with the Parts List.
4. Run AutoBOM and review its processed/skipped summary and Immediate window.
5. Investigate every skipped part or failed save before exporting comparison
   data.

The readable `.bas` modules under `solidworks/auto-bom/reference/` are for
review only; changing them does not change the compiled `.swp` behavior.

## 7. Compare model values and update the Parts List

After AutoBOM completes, export the SolidWorks assembly/component data needed
to review each part's model-derived `LENGTH` and `SHAPE`.

Compare those values with the initial Parts List and update the Parts List so it
reflects the accepted completed model. This update is the handoff between
modeling and nesting: **do not build the final linear nests from the old,
pre-model length and shape values.**

Before nesting:

- Account for every AutoBOM skip or save failure.
- Confirm all linear shapes have an accepted `LENGTH` and `SHAPE`.
- Confirm quantities still agree after assembly completion.
- Save a clearly identified **post-model Parts List** revision.

## 8. Create linear nests in 1D Cutting Optimizer

Use the post-model Parts List—after the AutoBOM length/shape review and
updates—as the source for 1D Cutting Optimizer.

Create the linear nests, review stock selections and quantities, and export all
nest results as CSV or TXT files into one job-specific folder. Preserve the
exact exports used for the final comparison.

> The repository's comparison tool currently checks linear nest exports. Plate
> parts are compared between SolidWorks and the Parts List, but plate nesting is
> reported as `NOT CHECKED` until a plate-nesting parser is added.

### Checkpoint before final comparison

- All expected linear parts appear in a nest.
- Shape, stock length, material, and required quantities are correct.
- Remnants, optimization decisions, and operator overrides have been reviewed.
- All final nest CSV/TXT files are together in one folder.

## 9. Export the final comparison inputs

Prepare these three inputs from the same accepted job revision:

1. **Parts List CSV** — exported from the post-model, post-AutoBOM-review Parts
   List.
2. **SolidWorks Assembly Visualization CSV** — exported from the completed
   assembly after AutoBOM has updated model properties.
3. **Linear nest folder** — containing every final 1D Cutting Optimizer CSV/TXT
   export for the job.

Do not mix an old Parts List, a newer assembly export, and older nest files. A
comparison is meaningful only when all three represent the same revision.

## 10. Run the final production comparison

From the repository root, run:

```powershell
py data-tools/production-comparison/compare_production_parts.py `
  --nests "C:\Job\Nests" `
  --parts "C:\Job\parts.csv" `
  --solidworks "C:\Job\solidworks.csv" `
  --output "C:\Job\Reports"
```

Open `production_part_comparison.xlsx` and begin with **Errors Requiring
Action**. Use the HTML or CSV outputs when sharing or filtering the same audit.

The detailed input requirements and report meanings are documented in
[`data-tools/production-comparison/production_part_comparison_README.md`](data-tools/production-comparison/production_part_comparison_README.md).

## 11. Correct, re-export, and repeat until accepted

The comparison is a quality gate, not merely a report to file away.

For every actionable discrepancy:

1. Identify whether the authoritative correction belongs in the Parts List,
   SolidWorks model, cut file, or nest.
2. Make the correction in that source.
3. Repeat any affected downstream stage. For example, a changed modeled shape
   length requires a Parts List update, a new nest, and new exports.
4. Re-export all affected comparison inputs.
5. Run the comparison again.

The workflow is complete only when every missing part, quantity mismatch,
material mismatch, length/shape mismatch, and nesting discrepancy is either
corrected or deliberately documented and approved.

## Final job record

Retain together:

- The source A-BOM and initial Parts List.
- The accepted post-model Parts List.
- AutoCAD orchestrator logs and processed cut-file archive.
- Approved SolidWorks parts and completed assembly.
- AutoBOM run notes, including any skipped parts.
- Final 1D Cutting Optimizer exports.
- Final comparison workbook, HTML report, and CSV detail files.
- Notes for every approved exception.
