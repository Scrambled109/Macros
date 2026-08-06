# Typical Production Workflow

Use this order for a complete job. Do not mark a stage complete until its output
has been reviewed.

1. **Create the Parts List.** Run the Job Assistant or
   `data-tools/bom-converter/bom_converter.py`. Check part numbers, quantities,
   descriptions, material, and thickness.
2. **Prepare cut files.** Use the Job Assistant DXF stage or
   `autocad/dxf-orchestrator/Master_Orchestrator.py`. Bevel drawings are
   processed headlessly and receive a `(B)` filename suffix. Review every output
   and the orchestrator logs.
3. **Create plate models.** In the Job Assistant, select one reviewed
   material/thickness folder. Then run
   `solidworks/cad-batch-converter/Main.RunBatch.swp` from SolidWorks. Confirm
   the requested thickness and each saved part in `BatchLog.txt`.
4. **Model shapes and beveled parts manually.** The automation does not replace
   engineering judgment for these parts.
5. **Assemble the complete SolidWorks model.** Resolve missing, duplicate, or
   incorrectly configured parts before continuing.
6. **Apply modified Parts List properties.** Work on a recoverable model copy,
   then run `solidworks/auto-bom/AutoBOMProperties.swp`. It uses the active Excel
   workbook or asks for one, shows column dropdowns every run, and processes
   selected components—or the entire assembly when none are selected. Review
   unmatched, skipped, duplicate, and save-failed parts.
7. **Update the Parts List.** Reconcile model-derived length and shape values.
8. **Create the linear nests.** Export the nest data required by the comparison
   tool.
9. **Run production comparison.** Use
   `data-tools/production-comparison/compare_production_parts.py` and correct
   every item on **Errors Requiring Action**.

At every handoff, keep the job number and part number consistent and retain a
dated or revision-controlled copy of the previous approved output.

