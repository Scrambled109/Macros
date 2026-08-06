from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE = (ROOT / "ModifiedPartsListProperties.bas").read_text(encoding="utf-8")
FORM = (ROOT / "ModifiedPartsListMapper.frm").read_text(encoding="utf-8")


class ModifiedPartsListSourceTests(unittest.TestCase):
    def test_active_excel_falls_back_to_workbook_picker(self) -> None:
        self.assertIn('GetObject(, "Excel.Application")', MODULE)
        self.assertIn("GetOpenFilename", MODULE)
        self.assertLess(
            MODULE.index('GetObject(, "Excel.Application")'),
            MODULE.index("GetOpenFilename"),
        )

    def test_mapping_form_is_shown_every_run(self) -> None:
        self.assertIn("New ModifiedPartsListMapper", MODULE)
        self.assertIn("mapper.Show vbModal", MODULE)
        self.assertIn("Begin VB.ComboBox cboPartNumber", FORM)
        self.assertIn("Begin VB.ComboBox cboDescription", FORM)
        self.assertIn("Begin VB.ComboBox cboRawMaterial", FORM)

    def test_selection_falls_back_to_whole_assembly(self) -> None:
        selection = MODULE.index("GetSelectedObjectCount2")
        entire_assembly = MODULE.index("swAssembly.GetComponents(False)")
        self.assertLess(selection, entire_assembly)
        self.assertIn("If result.Count > 0 Then", MODULE)

    def test_properties_target_global_and_referenced_configuration(self) -> None:
        self.assertIn('WriteMappedProperties model, ""', MODULE)
        self.assertIn("component.ReferencedConfiguration", MODULE)
        self.assertIn('manager.Add3 "Description"', MODULE)
        self.assertIn('manager.Add3 "Raw_Material"', MODULE)

    def test_ui_and_excel_ownership_are_restored(self) -> None:
        self.assertIn("RestoreAssemblyUI", MODULE)
        self.assertIn("If workbookWasOpened", MODULE)
        self.assertIn("If excelWasCreated", MODULE)


if __name__ == "__main__":
    unittest.main()
