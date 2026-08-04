import csv
import importlib.util
import tempfile
import unittest
from pathlib import Path


MODULE_PATH = Path(__file__).with_name("Master_Orchestrator.py")
SPEC = importlib.util.spec_from_file_location("master_orchestrator", MODULE_PATH)
orchestrator = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(orchestrator)


class OrchestratorHelpersTest(unittest.TestCase):
    def test_thickness_to_mils(self):
        self.assertEqual(orchestrator.thickness_to_mils("0.5"), "500")
        self.assertEqual(orchestrator.thickness_to_mils("3/8"), "375")
        self.assertEqual(orchestrator.thickness_to_mils("1-1/2"), "1500")
        self.assertEqual(orchestrator.thickness_to_mils("250"), "250")

    def test_safe_name(self):
        self.assertEqual(orchestrator.safe_name('a/b:c*?"d<e>|f'), "a-b-c---d-e--f")

    def test_bevel_detection_only_checks_text_entities(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            dxf = Path(temp_dir) / "part.dxf"
            dxf.write_text("0\nSTYLE\n1\nK\n0\nTEXT\n1\nplain\n", encoding="ascii")
            self.assertFalse(orchestrator.is_beveled_dxf(dxf))
            dxf.write_text("0\nMTEXT\n3\nWELD K BOTH SIDES\n", encoding="ascii")
            self.assertTrue(orchestrator.is_beveled_dxf(dxf))

    def test_load_parts_accepts_legacy_quanity_header(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            csv_path = Path(temp_dir) / "parts.csv"
            with csv_path.open("w", newline="", encoding="utf-8") as handle:
                writer = csv.DictWriter(
                    handle,
                    fieldnames=["PartNumber", "Quanity", "Thickness", "Material"],
                )
                writer.writeheader()
                writer.writerow(
                    {"PartNumber": "P1", "Quanity": "2", "Thickness": "1/2", "Material": "A36"}
                )
            parts = orchestrator.load_parts(csv_path)
            self.assertEqual(parts["P1"]["quantity"], "2")

    def test_parts_list_argument_defaults_to_workspace_file(self):
        args = orchestrator.parse_args(["--workspace", "job workspace"])
        self.assertIsNone(args.parts_list)


if __name__ == "__main__":
    unittest.main()
