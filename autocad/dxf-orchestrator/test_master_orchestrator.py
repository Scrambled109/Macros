import csv
import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest import mock


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

    def test_bevel_detection_recognizes_each_legacy_flag(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            dxf = Path(temp_dir) / "part.dxf"
            for note in ("BEVEL", "K", "V", "weld v both sides"):
                with self.subTest(note=note):
                    dxf.write_text(f"0\nTEXT\n1\n{note}\n", encoding="ascii")
                    self.assertTrue(orchestrator.is_beveled_dxf(dxf))

    def test_bevel_detection_recognizes_angle_annotations(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            dxf = Path(temp_dir) / "part.dxf"
            for note in ("V22.5", "RV9", "K 30", "V - 22.5", "RV/9"):
                with self.subTest(note=note):
                    dxf.write_text(f"0\nTEXT\n1\n{note}\n", encoding="ascii")
                    self.assertTrue(orchestrator.is_beveled_dxf(dxf))

    def test_bevel_detection_combines_mtext_chunks(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            dxf = Path(temp_dir) / "part.dxf"
            dxf.write_text("0\nMTEXT\n3\nRV\n1\n9\n0\nLINE\n", encoding="ascii")
            self.assertTrue(orchestrator.is_beveled_dxf(dxf))

    def test_similar_non_bevel_text_is_not_flagged(self):
        for note in ("EVERY", "VERIFY", "MARK", "STR-36"):
            with self.subTest(note=note):
                self.assertFalse(orchestrator.has_bevel_annotation(note))

    def test_snipe_synonyms_and_arrows_are_manual_review_flags(self):
        for note in (
            "SNIPE",
            "SNIPED END",
            "BVL OTHER SIDE",
            "CHAMFER",
            "BACK GOUGE",
            "APPLY SETS -> OTHER SIDE",
            "HOLD EDGE ↓",
        ):
            with self.subTest(note=note):
                self.assertTrue(orchestrator.has_bevel_annotation(note))

    def test_leader_entities_are_sent_to_manual_review(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            for entity in ("LEADER", "MLEADER"):
                with self.subTest(entity=entity):
                    dxf = Path(temp_dir) / f"{entity}.dxf"
                    dxf.write_text(f"0\n{entity}\n10\n1.0\n", encoding="ascii")
                    self.assertTrue(orchestrator.is_beveled_dxf(dxf))

    def test_write_script_has_one_autocad_enter_per_line(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            script = Path(temp_dir) / "job.scr"
            orchestrator.write_script(script, ["FILEDIA", "0", "_QUIT"])
            self.assertEqual(
                script.read_bytes(),
                b"FILEDIA\r\n0\r\n_QUIT\r\n",
            )
            self.assertNotIn(b"\r\r\n", script.read_bytes())

    @mock.patch.object(orchestrator.subprocess, "Popen")
    @mock.patch.object(orchestrator.subprocess, "run")
    def test_review_reuses_running_autocad(self, run, popen):
        run.return_value.returncode = 0

        opened = orchestrator.open_review_drawing(
            Path("acad.exe"), Path("part.dxf"), Path("review.scr")
        )

        self.assertTrue(opened)
        popen.assert_not_called()
        command = run.call_args.args[0]
        self.assertEqual(command[0], "powershell.exe")
        self.assertIn("-EncodedCommand", command)

    @mock.patch.object(orchestrator.subprocess, "run")
    def test_review_start_is_managed_by_powershell(self, run):
        run.return_value.returncode = 0

        opened = orchestrator.open_review_drawing(
            Path("acad.exe"), Path("part.dxf"), Path("review.scr")
        )

        self.assertTrue(opened)
        encoded = run.call_args.args[0][-1]
        powershell = orchestrator.base64.b64decode(encoded).decode("utf-16le")
        self.assertIn("Start-Process -FilePath", powershell)
        self.assertNotIn("/b", powershell)
        self.assertLess(powershell.index("Start-Process"), powershell.index("Documents.Open"))

    @mock.patch.object(orchestrator.subprocess, "Popen")
    @mock.patch.object(orchestrator.subprocess, "run")
    def test_review_does_not_start_second_autocad(self, run, popen):
        run.return_value.returncode = 2

        opened = orchestrator.open_review_drawing(
            Path("acad.exe"), Path("part.dxf"), Path("review.scr")
        )

        self.assertFalse(opened)
        popen.assert_not_called()

    @mock.patch.object(orchestrator.subprocess, "Popen")
    @mock.patch.object(orchestrator.subprocess, "run")
    def test_reused_session_disables_file_dialog_before_script(self, run, popen):
        run.return_value.returncode = 0

        orchestrator.open_review_drawing(
            Path("acad.exe"), Path("part.dxf"), Path("review.scr")
        )

        encoded = run.call_args.args[0][-1]
        powershell = orchestrator.base64.b64decode(encoded).decode("utf-16le")
        self.assertIn("$document.SetVariable('FILEDIA', 0)", powershell)
        self.assertLess(powershell.index("SetVariable"), powershell.index("SendCommand"))
        self.assertIn("Get-Process -Name acad", powershell)

    def test_manual_review_uses_unambiguous_finish_command(self):
        source = MODULE_PATH.read_text(encoding="utf-8")
        self.assertIn("(defun c:SPCFINISH", source)
        self.assertIn("type SPCFINISH", source)
        self.assertNotIn("(defun c:FINISH", source)

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
