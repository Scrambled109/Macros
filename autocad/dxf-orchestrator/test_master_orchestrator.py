import csv
import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest import mock


MODULE_PATH = Path(__file__).with_name("Master_Orchestrator.py")
LAYER_SCRIPT = Path(__file__).with_name("ColortoLayer.lsp")
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

    def test_snipe_synonyms_and_arrows_are_bevel_markers(self):
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

    def test_leader_entities_are_bevel_markers(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            for entity in ("LEADER", "MLEADER", "MULTILEADER"):
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

    def test_output_details_builds_expected_sorted_path(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            dxf = Path(temp_dir) / "P#1.dxf"
            parts = {
                "P#1": {"quantity": "2", "thickness": "3/8", "material": "A36"}
            }
            target, output = orchestrator.output_details(
                dxf, parts, Path(temp_dir)
            )
            self.assertEqual(target, "375-A36")
            self.assertEqual(output.name, "P-1_375-A36_2.dwg")

    def test_bevel_output_has_exact_parenthesized_suffix(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            dxf = Path(temp_dir) / "P#1.dxf"
            parts = {
                "P#1": {"quantity": "2", "thickness": "3/8", "material": "A36"}
            }
            target, output = orchestrator.output_details(
                dxf,
                parts,
                Path(temp_dir),
                beveled=True,
            )
            self.assertEqual(target, "375-A36")
            self.assertEqual(output.name, "P-1_375-A36_2(B).dwg")

    @mock.patch.object(orchestrator, "is_beveled_dxf", return_value=True)
    @mock.patch.object(orchestrator, "wait_file_stable", return_value=True)
    @mock.patch.object(orchestrator, "archive_original")
    @mock.patch.object(orchestrator.subprocess, "Popen")
    def test_bevel_uses_core_console_without_manual_review(
        self,
        popen,
        archive_original,
        _wait_file_stable,
        _is_beveled_dxf,
    ):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            dxf = root / "P1.dxf"
            dxf.write_text("0\nTEXT\n1\nBEVEL\n", encoding="ascii")
            archive = root / "archive"
            logs = root / "logs"
            archive.mkdir()
            logs.mkdir()
            script = root / "job_0001.scr"
            popen.return_value.wait.return_value = 0

            result = orchestrator.process_file(
                dxf,
                {},
                root,
                archive,
                logs,
                Path("accoreconsole.exe"),
                Path("ColorToLayer.lsp"),
                Path("seed.dwg"),
                Path("HashToDash.lsp"),
                script,
                10,
            )

            self.assertEqual(result, "bevel")
            popen.assert_called_once()
            archive_original.assert_called_once_with(dxf, archive)
            self.assertIn("P1_Unsorted_1(B).dwg", script.read_text(encoding="ascii"))

    def test_python_orchestrator_has_no_gui_review_path(self):
        source = MODULE_PATH.read_text(encoding="utf-8")
        self.assertNotIn("open_review_drawing", source)
        self.assertNotIn("SPCFINISH", source)
        self.assertNotIn("--acad-gui-path", source)
        self.assertNotIn("--review-timeout", source)

    def test_layer_script_overrides_bevel_text_and_arrows_to_plot(self):
        source = LAYER_SCRIPT.read_text(encoding="utf-8")
        self.assertIn("SPC:BevelTextP", source)
        self.assertIn('(cons 8 "PLOT")', source)
        self.assertIn('((0 . "LEADER,MLEADER,MULTILEADER"))', source)
        self.assertIn('((0 . "TEXT,MTEXT"))', source)
        self.assertIn("BEVEL(ED)?|BVL|CHAMFER", source)
        self.assertIn('"PIN STAMP LINE MARKING"', source)
        self.assertIn('"PIN STAMP TEXT"', source)

    def test_layer_script_parentheses_are_balanced(self):
        source = LAYER_SCRIPT.read_text(encoding="utf-8")
        depth = 0
        in_string = False
        in_comment = False
        escaped = False
        for char in source:
            if in_comment:
                if char == "\n":
                    in_comment = False
                continue
            if in_string:
                if escaped:
                    escaped = False
                elif char == "\\":
                    escaped = True
                elif char == '"':
                    in_string = False
                continue
            if char == '"':
                in_string = True
            elif char == ";":
                in_comment = True
            elif char == "(":
                depth += 1
            elif char == ")":
                depth -= 1
                self.assertGreaterEqual(depth, 0)
        self.assertFalse(in_string)
        self.assertEqual(depth, 0)

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
        self.assertEqual(args.workers, 2)

    def test_worker_count_is_bounded(self):
        self.assertEqual(orchestrator.parse_args(["--workers", "4"]).workers, 4)
        with self.assertRaises(SystemExit):
            orchestrator.parse_args(["--workers", "0"])
        with self.assertRaises(SystemExit):
            orchestrator.parse_args(["--workers", "5"])

    @mock.patch.object(orchestrator, "process_file", return_value="clean")
    def test_parallel_files_receive_unique_scripts(self, process_file):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            files = [root / "001" / "P1.dxf", root / "002" / "P2.dxf"]
            for path in files:
                path.parent.mkdir(exist_ok=True)
                path.write_text("0\nEOF\n", encoding="ascii")
            archive = root / "archive"
            logs = root / "logs"
            scripts = root / "scripts"
            archive.mkdir()
            logs.mkdir()
            scripts.mkdir()

            counts = orchestrator.process_files(
                files,
                parts={},
                workspace=root,
                archive_dir=archive,
                log_dir=logs,
                acad_console=Path("accoreconsole.exe"),
                lsp_path=Path("ColorToLayer.lsp"),
                seed_path=Path("seed.dwg"),
                h2d_path=Path("HashToDash.lsp"),
                temp_dir=scripts,
                workers=2,
                console_timeout=10,
            )

            self.assertEqual(counts, {"clean": 2, "bevel": 0, "failed": 0})
            script_paths = [call.args[9] for call in process_file.call_args_list]
            self.assertEqual(len(set(script_paths)), 2)

    @mock.patch.object(orchestrator, "process_file", return_value="clean")
    def test_duplicate_output_paths_are_rejected_before_autocad(self, process_file):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            files = [root / "001" / "P1.dxf", root / "002" / "P1.dxf"]
            for path in files:
                path.parent.mkdir(exist_ok=True)
                path.write_text("0\nEOF\n", encoding="ascii")
            for name in ("archive", "logs", "scripts"):
                (root / name).mkdir()

            counts = orchestrator.process_files(
                files,
                parts={},
                workspace=root,
                archive_dir=root / "archive",
                log_dir=root / "logs",
                acad_console=Path("accoreconsole.exe"),
                lsp_path=Path("ColorToLayer.lsp"),
                seed_path=Path("seed.dwg"),
                h2d_path=Path("HashToDash.lsp"),
                temp_dir=root / "scripts",
                workers=2,
                console_timeout=10,
            )

            self.assertEqual(counts["failed"], 2)
            process_file.assert_not_called()

    def test_bevel_marker_keeps_standard_and_bevel_outputs_distinct(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            standard = root / "001" / "P1.dxf"
            beveled = root / "002" / "P1.dxf"
            standard.parent.mkdir()
            beveled.parent.mkdir()
            standard.write_text("0\nTEXT\n1\nPLAIN\n", encoding="ascii")
            beveled.write_text("0\nTEXT\n1\nBEVEL\n", encoding="ascii")

            safe, collisions = orchestrator.split_output_collisions(
                [standard, beveled],
                {},
                root,
            )

            self.assertEqual(safe, [standard, beveled])
            self.assertEqual(collisions, [])


if __name__ == "__main__":
    unittest.main()
