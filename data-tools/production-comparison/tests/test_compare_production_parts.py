from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


MODULE_PATH = Path(__file__).resolve().parents[1] / "compare_production_parts.py"
SPEC = importlib.util.spec_from_file_location("production_comparison_tests", MODULE_PATH)
comparison = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
sys.modules[SPEC.name] = comparison
SPEC.loader.exec_module(comparison)


class ProductionComparisonTests(unittest.TestCase):
    def test_parts_list_hash_and_solidworks_configuration_normalize_identically(self):
        rules = comparison.DEFAULT_RULES

        parts = comparison.normalize_part_number(
            "DS10039-5505A-001#705X", "Parts List", rules
        )
        solidworks = comparison.normalize_part_number(
            "DS10039-5505A-001-705X(Default<As Machined>)",
            "SolidWorks",
            rules,
        )
        nest = comparison.normalize_part_number(
            "DS10039-5505A-001-705X", "Nest", rules
        )

        self.assertEqual(parts, "DS10039-5505A-001-705X")
        self.assertEqual(parts, solidworks)
        self.assertEqual(parts, nest)

    def test_small_end_to_end_run_reconciles_all_three_sources(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            nests = root / "nests"
            output = root / "output"
            nests.mkdir()
            parts = root / "parts.csv"
            solidworks = root / "solidworks.csv"
            (nests / "angle.csv").write_text(
                "AN3x3x.25_AH-36\n\n"
                "Bar;Length;Quantity;Cost/;%\n"
                "B1;480;1;0;\n\n"
                "Parts;Length;Quantity;Left End;Right End\n"
                "DS5505A-1-7026;39.75;1;\n\n"
                "Cut thickness;Bar ends Trim\n",
                encoding="utf-8",
            )
            parts.write_text(
                "PART NUMBER,DESCRIPTION,TOTAL QUANTITY,THICKNESS/SHAPE,LENGTH,MATERIAL\n"
                "DS5505A-1#7026,ANGLE,1,AN3x3x.25,39.75,AH-36\n",
                encoding="utf-8",
            )
            solidworks.write_text(
                "FILE NAME,QUANTITY,DESCRIPTION,SHAPE,LENGTH,MATERIAL\n"
                "DS5505A-1-7026(Default<As Machined>),1,ANGLE,AN3x3x.25,39.75,AH-36\n",
                encoding="utf-8",
            )
            arguments = [
                str(MODULE_PATH),
                "--nests",
                str(nests),
                "--parts",
                str(parts),
                "--solidworks",
                str(solidworks),
                "--output",
                str(output),
                "--no-open",
            ]

            with mock.patch.object(sys, "argv", arguments):
                self.assertEqual(comparison.main(), 0)

            summary_path = next(output.rglob("comparison_summary.json"))
            summary = json.loads(summary_path.read_text(encoding="utf-8"))
            self.assertEqual(summary["counts"]["part_numbers"], 1)
            self.assertEqual(summary["counts"]["missing_core"], 0)
            self.assertEqual(summary["counts"]["exact_matches"], 1)

    def test_no_open_error_path_never_prompts_for_console_input(self):
        result = subprocess.run(
            [
                sys.executable,
                str(MODULE_PATH),
                "--nests",
                "missing-nests",
                "--parts",
                "missing-parts.csv",
                "--solidworks",
                "missing-sw.csv",
                "--output",
                "missing-output",
                "--no-open",
            ],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )

        self.assertEqual(result.returncode, 1)
        self.assertNotIn("Press Enter to close", result.stdout + result.stderr)


if __name__ == "__main__":
    unittest.main()
