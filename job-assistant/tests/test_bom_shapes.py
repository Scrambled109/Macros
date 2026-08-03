"""Regression tests for structural-family shape preservation."""

import importlib.util
import sys
import unittest
from pathlib import Path

MODULE_PATH = (
    Path(__file__).resolve().parents[2]
    / "data-tools"
    / "bom-converter"
    / "bom_converter.py"
)
SPEC = importlib.util.spec_from_file_location("bom_converter", MODULE_PATH)
bom_converter = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = bom_converter
SPEC.loader.exec_module(bom_converter)


class StructuralShapeTests(unittest.TestCase):
    def test_tee_shape_is_preserved(self):
        self.assertEqual(
            bom_converter.parse_pbom_material_description(
                "TEE,STL,5.000D X2.690W X4.50#,STRL"
            ),
            ("TEE", "5.000DX2.690WX4.50#", "STL"),
        )

    def test_common_structural_abbreviations(self):
        cases = {
            "L,STL,4X4X1/2": ("ANGLE", "4X4X1/2", "STL"),
            "C,STL,10X20": ("CHANNEL", "10X20", "STL"),
            "W,STL,12X26": ("WIDE FLANGE", "12X26", "STL"),
            "WT,STL,6X15": ("TEE", "6X15", "STL"),
        }
        for source, expected in cases.items():
            with self.subTest(source=source):
                self.assertEqual(
                    bom_converter.parse_pbom_material_description(source), expected
                )

    def test_shape_heading_maps_to_thickness_shape(self):
        headers = [{"name": "THICKNESS/SHAPE"}]
        self.assertEqual(
            bom_converter.suggest_mapping("SHAPE", headers), "THICKNESS/SHAPE"
        )
        for heading in ("THICKNESS SHAPE", "THICKNESS/SHAPE", "STOCK SIZE"):
            self.assertEqual(
                bom_converter.suggest_mapping(heading, headers), "THICKNESS/SHAPE"
            )


if __name__ == "__main__":
    unittest.main()
