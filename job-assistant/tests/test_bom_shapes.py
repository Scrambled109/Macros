"""Regression tests for structural-family shape preservation."""

import importlib.util
import sys
import unittest
from pathlib import Path

import pandas as pd

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
            ("TEE", "5DX2.69WX4.5#", "STL"),
        )

    def test_built_up_tee_uses_complete_shape_not_plate_thickness(self):
        self.assertEqual(
            bom_converter.parse_pbom_material_description(
                "TEE,STL,7.125 X5.000 X9.7#,BUILT-UP"
            ),
            ("TEE", "7.125X5X9.7#", "STL"),
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

    def test_catia_name_spec_extracts_tee_shape_and_material(self):
        self.assertEqual(
            bom_converter.parse_catia_name_spec(
                "G5505D-13,BT7.125X5.0X9.7,,HSLA-65,"
            ),
            ("7.125X5X9.7", "HSLA-65"),
        )

    def test_catia_plate_ignores_weight_and_converts_thickness(self):
        self.assertEqual(
            bom_converter.parse_catia_name_spec(
                "G5505D-12,Plate,0.313,HSLA-65,Weight"
            ),
            ("5/16", "HSLA-65"),
        )

    def test_selected_material_keeps_last_term_before_weight(self):
        self.assertEqual(
            bom_converter.material_from_selected_value(
                "G5505D-12,Plate,0.313,HSLA-65,Weight"
            ),
            "HSLA-65",
        )

    def test_selected_material_skips_other_trailing_metadata_labels(self):
        for label in (
            "THICKN",
            "Thickness",
            "THK",
            "Length",
            "Width",
            "Mass",
            "Description",
        ):
            with self.subTest(label=label):
                value = f"G5505D-12,Plate,0.313,HSLA-65,{label}"
                self.assertEqual(
                    bom_converter.material_from_selected_value(value),
                    "HSLA-65",
                )
                self.assertEqual(
                    bom_converter.parse_catia_name_spec(value),
                    ("5/16", "HSLA-65"),
                )

    def test_selected_material_skips_multiple_metadata_labels(self):
        value = "G5505D-12,Plate,0.313,HSLA-65,THICKN,Weight"
        self.assertEqual(
            bom_converter.material_from_selected_value(value),
            "HSLA-65",
        )

    def test_plate_weight_converts_while_shape_keeps_full_designation(self):
        self.assertEqual(
            bom_converter.parse_pbom_material_description(
                "PLATE,STL,12.75# X120.000 X312.000"
            ),
            ("PLATE", "5/16", "STL"),
        )
        self.assertEqual(
            bom_converter.parse_pbom_material_description(
                "TEE,STL,7.125 X5.000 X9.7#,BUILT-UP"
            ),
            ("TEE", "7.125X5X9.7#", "STL"),
        )

    def test_selected_material_and_pbom_shape_are_both_preserved(self):
        dataframe = pd.DataFrame([{
            "PART": "DS5505A-1#4012",
            "DESC": "TEE,STL,7.125 X5.000 X9.7#,BUILT-UP",
            "AH": "G5505D-12,Plate,0.313,HSLA-65,Weight",
            "NameSpecA from CATIA": "G5505D-12,Plate,0.313,HSLA-65,Weight",
        }])
        records = bom_converter.build_records(
            dataframe,
            [
                {"source": "PART", "destination": "PART NUMBER"},
                {"source": "DESC", "destination": "DESCRIPTION"},
                {"source": "AH", "destination": "MATERIAL TYPE"},
            ],
            ("DS",),
            template_headers=[
                {"name": "PART NUMBER"},
                {"name": "DESCRIPTION"},
                {"name": "THICKNESS/SHAPE"},
                {"name": "MATERIAL TYPE"},
            ],
        )
        self.assertEqual(records[0]["THICKNESS/SHAPE"], "7.125X5X9.7#")
        self.assertEqual(records[0]["MATERIAL TYPE"], "HSLA-65")

    def test_catia_structural_variations_do_not_become_material(self):
        cases = {
            "G5505D,W12X26,A572-50,Weight": ("W12X26", "A572-50"),
            "G5505D,HSS6X6X3/8,A500-B": ("HSS6X6X3/8", "A500-B"),
            "G5505D,Plate,0.500,Weight": ("1/2", ""),
        }
        for source, expected in cases.items():
            with self.subTest(source=source):
                self.assertEqual(
                    bom_converter.parse_catia_name_spec(source), expected
                )

    def test_catia_metadata_populates_unmapped_output_columns(self):
        dataframe = pd.DataFrame(
            [{
                "PART": "DS5505A-1#4012",
                "DESC": "TEE",
                "NameSpecA from CATIA": "G5505D-13,BT7.125X5X9.7,,HS,",
            }]
        )
        records = bom_converter.build_records(
            dataframe,
            [
                {"source": "PART", "destination": "PART NUMBER"},
                {"source": "DESC", "destination": "DESCRIPTION"},
                {
                    "source": "NameSpecA from CATIA",
                    "destination": bom_converter.IGNORE_MAPPING,
                },
            ],
            ("DS",),
            template_headers=[
                {"name": "PART NUMBER"},
                {"name": "DESCRIPTION"},
                {"name": "THICKNESS/SHAPE"},
                {"name": "MATERIAL TYPE"},
            ],
        )
        self.assertEqual(records[0]["THICKNESS/SHAPE"], "7.125X5X9.7")
        self.assertEqual(records[0]["MATERIAL TYPE"], "HS")


    def test_mapping_wheel_does_not_move_pane_from_combobox(self):
        class Widget:
            def __init__(self, name, widget_class, parent=None):
                self.name = name
                self.widget_class = widget_class
                self.parent = parent
                self.scrolls = []
                self.registry = {name: self}
                if parent is not None:
                    self.registry = parent.registry
                    self.registry[name] = self

            def winfo_exists(self):
                return True

            def winfo_class(self):
                return self.widget_class

            def winfo_parent(self):
                return self.parent.name if self.parent is not None else ""

            def _nametowidget(self, name):
                return self.registry[name]

            def yview_scroll(self, amount, units):
                self.scrolls.append((amount, units))

        canvas = Widget(".mapping", "Canvas")
        combo = Widget(".mapping.combo", "TCombobox", canvas)
        label = Widget(".mapping.label", "TLabel", canvas)
        app = object.__new__(bom_converter.BomConverterApp)
        app.mapping_canvas = canvas

        combo_event = type("Event", (), {"widget": combo, "delta": -120})()
        self.assertIsNone(app.mapping_mousewheel(combo_event))
        self.assertEqual(canvas.scrolls, [])

        label_event = type("Event", (), {"widget": label, "delta": -120})()
        self.assertEqual(app.mapping_mousewheel(label_event), "break")
        self.assertEqual(canvas.scrolls, [(1, "units")])


if __name__ == "__main__":
    unittest.main()
