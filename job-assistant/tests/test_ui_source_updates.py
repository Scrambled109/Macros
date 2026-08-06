from __future__ import annotations

import ast
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]


class UserInterfaceSourceTests(unittest.TestCase):
    def test_both_color_mappings_send_gray_8_to_plot(self) -> None:
        paths = (
            ROOT / "autocad" / "dxf-orchestrator" / "ColortoLayer.lsp",
            ROOT / "autocad" / "commands" / "ColortoLayer.lsp",
        )
        for path in paths:
            with self.subTest(path=path):
                self.assertIn('(8 . "PLOT")', path.read_text(encoding="utf-8"))

    def test_all_python_interfaces_use_steel_america_blue(self) -> None:
        paths = (
            ROOT / "job-assistant" / "job_assistant.py",
            ROOT / "data-tools" / "bom-converter" / "bom_converter.py",
            ROOT
            / "data-tools"
            / "production-comparison"
            / "compare_production_parts.py",
            ROOT
            / "data-tools"
            / "epl_converter"
            / "epl_converter"
            / "ui.py",
            ROOT / "solidworks" / "cutfile-exporter" / "cutfile_exporter.py",
        )
        for path in paths:
            with self.subTest(path=path):
                source = path.read_text(encoding="utf-8")
                ast.parse(source)
                self.assertIn("#57a0d3", source.lower())

    def test_dark_surfaces_use_intentional_backgrounds(self) -> None:
        paths = (
            ROOT / "job-assistant" / "job_assistant.py",
            ROOT / "data-tools" / "bom-converter" / "bom_converter.py",
            ROOT
            / "data-tools"
            / "production-comparison"
            / "compare_production_parts.py",
            ROOT
            / "data-tools"
            / "epl_converter"
            / "epl_converter"
            / "ui.py",
            ROOT / "solidworks" / "cutfile-exporter" / "cutfile_exporter.py",
        )
        for path in paths:
            with self.subTest(path=path):
                source = path.read_text(encoding="utf-8").lower()
                if path.name == "bom_converter.py":
                    self.assertIn("#151d24", source)
                else:
                    self.assertIn("#070b0f", source)

    def test_bom_mapping_rows_support_hover_and_selection(self) -> None:
        source = (
            ROOT / "data-tools" / "bom-converter" / "bom_converter.py"
        ).read_text(encoding="utf-8")
        self.assertIn("MappingActive.TFrame", source)
        self.assertIn('widget.bind("<Enter>"', source)
        self.assertIn("select_mapping_row", source)

    def test_assistant_references_packaged_logo(self) -> None:
        source = (ROOT / "job-assistant" / "job_assistant.py").read_text(
            encoding="utf-8"
        )
        self.assertIn('"assets" / "steel_america_logo.png"', source)

    def test_bom_theme_startup_ignores_unsupported_ttk_options(self) -> None:
        source = (
            ROOT / "data-tools" / "bom-converter" / "bom_converter.py"
        ).read_text(encoding="utf-8")
        self.assertIn("def safe_style_configure", source)
        self.assertIn("traceback.print_exc()", source)


if __name__ == "__main__":
    unittest.main()
