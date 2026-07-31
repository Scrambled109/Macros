from __future__ import annotations

import unittest
import sys
import types
from unittest.mock import MagicMock, patch

try:
    import customtkinter  # noqa: F401
except ModuleNotFoundError:
    fake_ctk = types.ModuleType("customtkinter")
    for class_name in ("CTk", "CTkFrame", "CTkScrollableFrame", "CTkTextbox", "CTkLabel"):
        setattr(fake_ctk, class_name, type(class_name, (), {}))
    sys.modules["customtkinter"] = fake_ctk

from epl_converter.ui import ConverterApp
from epl_converter import ConversionError


class ConverterUITests(unittest.TestCase):
    def test_bop_picker_stores_selected_files_and_updates_visible_count(self) -> None:
        app = ConverterApp.__new__(ConverterApp)
        app.bop_files = []
        app.bop_box = MagicMock()
        app.bop_count_label = MagicMock()

        selected = (r"C:\Input\BOP-1.xlsm", r"C:\Input\BOP-2.xlsx")
        with patch("epl_converter.ui.filedialog.askopenfilenames", return_value=selected):
            app._select_bops()

        self.assertEqual(app.bop_files, list(selected))
        app.bop_box.insert.assert_called_once()
        app.bop_count_label.configure.assert_called_once_with(
            text="2 BOP/BOM file(s) selected"
        )

    def test_background_conversion_error_survives_until_tk_callback_runs(self) -> None:
        app = ConverterApp.__new__(ConverterApp)
        app.epl_files = [r"C:\Input\LOA000001-EPL.xlsx"]
        app.bop_files = [r"C:\Input\BOP.xlsm"]
        app.output_dir = MagicMock()
        app.output_dir.get.return_value = r"C:\Output"
        app.output_name = MagicMock()
        app.output_name.get.return_value = "Parts List"
        app.metadata_path = MagicMock()
        app.metadata_path.get.return_value = ""
        app.plate_template_path = MagicMock()
        app.plate_template_path.get.return_value = ""
        app.after = MagicMock()
        app._finish = MagicMock()

        with patch(
            "epl_converter.ui.convert_epls",
            side_effect=ConversionError("Test conversion failure"),
        ):
            app._convert()

        delayed_callback = app.after.call_args.args[1]
        delayed_callback()
        app._finish.assert_called_once_with("", "Test conversion failure")


if __name__ == "__main__":
    unittest.main()
