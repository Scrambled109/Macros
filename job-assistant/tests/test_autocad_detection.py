"""Tests for AutoCAD Core Console version detection."""

import importlib.util
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock


MODULE_PATH = (
    Path(__file__).resolve().parents[2]
    / "autocad"
    / "dxf-orchestrator"
    / "Master_Orchestrator.py"
)
SPEC = importlib.util.spec_from_file_location("master_orchestrator", MODULE_PATH)
master_orchestrator = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = master_orchestrator
SPEC.loader.exec_module(master_orchestrator)


class AutoCADDetectionTests(unittest.TestCase):
    def _console(self, root: Path, year: int) -> Path:
        path = root / "Autodesk" / f"AutoCAD {year}" / "accoreconsole.exe"
        path.parent.mkdir(parents=True)
        path.touch()
        return path

    def test_prefers_2026_when_both_supported_versions_exist(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            expected = self._console(root, 2026)
            self._console(root, 2025)
            self.assertEqual(
                master_orchestrator.detect_autocad_console(
                    program_files=root
                ),
                expected,
            )

    def test_falls_back_to_autocad_2025(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            expected = self._console(root, 2025)
            self.assertEqual(
                master_orchestrator.detect_autocad_console(
                    program_files=root
                ),
                expected,
            )

    def test_detects_another_installed_version_after_supported_versions(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            expected = self._console(root, 2024)
            self.assertEqual(
                master_orchestrator.detect_autocad_console(
                    program_files=root
                ),
                expected,
            )

    def test_explicit_or_environment_path_overrides_detection(self):
        explicit = Path("D:/AutoCAD/accoreconsole.exe")
        self.assertEqual(
            master_orchestrator.detect_autocad_console(explicit),
            explicit,
        )
        with mock.patch.dict(
            os.environ,
            {"ACAD_CONSOLE_PATH": str(explicit)},
            clear=False,
        ):
            self.assertEqual(
                master_orchestrator.detect_autocad_console(),
                explicit,
            )


if __name__ == "__main__":
    unittest.main()
