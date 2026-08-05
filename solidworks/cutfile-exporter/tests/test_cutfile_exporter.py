from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import cutfile_exporter  # noqa: E402
from cutfile_exporter import (  # noqa: E402
    ExportRecord,
    _reject_duplicate_output_names,
    discover_parts,
    export_many,
)
from solidworks_adapter import SolidWorksExportError  # noqa: E402


class CutfileExporterTests(unittest.TestCase):
    def test_discovery_ignores_solidworks_lock_files(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            real_part = folder / "PART.SLDPRT"
            real_part.touch()
            (folder / "~$PART.SLDPRT").touch()

            result = discover_parts(folder, recursive=False)

            self.assertEqual(result, [real_part])

    def test_rejects_same_named_outputs_from_different_folders(self):
        sources = [
            Path("first") / "PART.SLDPRT",
            Path("second") / "part.sldprt",
        ]

        with self.assertRaisesRegex(SolidWorksExportError, "same DXF filename"):
            _reject_duplicate_output_names(sources)

    def test_export_many_connects_once_and_processes_serially(self):
        sources = [Path("first.SLDPRT"), Path("second.SLDPRT")]
        session = object()
        records = [
            ExportRecord(str(source), f"{source.stem}.dxf", "OK")
            for source in sources
        ]
        with tempfile.TemporaryDirectory() as directory, mock.patch.object(
            cutfile_exporter,
            "_initialize_com_thread",
            return_value=None,
        ), mock.patch.object(
            cutfile_exporter.SolidWorksSession,
            "connect",
            return_value=session,
        ) as connect, mock.patch.object(
            cutfile_exporter,
            "export_part",
            side_effect=records,
        ) as export_part, mock.patch.object(
            cutfile_exporter,
            "write_report",
        ):
            output = Path(directory)
            actual = export_many(sources, output)

        self.assertEqual(actual, records)
        connect.assert_called_once_with(visible=True)
        self.assertEqual(
            [call.args[1] for call in export_part.call_args_list],
            sources,
        )

    def test_gui_runs_solidworks_batch_off_the_tk_event_thread(self):
        source = Path(cutfile_exporter.__file__).read_text(encoding="utf-8")
        self.assertIn(
            "threading.Thread(target=worker, daemon=True).start()",
            source,
        )
        self.assertIn("root.after(50, poll_worker)", source)


if __name__ == "__main__":
    unittest.main()
