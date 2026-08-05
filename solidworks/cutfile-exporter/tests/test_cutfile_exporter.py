from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from cutfile_exporter import discover_parts  # noqa: E402


class CutfileExporterTests(unittest.TestCase):
    def test_discovery_ignores_solidworks_lock_files(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            real_part = folder / "PART.SLDPRT"
            real_part.touch()
            (folder / "~$PART.SLDPRT").touch()

            result = discover_parts(folder, recursive=False)

            self.assertEqual(result, [real_part])


if __name__ == "__main__":
    unittest.main()
