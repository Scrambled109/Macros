from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from run_macro import (  # noqa: E402
    SolidWorksRunnerError,
    connect_solidworks,
    run_macro,
)


class _FakeSolidWorks:
    def __init__(self) -> None:
        self.Visible = True
        self.document_visibility: list[tuple[bool, int]] = []

    def DocumentVisible(self, visible: bool, document_type: int) -> None:
        self.document_visibility.append((visible, document_type))


class SolidWorksRunnerTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.macro = Path(self.temp.name) / "Main.RunBatch.swp"
        self.macro.write_bytes(b"test macro placeholder")

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_reuses_active_session_without_starting_another(self) -> None:
        app = object()

        def unexpected_dispatch():
            self.fail("Dispatch must not run when an active session exists")

        connection = connect_solidworks(
            get_active=lambda: app,
            dispatch=unexpected_dispatch,
        )

        self.assertIs(connection.app, app)
        self.assertTrue(connection.reused)
        self.assertEqual(connection.launch_method, "active COM session")

    def test_dispatches_once_when_no_session_or_explicit_executable(self) -> None:
        app = object()
        calls = 0

        def dispatch():
            nonlocal calls
            calls += 1
            return app

        connection = connect_solidworks(
            get_active=lambda: None,
            dispatch=dispatch,
        )

        self.assertIs(connection.app, app)
        self.assertFalse(connection.reused)
        self.assertEqual(calls, 1)

    def test_hides_graphics_during_macro_and_restores_afterward(self) -> None:
        app = _FakeSolidWorks()
        observed = {}

        def invoke(current, macro, module, procedure):
            observed.update(
                visible=current.Visible,
                macro=macro,
                module=module,
                procedure=procedure,
            )
            return True, 0

        run_macro(app, self.macro, invoke=invoke)

        self.assertFalse(observed["visible"])
        self.assertEqual(observed["macro"], self.macro.resolve())
        self.assertEqual(observed["module"], "CADBatch")
        self.assertEqual(observed["procedure"], "main")
        self.assertEqual(app.document_visibility, [(False, 1), (True, 1)])
        self.assertTrue(app.Visible)

    def test_macro_failure_still_restores_solidworks_ui(self) -> None:
        app = _FakeSolidWorks()

        with self.assertRaisesRegex(SolidWorksRunnerError, "error code 17"):
            run_macro(
                app,
                self.macro,
                invoke=lambda *_args: (False, 17),
            )

        self.assertEqual(app.document_visibility, [(False, 1), (True, 1)])
        self.assertTrue(app.Visible)

    def test_rejects_non_compiled_macro(self) -> None:
        source = Path(self.temp.name) / "Main.bas"
        source.write_text("Sub main(): End Sub", encoding="ascii")

        with self.assertRaisesRegex(SolidWorksRunnerError, "compiled .swp"):
            run_macro(_FakeSolidWorks(), source, invoke=lambda *_args: (True, 0))

    def test_text_marking_suppresses_per_segment_display_updates(self) -> None:
        source = (
            Path(__file__).resolve().parents[1] / "TextMarking.bas"
        ).read_text(encoding="utf-8")

        self.assertIn("skm.AddToDB = True", source)
        self.assertIn("skm.DisplayWhenAdded = False", source)
        self.assertIn(
            "skm.DisplayWhenAdded = prevDisplayWhenAdded",
            source,
        )


if __name__ == "__main__":
    unittest.main()
