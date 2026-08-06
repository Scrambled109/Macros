from __future__ import annotations

import sys
import tempfile
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from run_macro import (  # noqa: E402
    DISP_E_PARAMNOTOPTIONAL,
    SW_METHODS_WITHOUT_ARGUMENTS,
    SolidWorksRunnerError,
    _invoke_macro,
    connect_solidworks,
    resolve_entry_point,
    run_macro,
    wait_until_macro_ready,
)


class _FakeSolidWorks:
    def __init__(self) -> None:
        self.Visible = True
        self.UserControl = False
        self.CommandInProgress = False


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

        connection = connect_solidworks(get_active=lambda: None, dispatch=dispatch)

        self.assertIs(connection.app, app)
        self.assertFalse(connection.reused)
        self.assertEqual(calls, 1)

    def test_waits_past_early_com_registration_until_vba_is_ready(self) -> None:
        responses = [RuntimeError("server busy"), [], ["CADBatch.main"]]
        clock = iter([0.0, 0.1, 0.2, 0.3])

        def methods(_app, _macro):
            response = responses.pop(0)
            if isinstance(response, Exception):
                raise response
            return response

        available = wait_until_macro_ready(
            object(),
            self.macro,
            timeout=5,
            monotonic=lambda: next(clock),
            sleep=lambda _seconds: None,
            methods=methods,
        )

        self.assertEqual(available, ["CADBatch.main"])

    def test_macro_method_filter_matches_solidworks_enum(self) -> None:
        self.assertEqual(SW_METHODS_WITHOUT_ARGUMENTS, 1)

    def test_resolves_main_from_actual_compiled_module_name(self) -> None:
        self.assertEqual(
            resolve_entry_point(["Main_RunBatch1.Helper", "Main_RunBatch1.main"]),
            ("Main_RunBatch1", "main"),
        )

    def test_resolves_module_only_response_from_get_macro_methods(self) -> None:
        self.assertEqual(
            resolve_entry_point(["Main_RunBatch1"]),
            ("Main_RunBatch1", "main"),
        )

    def test_runs_without_forcing_document_visibility(self) -> None:
        app = _FakeSolidWorks()
        observed = {}

        def invoke(current, macro, module, procedure):
            observed.update(
                command=current.CommandInProgress,
                macro=macro,
                module=module,
                procedure=procedure,
            )
            return True, 0

        resolved = run_macro(
            app,
            self.macro,
            ready=lambda *_args, **_kwargs: ["CompiledName.main"],
            invoke=invoke,
        )

        self.assertTrue(observed["command"])
        self.assertEqual(observed["macro"], self.macro.resolve())
        self.assertEqual(observed["module"], "CompiledName")
        self.assertEqual(observed["procedure"], "main")
        self.assertEqual(resolved, ("CompiledName", "main"))
        self.assertFalse(app.CommandInProgress)
        self.assertTrue(app.Visible)
        self.assertTrue(app.UserControl)
        self.assertFalse(hasattr(app, "document_visibility"))

    def test_macro_failure_still_clears_performance_flag(self) -> None:
        app = _FakeSolidWorks()

        with self.assertRaisesRegex(SolidWorksRunnerError, "error code 17"):
            run_macro(
                app,
                self.macro,
                ready=lambda *_args, **_kwargs: ["CADBatch.main"],
                invoke=lambda *_args: (False, 17),
            )

        self.assertFalse(app.CommandInProgress)
        self.assertTrue(app.Visible)

    def test_generated_pywin32_runmacro2_tuple_is_supported(self) -> None:
        class GeneratedWrapper:
            def RunMacro2(self, *_args):
                self.args = _args
                return True, 0

        app = GeneratedWrapper()
        self.assertEqual(
            _invoke_macro(app, self.macro, "CADBatch", "main"),
            (True, 0),
        )
        self.assertEqual(len(app.args), 4)

    def test_dynamic_wrapper_retries_with_required_out_parameter(self) -> None:
        class ParameterNotOptional(Exception):
            hresult = DISP_E_PARAMNOTOPTIONAL

        class ErrorVariant:
            def __init__(self, _kind, value):
                self.value = value

        class DynamicWrapper:
            def __init__(self):
                self.calls = []

            def RunMacro2(self, *args):
                self.calls.append(args)
                if len(args) == 4:
                    raise ParameterNotOptional("Parameter not optional")
                args[-1].value = 0
                return True

        app = DynamicWrapper()
        self.assertEqual(
            _invoke_macro(
                app,
                self.macro,
                "CADBatch",
                "main",
                variant_factory=ErrorVariant,
                byref_i4=99,
            ),
            (True, 0),
        )
        self.assertEqual([len(call) for call in app.calls], [4, 5])

    def test_real_com_failure_is_not_retried(self) -> None:
        class MacroFailure(Exception):
            hresult = -123

        class FailingWrapper:
            def __init__(self):
                self.calls = 0

            def RunMacro2(self, *_args):
                self.calls += 1
                raise MacroFailure("macro failed")

        app = FailingWrapper()
        with self.assertRaisesRegex(MacroFailure, "macro failed"):
            _invoke_macro(app, self.macro, "CADBatch", "main")
        self.assertEqual(app.calls, 1)

    def test_rejects_non_compiled_macro(self) -> None:
        source = Path(self.temp.name) / "Main.bas"
        source.write_text("Sub main(): End Sub", encoding="ascii")

        with self.assertRaisesRegex(SolidWorksRunnerError, "compiled .swp"):
            run_macro(
                _FakeSolidWorks(),
                source,
                ready=lambda *_args, **_kwargs: ["CADBatch.main"],
                invoke=lambda *_args: (True, 0),
            )


if __name__ == "__main__":
    unittest.main()
