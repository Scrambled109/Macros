"""Run the compiled CAD batch macro through one reusable SolidWorks session.

The Engineering Job Assistant starts this helper as a separate process so the
Tkinter dashboard remains responsive while SolidWorks is busy.  The helper
attaches to the running COM server first, starts SolidWorks only when needed,
hides application/document graphics for the batch, and restores the UI when
the macro returns.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import os
from pathlib import Path
import subprocess
import sys
import time
import traceback
from typing import Callable


SOLIDWORKS_PROGIDS = ("SldWorks.Application.33", "SldWorks.Application")
SW_DOC_PART = 1
SW_RUN_MACRO_UNLOAD_AFTER_RUN = 1


class SolidWorksRunnerError(RuntimeError):
    """A SolidWorks startup or macro error that can be shown to the operator."""


@dataclass(frozen=True)
class SolidWorksConnection:
    app: object
    reused: bool
    launch_method: str


def _get_active_solidworks():
    import win32com.client

    for prog_id in SOLIDWORKS_PROGIDS:
        try:
            return win32com.client.GetActiveObject(prog_id)
        except Exception:
            continue
    return None


def _dispatch_solidworks():
    import win32com.client

    failures: list[str] = []
    for prog_id in SOLIDWORKS_PROGIDS:
        try:
            return win32com.client.Dispatch(prog_id)
        except Exception as exc:
            failures.append(f"{prog_id}: {exc}")
    raise SolidWorksRunnerError(
        "Could not start SolidWorks through COM. " + "; ".join(failures)
    )


def connect_solidworks(
    executable: Path | None = None,
    *,
    timeout: float = 90.0,
    get_active: Callable[[], object | None] | None = None,
    dispatch: Callable[[], object] | None = None,
    launch: Callable[..., object] | None = None,
    monotonic: Callable[[], float] | None = None,
    sleep: Callable[[float], None] | None = None,
) -> SolidWorksConnection:
    """Reuse the active instance, otherwise start exactly one SolidWorks app."""

    get_active = get_active or _get_active_solidworks
    dispatch = dispatch or _dispatch_solidworks
    launch = launch or subprocess.Popen
    monotonic = monotonic or time.monotonic
    sleep = sleep or time.sleep

    app = get_active()
    if app is not None:
        return SolidWorksConnection(app, reused=True, launch_method="active COM session")

    if executable is None:
        return SolidWorksConnection(
            dispatch(), reused=False, launch_method="registered COM server"
        )

    executable = Path(executable).resolve()
    if not executable.is_file():
        raise SolidWorksRunnerError(
            f"Configured SolidWorks executable was not found: {executable}"
        )
    launch(
        [str(executable)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    deadline = monotonic() + max(timeout, 1.0)
    while monotonic() < deadline:
        app = get_active()
        if app is not None:
            return SolidWorksConnection(
                app, reused=False, launch_method="configured executable"
            )
        sleep(0.5)
    raise SolidWorksRunnerError(
        f"SolidWorks started but did not register its COM server within {timeout:g} seconds."
    )


def _invoke_macro(
    app,
    macro: Path,
    module: str,
    procedure: str,
) -> tuple[bool, int]:
    import pythoncom
    from win32com.client import VARIANT

    error = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    succeeded = app.RunMacro2(
        str(macro),
        module,
        procedure,
        SW_RUN_MACRO_UNLOAD_AFTER_RUN,
        error,
    )
    return bool(succeeded), int(error.value)


def run_macro(
    app,
    macro: Path,
    *,
    module: str = "CADBatch",
    procedure: str = "main",
    show_after_run: bool = True,
    invoke: Callable[[object, Path, str, str], tuple[bool, int]] | None = None,
) -> None:
    """Run a compiled SWP with graphics suppressed and restore the UI safely."""

    macro = Path(macro).resolve()
    if not macro.is_file():
        raise SolidWorksRunnerError(f"SolidWorks macro was not found: {macro}")
    if macro.suffix.casefold() != ".swp":
        raise SolidWorksRunnerError(f"Expected a compiled .swp macro: {macro}")

    invoke = invoke or _invoke_macro
    document_windows_hidden = False
    try:
        try:
            app.Visible = False
        except Exception as exc:
            print(f"WARNING: could not hide the SolidWorks application: {exc}")
        try:
            app.DocumentVisible(False, SW_DOC_PART)
            document_windows_hidden = True
        except Exception as exc:
            print(f"WARNING: could not suppress part windows: {exc}")

        succeeded, error_code = invoke(app, macro, module, procedure)
        if not succeeded:
            raise SolidWorksRunnerError(
                f"SolidWorks RunMacro2 failed with error code {error_code}."
            )
    finally:
        if document_windows_hidden:
            try:
                app.DocumentVisible(True, SW_DOC_PART)
            except Exception as exc:
                print(f"WARNING: could not restore part-window visibility: {exc}")
        if show_after_run:
            try:
                app.Visible = True
            except Exception as exc:
                print(f"WARNING: could not restore the SolidWorks window: {exc}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the CAD batch converter in a reusable background SolidWorks session."
    )
    parser.add_argument("--macro", required=True, type=Path)
    parser.add_argument("--module", default="CADBatch")
    parser.add_argument("--procedure", default="main")
    parser.add_argument("--solidworks-executable", type=Path)
    parser.add_argument("--connect-timeout", type=float, default=90.0)
    parser.add_argument(
        "--leave-hidden",
        action="store_true",
        help="Keep SolidWorks hidden after the macro instead of restoring it for review.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if os.name != "nt":
        print("ERROR: SolidWorks automation requires Windows.", file=sys.stderr)
        return 2

    try:
        import pythoncom
    except ImportError:
        print(
            "ERROR: pywin32 is not installed. Run: py -m pip install -r requirements.txt",
            file=sys.stderr,
        )
        return 2

    pythoncom.CoInitialize()
    try:
        connection = connect_solidworks(
            args.solidworks_executable,
            timeout=args.connect_timeout,
        )
        action = "Reusing" if connection.reused else "Started"
        print(f"{action} SolidWorks via {connection.launch_method}.", flush=True)
        print(f"Running {args.macro}::{args.module}.{args.procedure}", flush=True)
        run_macro(
            connection.app,
            args.macro,
            module=args.module,
            procedure=args.procedure,
            show_after_run=not args.leave_hidden,
        )
        print("SolidWorks macro completed and returned control.", flush=True)
        return 0
    except SolidWorksRunnerError as exc:
        print(f"ERROR: {exc}", file=sys.stderr, flush=True)
        return 1
    except Exception as exc:
        print(f"ERROR: unexpected SolidWorks automation failure: {exc}", file=sys.stderr)
        traceback.print_exc()
        return 1
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
