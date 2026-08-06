"""Run the compiled CAD batch macro through one reusable SolidWorks session.

The runner deliberately leaves document visibility to the VBA macro. Hiding all
part documents from an out-of-process controller changes how newly imported
documents are activated and made the previous runner less stable than running
the same SWP manually.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import os
from pathlib import Path
import re
import subprocess
import sys
import time
import traceback
from typing import Callable


SOLIDWORKS_PROGIDS = ("SldWorks.Application.33", "SldWorks.Application")
SW_RUN_MACRO_UNLOAD_AFTER_RUN = 1
# swMacroMethods_e.swMethodsWithoutArguments. This is 1, not a zero-based
# selector; passing 0 makes GetMacroMethods return no modules on current
# SolidWorks versions and causes a false 120-second readiness timeout.
SW_METHODS_WITHOUT_ARGUMENTS = 1
DISP_E_PARAMNOTOPTIONAL = -2147352561


class SolidWorksRunnerError(RuntimeError):
    """A SolidWorks startup or macro error that can be shown to the operator."""


_CUT_FILE_PART_SUFFIX = re.compile(
    r"^(?:(?P<legacy>.+)_\d+-[^_]+_\d+|(?P<current>.+)_\d+)(?:\(B\))?$",
    re.IGNORECASE,
)


def solidworks_part_stem(cut_file_stem: str) -> str:
    """Return the part number from legacy or current orchestrator filenames."""
    match = _CUT_FILE_PART_SUFFIX.fullmatch(cut_file_stem)
    return (match.group("legacy") or match.group("current")) if match else cut_file_stem


def normalize_plate_model_filenames(
    output_folder: Path, candidates: list[Path] | None = None
) -> list[Path]:
    """Remove suffixes from this batch's generated .SLDPRT files."""
    output_folder = Path(output_folder)
    if not output_folder.is_dir():
        raise SolidWorksRunnerError(
            f"SolidWorks output folder does not exist: {output_folder}"
        )

    plan: list[tuple[Path, Path]] = []
    destinations: dict[str, Path] = {}
    sources = (
        candidates
        if candidates is not None
        else [
            path
            for path in output_folder.rglob("*")
            if path.is_file() and path.suffix.casefold() == ".sldprt"
        ]
    )
    for source in sorted(sources, key=lambda path: str(path).casefold()):
        target = source.with_name(f"{solidworks_part_stem(source.stem)}{source.suffix}")
        if target == source:
            continue
        key = os.path.normcase(str(target.resolve()))
        previous = destinations.get(key)
        if previous is not None:
            raise SolidWorksRunnerError(
                "Two generated SolidWorks parts would have the same original "
                f"name: {previous.name} and {source.name} -> {target.name}"
            )
        if target.exists():
            raise SolidWorksRunnerError(
                f"Cannot rename {source.name} to {target.name} because that file "
                "already exists."
            )
        destinations[key] = source
        plan.append((source, target))

    renamed: list[Path] = []
    try:
        for source, target in plan:
            source.rename(target)
            renamed.append(target)
    except OSError as exc:
        for target in reversed(renamed):
            source = next(old for old, new in plan if new == target)
            try:
                target.rename(source)
            except OSError:
                pass
        raise SolidWorksRunnerError(
            f"Could not normalize SolidWorks part filenames: {exc}"
        ) from exc
    return renamed


def snapshot_plate_models(output_folder: Path | None) -> dict[Path, tuple[int, int]]:
    """Capture size and modification time for existing plate models."""
    if output_folder is None or not output_folder.is_dir():
        return {}
    result: dict[Path, tuple[int, int]] = {}
    for path in output_folder.rglob("*"):
        if path.is_file() and path.suffix.casefold() == ".sldprt":
            stat = path.stat()
            result[path.resolve()] = (stat.st_size, stat.st_mtime_ns)
    return result


def changed_plate_models(
    before: dict[Path, tuple[int, int]], output_folder: Path
) -> list[Path]:
    """Return models created or changed after the macro ran."""
    after = snapshot_plate_models(output_folder)
    return sorted(
        (path for path, signature in after.items() if before.get(path) != signature),
        key=lambda path: str(path).casefold(),
    )


def expected_part_stems(source_folder: Path | None) -> set[str]:
    """Return normalized part stems represented by selected source DWGs."""
    if source_folder is None or not source_folder.is_dir():
        return set()
    return {
        solidworks_part_stem(path.stem).casefold()
        for path in source_folder.iterdir()
        if path.is_file() and path.suffix.casefold() == ".dwg"
    }


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
    timeout: float = 120.0,
    get_active: Callable[[], object | None] | None = None,
    dispatch: Callable[[], object] | None = None,
    launch: Callable[..., object] | None = None,
    monotonic: Callable[[], float] | None = None,
    sleep: Callable[[float], None] | None = None,
) -> SolidWorksConnection:
    """Reuse the active instance, otherwise start exactly one SolidWorks app.

    Starting the executable is followed by a COM-registration wait. Macro-level
    readiness is checked separately because SolidWorks can enter the Running
    Object Table while its add-ins and VBA host are still starting.
    """

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


def _available_macro_methods(app, macro: Path) -> list[str]:
    methods = app.GetMacroMethods(str(macro), SW_METHODS_WITHOUT_ARGUMENTS)
    if methods is None:
        return []
    if isinstance(methods, str):
        return [methods]
    return [str(method) for method in methods]


def wait_until_macro_ready(
    app,
    macro: Path,
    *,
    timeout: float = 120.0,
    monotonic: Callable[[], float] | None = None,
    sleep: Callable[[float], None] | None = None,
    methods: Callable[[object, Path], list[str]] | None = None,
) -> list[str]:
    """Wait until the VBA host can inspect the compiled macro.

    A non-empty GetMacroMethods result proves more than an early COM attach: the
    SolidWorks automation object, VBA host, and SWP file are all available.
    """

    monotonic = monotonic or time.monotonic
    sleep = sleep or time.sleep
    methods = methods or _available_macro_methods
    deadline = monotonic() + max(timeout, 1.0)
    last_error = "SolidWorks returned no macro methods."
    while monotonic() < deadline:
        try:
            available = methods(app, macro)
            if available:
                return available
        except Exception as exc:
            last_error = str(exc)
        sleep(0.5)
    raise SolidWorksRunnerError(
        "SolidWorks registered with Windows but its VBA host did not become "
        f"ready within {timeout:g} seconds. Last response: {last_error}"
    )


def resolve_entry_point(
    methods: list[str],
    *,
    preferred_module: str = "",
    preferred_procedure: str = "main",
) -> tuple[str, str]:
    """Resolve the module reported by GetMacroMethods and its procedure.

    SolidWorks documents GetMacroMethods as returning module names. Some COM
    wrappers or older releases have exposed ``Module.Procedure`` strings, so
    both shapes are accepted.
    """

    parsed: list[tuple[str, str]] = []
    modules: list[str] = []
    for method in methods:
        module, separator, procedure = method.rpartition(".")
        if separator and module and procedure:
            parsed.append((module, procedure))
        elif method.strip():
            modules.append(method.strip())
    if not parsed and not modules:
        raise SolidWorksRunnerError(
            "SolidWorks could inspect the macro, but no parameter-free entry "
            f"points were found. Reported methods: {methods!r}"
        )

    if preferred_module:
        for module, procedure in parsed:
            if (
                module.casefold() == preferred_module.casefold()
                and procedure.casefold() == preferred_procedure.casefold()
            ):
                return module, procedure
        for module in modules:
            if module.casefold() == preferred_module.casefold():
                return module, preferred_procedure
    for module, procedure in parsed:
        if procedure.casefold() == preferred_procedure.casefold():
            return module, procedure
    if parsed:
        return parsed[0]
    return modules[0], preferred_procedure


def _missing_out_parameter(error: Exception) -> bool:
    """Return whether COM rejected the omitted RunMacro2 output argument."""

    hresult = getattr(error, "hresult", None)
    if hresult is None and error.args:
        hresult = error.args[0]
    return isinstance(error, TypeError) or hresult == DISP_E_PARAMNOTOPTIONAL


def _invoke_macro(
    app,
    macro: Path,
    module: str,
    procedure: str,
    *,
    variant_factory=None,
    byref_i4: int | None = None,
) -> tuple[bool, int]:
    """Call RunMacro2 through either generated or dynamic pywin32 wrappers."""

    # makepy-generated wrappers omit the final [out] parameter and return it in
    # a tuple. Dynamic dispatch wrappers require an explicit by-reference VARIANT.
    try:
        result = app.RunMacro2(
            str(macro), module, procedure, SW_RUN_MACRO_UNLOAD_AFTER_RUN
        )
        if isinstance(result, tuple):
            succeeded = bool(result[0])
            error_code = int(result[1]) if len(result) > 1 else 0
            return succeeded, error_code
        return bool(result), 0
    except Exception as exc:
        if not _missing_out_parameter(exc):
            raise
        if variant_factory is None or byref_i4 is None:
            import pythoncom
            from win32com.client import VARIANT

            variant_factory = VARIANT
            byref_i4 = pythoncom.VT_BYREF | pythoncom.VT_I4
        error = variant_factory(byref_i4, 0)
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
    module: str = "",
    procedure: str = "main",
    ready_timeout: float = 120.0,
    show_after_run: bool = True,
    ready: Callable[..., list[str]] | None = None,
    invoke: Callable[[object, Path, str, str], tuple[bool, int]] | None = None,
) -> tuple[str, str]:
    """Wait for SolidWorks, resolve the SWP entry point, and run it safely."""

    macro = Path(macro).resolve()
    if not macro.is_file():
        raise SolidWorksRunnerError(f"SolidWorks macro was not found: {macro}")
    if macro.suffix.casefold() != ".swp":
        raise SolidWorksRunnerError(f"Expected a compiled .swp macro: {macro}")

    ready = ready or wait_until_macro_ready
    invoke = invoke or _invoke_macro
    command_flag_set = False
    available = ready(app, macro, timeout=ready_timeout)
    resolved_module, resolved_procedure = resolve_entry_point(
        available,
        preferred_module=module,
        preferred_procedure=procedure,
    )
    try:
        # Officially intended for a sequence of calls from an out-of-process
        # client. It reduces redraw/update overhead without making new part
        # documents invisible to the macro itself.
        try:
            app.CommandInProgress = True
            command_flag_set = True
        except Exception as exc:
            print(f"WARNING: could not enable CommandInProgress: {exc}")

        succeeded, error_code = invoke(
            app, macro, resolved_module, resolved_procedure
        )
        if not succeeded:
            raise SolidWorksRunnerError(
                f"SolidWorks RunMacro2 failed with error code {error_code}."
            )
        return resolved_module, resolved_procedure
    finally:
        if command_flag_set:
            try:
                app.CommandInProgress = False
            except Exception as exc:
                print(f"WARNING: could not clear CommandInProgress: {exc}")
        if show_after_run:
            try:
                app.Visible = True
                app.UserControl = True
            except Exception as exc:
                print(f"WARNING: could not restore the SolidWorks window: {exc}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run the CAD batch converter in a reusable SolidWorks session."
    )
    parser.add_argument("--macro", required=True, type=Path)
    parser.add_argument("--module", default="")
    parser.add_argument("--procedure", default="main")
    parser.add_argument("--solidworks-executable", type=Path)
    parser.add_argument("--connect-timeout", type=float, default=120.0)
    parser.add_argument("--ready-timeout", type=float, default=120.0)
    parser.add_argument(
        "--normalize-output",
        type=Path,
        help=(
            "After a successful CAD batch, remove the orchestrator material and "
            "quantity suffix from generated SolidWorks part filenames."
        ),
    )
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
        print(f"Waiting for the SolidWorks VBA host and {args.macro.name}.", flush=True)
        output_folder = args.normalize_output
        source_folder = None
        if args.macro.name.casefold() == "main.runbatch.swp":
            if output_folder is None:
                configured_output = os.environ.get("MACROS_OUTPUT_FOLDER", "").strip()
                output_folder = Path(configured_output) if configured_output else None
            configured_source = os.environ.get("MACROS_SOURCE_FOLDER", "").strip()
            source_folder = Path(configured_source) if configured_source else None
        before_models = snapshot_plate_models(output_folder)
        module, procedure = run_macro(
            connection.app,
            args.macro,
            module=args.module,
            procedure=args.procedure,
            ready_timeout=args.ready_timeout,
            show_after_run=not args.leave_hidden,
        )
        if output_folder is not None:
            changed = changed_plate_models(before_models, output_folder)
            expected = expected_part_stems(source_folder)
            matched = [
                part
                for part in changed
                if not expected
                or solidworks_part_stem(part.stem).casefold() in expected
            ]
            if not matched:
                source_text = str(source_folder) if source_folder else "configured source"
                raise SolidWorksRunnerError(
                    "The macro returned success but created or updated no SolidWorks "
                    f"parts matching DWGs in {source_text}. The batch was not accepted."
                )
            renamed = normalize_plate_model_filenames(output_folder, matched)
            print(
                f"Verified {len(matched)} current-batch part(s); normalized "
                f"{len(renamed)} filename(s).",
                flush=True,
            )
            for part in renamed:
                print(f"  {part.name}", flush=True)
        print(f"SolidWorks macro {module}.{procedure} completed.", flush=True)
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
