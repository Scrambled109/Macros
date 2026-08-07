"""Run the compiled CAD batch macro through one reusable SolidWorks session.

The runner deliberately leaves document visibility to the VBA macro. Hiding all
part documents from an out-of-process controller changes how newly imported
documents are activated and made the previous runner less stable than running
the same SWP manually.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import time
import traceback
from typing import Callable
import uuid


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


def publish_plate_outputs(
    parts: list[Path],
    destination: Path,
    expected_stems: set[str],
    *,
    batch_log: Path | None = None,
) -> list[Path]:
    """Atomically copy verified current-batch outputs back to job staging."""
    destination = Path(destination)
    destination.mkdir(parents=True, exist_ok=True)
    unrelated = [
        path
        for path in destination.iterdir()
        if path.is_file()
        and path.suffix.casefold() == ".sldprt"
        and solidworks_part_stem(path.stem).casefold() not in expected_stems
    ]
    if unrelated:
        examples = ", ".join(path.name for path in unrelated[:5])
        raise SolidWorksRunnerError(
            "The published staging folder contains unrelated part files "
            f"(for example: {examples}). Move that folder aside before rerunning."
        )

    published: list[Path] = []
    for source in parts:
        target = destination / source.name
        temporary = destination / f".{source.name}.{uuid.uuid4().hex}.tmp"
        try:
            shutil.copy2(source, temporary)
            if temporary.stat().st_size != source.stat().st_size:
                raise OSError("copied file size does not match the local result")
            os.replace(temporary, target)
        except OSError as exc:
            if temporary.exists():
                temporary.unlink()
            raise SolidWorksRunnerError(
                f"Could not publish {source.name} to {destination}: {exc}"
            ) from exc
        published.append(target)

    if batch_log is not None and batch_log.is_file():
        log_target = destination / batch_log.name
        log_temporary = destination / f".{batch_log.name}.{uuid.uuid4().hex}.tmp"
        try:
            shutil.copy2(batch_log, log_temporary)
            os.replace(log_temporary, log_target)
        except OSError as exc:
            if log_temporary.exists():
                log_temporary.unlink()
            raise SolidWorksRunnerError(
                f"Parts were published, but BatchLog.txt could not be copied: {exc}"
            ) from exc
    return published


def cleanup_speed_workspace(workspace: Path) -> None:
    """Remove only a marked assistant-owned local speed workspace."""
    workspace = Path(workspace).resolve()
    marker = workspace / ".engineering-job-assistant-plate-batch"
    if not marker.is_file():
        raise SolidWorksRunnerError(
            f"Refusing to remove unmarked local workspace: {workspace}"
        )
    shutil.rmtree(workspace)


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


def _running_solidworks_sessions() -> dict[int, object]:
    """Return every SolidWorks COM session currently registered with Windows."""
    import pythoncom
    import win32com.client

    sessions: dict[int, object] = {}
    table = pythoncom.GetRunningObjectTable()
    enumerator = table.EnumRunning()
    bind_context = pythoncom.CreateBindCtx(0)
    while True:
        monikers = enumerator.Next(1)
        if not monikers:
            break
        moniker = monikers[0]
        try:
            # Do not filter by moniker text. Different SolidWorks releases have
            # used SldWorks..., SolidWorks_PID..., and opaque ROT names. Probe
            # the object itself for the SolidWorks GetProcessID API instead.
            app = win32com.client.Dispatch(table.GetObject(moniker))
            process_id_value = app.GetProcessID
            process_id = int(
                process_id_value()
                if callable(process_id_value)
                else process_id_value
            )
            sessions[process_id] = app
        except Exception:
            continue
    return sessions


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


def _dispatch_new_solidworks():
    """Create a separate COM server instead of attaching to an existing one."""
    import win32com.client

    failures: list[str] = []
    for prog_id in SOLIDWORKS_PROGIDS:
        try:
            # Dispatch can reuse the active local server. DispatchEx calls
            # CoCreateInstanceEx, which gives each parallel worker its own
            # SolidWorks application object and process.
            return win32com.client.DispatchEx(prog_id)
        except Exception as exc:
            failures.append(f"{prog_id}: {exc}")
    raise SolidWorksRunnerError(
        "Could not create a dedicated SolidWorks COM instance. "
        + "; ".join(failures)
    )


def connect_solidworks(
    executable: Path | None = None,
    *,
    timeout: float = 120.0,
    get_active: Callable[[], object | None] | None = None,
    dispatch: Callable[[], object] | None = None,
    dispatch_new: Callable[[], object] | None = None,
    get_sessions: Callable[[], dict[int, object]] | None = None,
    launch: Callable[..., object] | None = None,
    force_new: bool = False,
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
    dispatch_new = dispatch_new or _dispatch_new_solidworks
    get_sessions = get_sessions or _running_solidworks_sessions
    launch = launch or subprocess.Popen
    monotonic = monotonic or time.monotonic
    sleep = sleep or time.sleep

    if force_new:
        if executable is None:
            raise SolidWorksRunnerError(
                "Dedicated SolidWorks workers require --solidworks-executable "
                "so the configured installation can be validated."
            )
        executable = Path(executable).resolve()
        if not executable.is_file():
            raise SolidWorksRunnerError(
                f"Configured SolidWorks executable was not found: {executable}"
            )

        sessions_before = set(get_sessions())
        app = dispatch_new()
        process_id_value = app.GetProcessID
        process_id = int(
            process_id_value() if callable(process_id_value) else process_id_value
        )
        if process_id in sessions_before:
            try:
                app.ExitApp()
            except Exception:
                pass
            raise SolidWorksRunnerError(
                "SolidWorks COM returned an already-running process while a "
                "dedicated worker was requested. Refusing to share PID "
                f"{process_id} between parallel workers."
            )
        return SolidWorksConnection(
            app,
            reused=False,
            launch_method=f"dedicated COM server process {process_id}",
        )

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
        "--new-instance",
        action="store_true",
        help="Create a dedicated SolidWorks COM instance instead of reusing one.",
    )
    parser.add_argument("--worker-ready-file", type=Path)
    parser.add_argument("--worker-start-signal", type=Path)
    parser.add_argument("--worker-start-timeout", type=float, default=240.0)
    parser.add_argument(
        "--quit-after-run",
        action="store_true",
        help="Exit the dedicated SolidWorks instance when this worker finishes.",
    )
    parser.add_argument(
        "--normalize-output",
        type=Path,
        help=(
            "After a successful CAD batch, remove the orchestrator material and "
            "quantity suffix from generated SolidWorks part filenames."
        ),
    )
    parser.add_argument(
        "--publish-output",
        type=Path,
        help="Copy verified local plate results to this job staging folder.",
    )
    parser.add_argument(
        "--cleanup-workspace",
        type=Path,
        help="Remove this marked local speed workspace after verified publication.",
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
    connection = None
    try:
        connection = connect_solidworks(
            args.solidworks_executable,
            timeout=args.connect_timeout,
            force_new=args.new_instance,
        )
        action = "Reusing" if connection.reused else "Started"
        print(f"{action} SolidWorks via {connection.launch_method}.", flush=True)
        solidworks_pid_value = connection.app.GetProcessID
        solidworks_pid = int(
            solidworks_pid_value()
            if callable(solidworks_pid_value)
            else solidworks_pid_value
        )
        print(f"SolidWorks process ID: {solidworks_pid}", flush=True)
        if args.worker_ready_file is not None:
            args.worker_ready_file.parent.mkdir(parents=True, exist_ok=True)
            temporary_ready = args.worker_ready_file.with_suffix(
                args.worker_ready_file.suffix + ".tmp"
            )
            temporary_ready.write_text(
                json.dumps(
                    {
                        "controller_pid": os.getpid(),
                        "solidworks_pid": solidworks_pid,
                    }
                )
                + "\n",
                encoding="utf-8",
            )
            os.replace(temporary_ready, args.worker_ready_file)
        if args.worker_start_signal is not None:
            deadline = time.monotonic() + max(args.worker_start_timeout, 1.0)
            while not args.worker_start_signal.is_file():
                if time.monotonic() >= deadline:
                    raise SolidWorksRunnerError(
                        "Timed out waiting for the parallel-worker start signal."
                    )
                time.sleep(0.1)
            signal = args.worker_start_signal.read_text(
                encoding="utf-8", errors="replace"
            ).strip().casefold()
            if signal != "start":
                raise SolidWorksRunnerError(
                    "Parallel launch was aborted before any macro executed."
                )
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
        expected_before_run = expected_part_stems(source_folder)
        if expected_before_run:
            unrelated = [
                part
                for part in before_models
                if solidworks_part_stem(part.stem).casefold() not in expected_before_run
            ]
            if unrelated:
                examples = ", ".join(part.name for part in unrelated[:5])
                raise SolidWorksRunnerError(
                    "The selected SolidWorks output folder already contains "
                    f"{len(unrelated)} part(s) unrelated to the selected DWGs "
                    f"(for example: {examples}). Move that staging folder aside "
                    "before rerunning so unrelated models cannot be accepted or moved."
                )
        macro_started = time.perf_counter()
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
            expected = expected_before_run
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
            final_parts = [
                part.with_name(
                    f"{solidworks_part_stem(part.stem)}{part.suffix}"
                )
                for part in matched
            ]
            missing_final = [part for part in final_parts if not part.is_file()]
            if missing_final:
                raise SolidWorksRunnerError(
                    "Filename normalization lost expected current-batch outputs: "
                    + ", ".join(part.name for part in missing_final[:5])
                )
            macro_seconds = time.perf_counter() - macro_started
            print(
                f"Verified {len(matched)} current-batch part(s); normalized "
                f"{len(renamed)} filename(s). Macro time: {macro_seconds:.1f} s.",
                flush=True,
            )
            if args.publish_output is not None:
                publish_started = time.perf_counter()
                published = publish_plate_outputs(
                    final_parts,
                    args.publish_output,
                    expected,
                    batch_log=output_folder / "BatchLog.txt",
                )
                publish_seconds = time.perf_counter() - publish_started
                print(
                    f"Published {len(published)} verified part(s) to "
                    f"{args.publish_output} in {publish_seconds:.1f} s.",
                    flush=True,
                )
                for part in published:
                    print(f"  {part.name}", flush=True)
                if args.cleanup_workspace is not None:
                    cleanup_speed_workspace(args.cleanup_workspace)
                    print(
                        f"Removed local speed workspace: {args.cleanup_workspace}",
                        flush=True,
                    )
            else:
                for part in final_parts:
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
        if args.quit_after_run and connection is not None:
            try:
                connection.app.ExitApp()
                print("Closed dedicated SolidWorks worker instance.", flush=True)
            except Exception as exc:
                print(
                    f"WARNING: could not close dedicated SolidWorks worker: {exc}",
                    file=sys.stderr,
                    flush=True,
                )
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
