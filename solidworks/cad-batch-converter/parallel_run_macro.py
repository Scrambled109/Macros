"""Experimental parallel SolidWorks plate-batch controller."""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

from run_macro import (
    SolidWorksRunnerError,
    cleanup_speed_workspace,
    expected_part_stems,
    publish_plate_outputs,
    solidworks_part_stem,
)


@dataclass
class Worker:
    number: int
    source: Path
    filtered: Path
    output: Path
    ready: Path
    log: Path
    process: subprocess.Popen | None = None
    handle: object | None = None


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Run isolated SolidWorks plate workers in parallel."
    )
    parser.add_argument("--macro", required=True, type=Path)
    parser.add_argument("--runner", required=True, type=Path)
    parser.add_argument("--solidworks-executable", required=True, type=Path)
    parser.add_argument("--source", required=True, type=Path)
    parser.add_argument("--workspace", required=True, type=Path)
    parser.add_argument("--publish-output", required=True, type=Path)
    parser.add_argument("--extrude-depth-meters", required=True, type=float)
    parser.add_argument("--workers", type=int, default=2)
    parser.add_argument("--startup-timeout", type=float, default=300.0)
    return parser


def prepare_workers(source: Path, workspace: Path, requested: int) -> list[Worker]:
    drawings = sorted(
        (
            path
            for path in source.iterdir()
            if path.is_file() and path.suffix.casefold() == ".dwg"
        ),
        key=lambda path: path.name.casefold(),
    )
    if not drawings:
        raise SolidWorksRunnerError(f"No DWGs were found in {source}")
    count = max(1, min(int(requested), len(drawings)))
    root = workspace / "Parallel Workers"
    root.mkdir(parents=True, exist_ok=False)
    workers: list[Worker] = []
    for index in range(count):
        worker_root = root / f"worker-{index + 1:02d}"
        item = Worker(
            number=index + 1,
            source=worker_root / "Source DWGs",
            filtered=worker_root / "Filtered DWGs",
            output=worker_root / "SolidWorks Parts",
            ready=worker_root / "session-ready.json",
            log=worker_root / "worker.log",
        )
        for folder in (item.source, item.filtered, item.output):
            folder.mkdir(parents=True, exist_ok=False)
        workers.append(item)
    for index, drawing in enumerate(drawings):
        target_worker = workers[index % len(workers)]
        shutil.copy2(drawing, target_worker.source / drawing.name)
    return workers


def worker_command(
    worker: Worker,
    *,
    runner: Path,
    macro: Path,
    solidworks_executable: Path,
    start_signal: Path,
) -> list[str]:
    return [
        sys.executable,
        "-u",
        str(runner),
        "--macro",
        str(macro),
        "--solidworks-executable",
        str(solidworks_executable),
        "--normalize-output",
        str(worker.output),
        "--new-instance",
        "--quit-after-run",
        "--leave-hidden",
        "--worker-ready-file",
        str(worker.ready),
        "--worker-start-signal",
        str(start_signal),
    ]


def abort_workers(workers: list[Worker], start_signal: Path) -> None:
    start_signal.write_text("abort\n", encoding="utf-8")
    deadline = time.monotonic() + 30.0
    for worker in workers:
        if worker.process is None:
            continue
        remaining = max(0.1, deadline - time.monotonic())
        try:
            worker.process.wait(timeout=remaining)
        except subprocess.TimeoutExpired:
            worker.process.terminate()
    for worker in workers:
        if worker.handle is not None:
            worker.handle.close()
            worker.handle = None


def wait_for_worker_session(worker: Worker, timeout: float) -> int:
    """Wait for one controller to attach to its newly launched SolidWorks PID."""
    deadline = time.monotonic() + max(timeout, 1.0)
    while time.monotonic() < deadline:
        if worker.ready.is_file():
            data = json.loads(worker.ready.read_text(encoding="utf-8"))
            return int(data["solidworks_pid"])
        if worker.process is not None and worker.process.poll() is not None:
            raise SolidWorksRunnerError(
                f"Worker {worker.number} exited before its SolidWorks session "
                f"became ready. Review {worker.log}"
            )
        time.sleep(0.2)
    raise SolidWorksRunnerError(
        f"Worker {worker.number} timed out while starting its dedicated "
        f"SolidWorks session. Review {worker.log}"
    )


def wait_for_unique_sessions(
    workers: list[Worker], start_signal: Path, timeout: float
) -> dict[int, int]:
    deadline = time.monotonic() + max(timeout, 1.0)
    sessions: dict[int, int] = {}
    while len(sessions) < len(workers):
        for worker in workers:
            if worker.number in sessions:
                continue
            if worker.ready.is_file():
                data = json.loads(worker.ready.read_text(encoding="utf-8"))
                sessions[worker.number] = int(data["solidworks_pid"])
                print(
                    f"Worker {worker.number} ready on SolidWorks PID "
                    f"{sessions[worker.number]}.",
                    flush=True,
                )
                continue
            if worker.process is not None and worker.process.poll() is not None:
                abort_workers(workers, start_signal)
                raise SolidWorksRunnerError(
                    f"Worker {worker.number} exited before its SolidWorks session "
                    f"became ready. Review {worker.log}"
                )
        if time.monotonic() >= deadline:
            abort_workers(workers, start_signal)
            raise SolidWorksRunnerError(
                "Timed out while starting dedicated SolidWorks worker sessions."
            )
        time.sleep(0.2)
    if len(set(sessions.values())) != len(sessions):
        abort_workers(workers, start_signal)
        raise SolidWorksRunnerError(
            "SolidWorks did not create one unique process per worker. No macro "
            "was run and nothing was published."
        )
    return sessions


def combine_batch_logs(workers: list[Worker], destination: Path) -> Path:
    lines = ["PARALLEL SOLIDWORKS BATCH LOG", ""]
    for worker in workers:
        batch_log = worker.output / "BatchLog.txt"
        lines.append(f"===== WORKER {worker.number:02d} =====")
        if batch_log.is_file():
            lines.append(batch_log.read_text(encoding="utf-8", errors="replace"))
        else:
            lines.append(f"Missing worker batch log: {batch_log}")
        lines.append("")
    destination.write_text("\n".join(lines), encoding="utf-8")
    return destination


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    started = time.perf_counter()
    workers: list[Worker] = []
    start_signal = args.workspace / "parallel-start.signal"
    try:
        expected = expected_part_stems(args.source)
        if not expected:
            raise SolidWorksRunnerError(
                f"No expected plate names could be read from {args.source}"
            )
        # Fail before doing expensive work if job staging is contaminated.
        publish_plate_outputs([], args.publish_output, expected)
        workers = prepare_workers(args.source, args.workspace, args.workers)
        print(
            f"Starting {len(workers)} isolated SolidWorks worker(s) for "
            f"{len(expected)} DWG(s).",
            flush=True,
        )
        sessions: dict[int, int] = {}
        for worker in workers:
            environment = os.environ.copy()
            environment.update(
                {
                    "MACROS_SOURCE_FOLDER": str(worker.source),
                    "MACROS_FILTERED_FOLDER": str(worker.filtered),
                    "MACROS_OUTPUT_FOLDER": str(worker.output),
                    "MACROS_EXTRUDE_DEPTH_METERS": format(
                        args.extrude_depth_meters, ".12g"
                    ),
                }
            )
            worker.handle = worker.log.open("w", encoding="utf-8")
            worker.process = subprocess.Popen(
                worker_command(
                    worker,
                    runner=args.runner,
                    macro=args.macro,
                    solidworks_executable=args.solidworks_executable,
                    start_signal=start_signal,
                ),
                cwd=str(args.runner.parent),
                env=environment,
                stdout=worker.handle,
                stderr=subprocess.STDOUT,
            )
            solidworks_pid = wait_for_worker_session(
                worker, args.startup_timeout
            )
            if solidworks_pid in sessions.values():
                abort_workers(workers, start_signal)
                raise SolidWorksRunnerError(
                    f"Worker {worker.number} attached to already-claimed "
                    f"SolidWorks PID {solidworks_pid}. No macro was run and "
                    "nothing was published."
                )
            sessions[worker.number] = solidworks_pid
            print(
                f"Worker {worker.number} ready on unique SolidWorks PID "
                f"{solidworks_pid}; starting the next instance.",
                flush=True,
            )
        print(
            "Verified unique SolidWorks PIDs: "
            + ", ".join(str(pid) for pid in sessions.values()),
            flush=True,
        )
        start_signal.write_text("start\n", encoding="utf-8")

        failures: list[str] = []
        for worker in workers:
            assert worker.process is not None
            exit_code = worker.process.wait()
            if worker.handle is not None:
                worker.handle.close()
                worker.handle = None
            print(
                f"Worker {worker.number} finished with exit code {exit_code}.",
                flush=True,
            )
            if exit_code != 0:
                failures.append(f"worker {worker.number}: {worker.log}")
        if failures:
            raise SolidWorksRunnerError(
                "Parallel batch failed; local workspace was preserved. "
                + "; ".join(failures)
            )

        parts = sorted(
            (
                path
                for worker in workers
                for path in worker.output.glob("*.SLDPRT")
                if path.is_file()
            ),
            key=lambda path: path.name.casefold(),
        )
        actual = {solidworks_part_stem(path.stem).casefold() for path in parts}
        missing = sorted(expected - actual)
        extra = sorted(actual - expected)
        if missing or extra:
            raise SolidWorksRunnerError(
                f"Parallel output set did not match the selected DWGs. "
                f"Missing={missing[:10]}, unexpected={extra[:10]}. "
                "Nothing was published."
            )

        combined_log = combine_batch_logs(
            workers, args.workspace / "BatchLog.txt"
        )
        published = publish_plate_outputs(
            parts,
            args.publish_output,
            expected,
            batch_log=combined_log,
        )
        elapsed = time.perf_counter() - started
        print(
            f"Published {len(published)} verified parts from {len(workers)} "
            f"SolidWorks instances in {elapsed:.1f} s.",
            flush=True,
        )
        cleanup_speed_workspace(args.workspace)
        print("Removed verified local parallel workspace.", flush=True)
        return 0
    except SolidWorksRunnerError as exc:
        if any(
            worker.process is not None and worker.process.poll() is None
            for worker in workers
        ):
            abort_workers(workers, start_signal)
        print(f"ERROR: {exc}", file=sys.stderr, flush=True)
        return 1
    except Exception as exc:
        if any(
            worker.process is not None and worker.process.poll() is None
            for worker in workers
        ):
            abort_workers(workers, start_signal)
        print(f"ERROR: unexpected parallel-controller failure: {exc}", file=sys.stderr)
        return 1
    finally:
        for worker in workers:
            if worker.handle is not None:
                worker.handle.close()


if __name__ == "__main__":
    raise SystemExit(main())
