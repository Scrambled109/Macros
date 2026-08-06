#!/usr/bin/env python3
"""Python implementation of the AutoCAD DXF master orchestrator.

This is the implementation launched by the Engineering Job Assistant. It uses
only the Python standard library and keeps input, output, archive, and logging
behavior deterministic across unattended runs.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path


SUPPORTED_AUTOCAD_YEARS = (2026, 2025)


def autocad_console_candidates(program_files: Path | str | None = None) -> list[Path]:
    """Return installed-version candidates from newest to oldest."""
    root = Path(
        program_files
        if program_files is not None
        else os.environ.get("ProgramFiles", r"C:\Program Files")
    )
    autodesk = root / "Autodesk"
    candidates = [
        autodesk / f"AutoCAD {year}" / "accoreconsole.exe"
        for year in SUPPORTED_AUTOCAD_YEARS
    ]
    try:
        installed = sorted(
            (
                path
                for path in autodesk.glob("AutoCAD *")
                if path.is_dir()
            ),
            key=lambda path: tuple(
                int(part) for part in re.findall(r"\d+", path.name)
            ) or (0,),
            reverse=True,
        )
    except OSError:
        installed = []

    for folder in installed:
        candidate = folder / "accoreconsole.exe"
        if candidate not in candidates:
            candidates.append(candidate)
    return candidates


def detect_autocad_console(
    explicit: Path | str | None = None,
    program_files: Path | str | None = None,
) -> Path:
    """Resolve an override or select the newest installed Core Console."""
    configured = explicit or os.environ.get("ACAD_CONSOLE_PATH")
    if configured:
        return Path(configured)

    candidates = autocad_console_candidates(program_files)
    for candidate in candidates:
        if candidate.is_file():
            return candidate

    checked = "\n  ".join(str(path) for path in candidates)
    raise FileNotFoundError(
        "AutoCAD Core Console was not found. Install AutoCAD 2025/2026, "
        "or set an explicit Core Console path. Checked:\n  "
        f"{checked}"
    )
CONSOLE_TIMEOUT_SECONDS = 180
MINIMUM_DWG_BYTES = 1024

HASH_TO_DASH_LISP = """(defun c:HashToDash ( / ss i ent oldStr newStr)
  (if (setq ss (ssget "X" '((0 . "TEXT,MTEXT"))))
    (progn
      (setq i 0)
      (while (< i (sslength ss))
        (setq ent (entget (ssname ss i)))
        (setq oldStr (cdr (assoc 1 ent)))
        (setq newStr oldStr)
        (while (vl-string-search "#" newStr)
          (setq newStr (vl-string-subst "-" "#" newStr))
        )
        (if (/= oldStr newStr)
          (entmod (subst (cons 1 newStr) (assoc 1 ent) ent))
        )
        (setq i (1+ i))
      )
    )
  )
  (princ)
)
"""


def thickness_to_mils(raw: str | None) -> str:
    """Normalize decimal or fractional inches to thousandths of an inch."""
    value = (raw or "").strip()
    if re.fullmatch(r"\d*\.\d+", value):
        return str(round(float(value) * 1000))

    match = re.fullmatch(r"(?:(\d+)[-\s])?(\d+)\s*/\s*(\d+)", value)
    if match:
        whole = float(match.group(1) or 0)
        numerator = float(match.group(2))
        denominator = float(match.group(3))
        if denominator:
            return str(round((whole + numerator / denominator) * 1000))
    return value


def safe_name(value: str) -> str:
    """Replace characters illegal in Windows filenames with hyphens."""
    return re.sub(r'[\\/:*?"<>|]', "-", value)


def has_bevel_annotation(value: str) -> bool:
    """Recognize conservative bevel/snipe notes from drawing text.

    Missing the filename marker is more costly than marking one extra drawing,
    so this intentionally accepts punctuation, spacing, common shop synonyms,
    and arrow glyphs used in exported callouts.
    """
    return bool(
        re.search(
            r"\b(?:BEVEL(?:ED)?|BVL|CHAMFER|SNIP(?:E|ED|ING)?|BACK\s*GOUGE|GOUGE)\b"
            r"|\b(?:K|V|RV)\b"
            r"|\b(?:K|V|RV)\s*[-–—:/]?\s*\d+(?:\.\d+)?\b"
            r"|(?:-{1,2}>|<{1,2}-|[←↑→↓↖↗↘↙⇐⇑⇒⇓➔➜➤►◄△▽])",
            value,
            re.IGNORECASE,
        )
    )


def is_beveled_dxf(path: Path) -> bool:
    """Look for bevel annotations in ASCII DXF TEXT/MTEXT entities."""
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError:
        return False

    entity_type = ""
    text_chunks: list[str] = []
    for index in range(0, len(lines) - 1, 2):
        code = lines[index].strip()
        value = lines[index + 1]
        if code == "0":
            if entity_type in {"TEXT", "MTEXT"} and has_bevel_annotation(
                "".join(text_chunks)
            ):
                return True
            entity_type = value.strip().upper()
            # LEADER/MLEADER entities explicitly attach an arrow to a note.
            # Mark them as bevel drawings even when their annotation text is
            # stored in a referenced block rather than inline in the DXF.
            if entity_type in {"LEADER", "MLEADER", "MULTILEADER"}:
                return True
            text_chunks = []
        elif code in {"1", "3"} and entity_type in {"TEXT", "MTEXT"}:
            text_chunks.append(value)
    return entity_type in {"TEXT", "MTEXT"} and has_bevel_annotation(
        "".join(text_chunks)
    )


def wait_file_stable(path: Path, timeout_seconds: float = 30) -> bool:
    """Wait for a valid DWG to retain its size and become readable."""
    deadline = time.monotonic() + timeout_seconds
    previous_size = -1
    while time.monotonic() < deadline:
        try:
            size = path.stat().st_size
            if size > MINIMUM_DWG_BYTES and size == previous_size:
                with path.open("rb"):
                    return True
            previous_size = size
        except OSError:
            pass
        time.sleep(0.4)
    return False


def load_parts(path: Path) -> dict[str, dict[str, str]]:
    if not path.exists():
        print(f"WARNING: CSV not found at {path}. Parts will be left in Unsorted.")
        return {}

    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        headers = reader.fieldnames or []
        required = {"PartNumber", "Thickness", "Material"}
        missing = sorted(required.difference(headers))
        if missing:
            raise ValueError(f"parts.csv is missing required column(s): {', '.join(missing)}")

        parts: dict[str, dict[str, str]] = {}
        for row in reader:
            part = (row.get("PartNumber") or "").strip()
            if not part:
                continue
            parts[part] = {
                "quantity": (row.get("Quantity") or row.get("Quanity") or "1").strip(),
                "thickness": (row.get("Thickness") or "").strip(),
                "material": (row.get("Material") or "").strip(),
            }
    print(f"Successfully loaded data for {len(parts)} parts from CSV.")
    return parts


def autocad_path(path: Path) -> str:
    return str(path.resolve()).replace("\\", "/")


def write_script(path: Path, lines: list[str]) -> None:
    # Do not pass pre-built CRLF text through Path.write_text on Windows.  Its
    # normal newline translation turns each CRLF into CRCRLF, which AutoCAD
    # interprets as an empty response between a command and its value (for
    # example FILEDIA, Enter, blank, then a stray `0` command).
    with path.open("w", encoding="ascii", newline="") as handle:
        handle.write("\r\n".join(lines) + "\r\n")


def output_details(
    dxf: Path,
    parts: dict[str, dict[str, str]],
    workspace: Path,
    *,
    beveled: bool = False,
) -> tuple[str, Path]:
    """Return the sorting directory and suffix-aware final DWG path."""
    original_stem = dxf.stem
    part = parts.get(original_stem)
    if part:
        quantity = part["quantity"]
        target_name = safe_name(
            f"{thickness_to_mils(part['thickness'])}-{part['material']}"
        )
    else:
        quantity = "1"
        target_name = "Unsorted"

    working_name = original_stem.replace("#", "-")
    working_name = re.sub(r"\s+_", "_", working_name)
    working_name = re.sub(r"_\d+-[A-Za-z0-9-]+_\d+$", "", working_name)
    working_name = re.sub(r"_\d+_\d+$", "", working_name)
    working_name = safe_name(f"{working_name}_{target_name}_{quantity}")
    if beveled:
        working_name += "(B)"
    target_dir = workspace / target_name
    target_dir.mkdir(exist_ok=True)
    return target_name, target_dir / f"{working_name}.dwg"


def archive_original(source: Path, archive_dir: Path) -> None:
    destination = archive_dir / source.name
    suffix = 2
    while destination.exists():
        destination = archive_dir / f"{source.stem}_{suffix}{source.suffix}"
        suffix += 1
    shutil.move(str(source), str(destination))


def process_file(
    dxf: Path,
    parts: dict[str, dict[str, str]],
    workspace: Path,
    archive_dir: Path,
    log_dir: Path,
    acad_console: Path,
    lsp_path: Path,
    seed_path: Path,
    h2d_path: Path,
    script_path: Path,
    console_timeout: int,
) -> str:
    beveled = is_beveled_dxf(dxf)
    target_name, final_dwg = output_details(
        dxf,
        parts,
        workspace,
        beveled=beveled,
    )
    escaped_dwg = autocad_path(final_dwg)
    preamble = [
        "FILEDIA", "0", "SECURELOAD", "0", "-INSERT",
        f'"{autocad_path(seed_path)}"', "0,0", "1", "1", "0",
        "ERASE", "L", "", f'(load "{autocad_path(lsp_path)}")',
        "ColorToLayer", f'(load "{autocad_path(h2d_path)}")', "HashToDash",
    ]
    result_kind = "bevel" if beveled else "clean"
    if beveled:
        print(
            "    [B] Bevel detected. Marking the DWG with (B) and running "
            "through AutoCAD Core Console...",
            flush=True,
        )
    else:
        print(
            "    [+] Standard part. Running through AutoCAD Core Console...",
            flush=True,
        )
    save_lisp = (
        f'(if (findfile "{escaped_dwg}") '
        f'(command "_.SAVEAS" "2018" "{escaped_dwg}" "Y") '
        f'(command "_.SAVEAS" "2018" "{escaped_dwg}"))'
    )
    write_script(script_path, preamble + [
        save_lisp, "SECURELOAD", "1",
        "FILEDIA", "1", "_QUIT", "Y",
    ])
    log_key = safe_name(f"{dxf.stem}_{script_path.stem}")
    log_path = log_dir / f"{log_key}.log"
    error_path = log_dir / f"{log_key}.err.log"
    with log_path.open("w", encoding="utf-8") as stdout, error_path.open(
        "w", encoding="utf-8"
    ) as stderr:
        process = subprocess.Popen(
            [str(acad_console), "/i", str(dxf), "/s", str(script_path)],
            stdout=stdout,
            stderr=stderr,
        )
        try:
            exit_code = process.wait(timeout=console_timeout)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait()
            print(f"    [X] TIMEOUT after {console_timeout}s. See {log_path}. Original DXF kept.")
            return "failed"

    if exit_code == 0 and wait_file_stable(final_dwg):
        archive_original(dxf, archive_dir)
        print(f"    [SUCCESS] Converted to {final_dwg.name} in {target_name}")
        return result_kind
    print(f"    [X] ERROR: exit={exit_code}; DWG missing or too small. Original DXF kept. Log: {log_path}")
    return "failed"


def split_output_collisions(
    files: list[Path],
    parts: dict[str, dict[str, str]],
    workspace: Path,
) -> tuple[list[Path], list[Path]]:
    """Keep inputs with unique output paths and reject ambiguous duplicates."""

    by_destination: dict[str, list[Path]] = {}
    for dxf in files:
        _target, destination = output_details(
            dxf,
            parts,
            workspace,
            beveled=is_beveled_dxf(dxf),
        )
        key = os.path.normcase(str(destination.resolve()))
        by_destination.setdefault(key, []).append(dxf)

    collisions = {
        dxf
        for grouped_files in by_destination.values()
        if len(grouped_files) > 1
        for dxf in grouped_files
    }
    safe_files = [dxf for dxf in files if dxf not in collisions]
    return safe_files, sorted(collisions, key=lambda path: str(path).casefold())


def process_files(
    files: list[Path],
    *,
    parts: dict[str, dict[str, str]],
    workspace: Path,
    archive_dir: Path,
    log_dir: Path,
    acad_console: Path,
    lsp_path: Path,
    seed_path: Path,
    h2d_path: Path,
    temp_dir: Path,
    workers: int,
    console_timeout: int,
) -> dict[str, int]:
    """Run standard and bevel-marked DXFs through console workers."""

    counts = {"clean": 0, "bevel": 0, "failed": 0}
    safe_files, collisions = split_output_collisions(files, parts, workspace)
    for dxf in collisions:
        print(
            f"    [X] {dxf}: another input resolves to the same production DWG. "
            "Neither duplicate was processed; rename or remove the stale input.",
            flush=True,
        )
        counts["failed"] += 1

    if not safe_files:
        return counts

    worker_count = max(1, workers)
    print(
        f"\n=== RUNNING {len(safe_files)} DXF FILE(S) WITH "
        f"{worker_count} AUTOCAD CORE WORKER(S) ===",
        flush=True,
    )
    with ThreadPoolExecutor(
        max_workers=worker_count,
        thread_name_prefix="autocad-core",
    ) as executor:
        futures = {}
        for index, dxf in enumerate(safe_files, start=1):
            script_path = temp_dir / f"job_{index:04d}.scr"
            future = executor.submit(
                process_file,
                dxf,
                parts,
                workspace,
                archive_dir,
                log_dir,
                acad_console,
                lsp_path,
                seed_path,
                h2d_path,
                script_path,
                console_timeout,
            )
            futures[future] = dxf

        for future in as_completed(futures):
            dxf = futures[future]
            try:
                result = future.result()
            except Exception as exc:
                print(
                    f"    [X] {dxf.name}: unexpected worker error: {exc}. "
                    "Original DXF kept.",
                    file=sys.stderr,
                    flush=True,
                )
                result = "failed"
            counts[result if result in counts else "failed"] += 1
    return counts


def worker_count(value: str) -> int:
    try:
        parsed = int(value)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("workers must be a positive integer") from exc
    if parsed < 1:
        raise argparse.ArgumentTypeError("workers must be at least 1")
    return parsed


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--acad-console-path",
        type=Path,
        default=None,
        help=(
            "Explicit accoreconsole.exe path. By default the newest installed "
            "AutoCAD version is detected, preferring 2026 then 2025."
        ),
    )
    parser.add_argument("--workspace", type=Path, default=Path.cwd())
    parser.add_argument(
        "--parts-list",
        type=Path,
        default=None,
        help="Required CSV Parts List (defaults to 'Parts List.csv' in the workspace).",
    )
    parser.add_argument("--console-timeout", type=int, default=CONSOLE_TIMEOUT_SECONDS)
    parser.add_argument(
        "--workers",
        type=worker_count,
        default=2,
        help=(
            "Concurrent AutoCAD Core Console jobs for all drawings (default 2; "
            "no software cap). Detected bevel drawings are named with (B)."
        ),
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    script_dir = Path(__file__).resolve().parent
    workspace = args.workspace.resolve()
    parts_list = (args.parts_list or workspace / "Parts List.csv").resolve()
    try:
        acad_console = detect_autocad_console(args.acad_console_path)
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    required = {
        "Seed DWG": script_dir / "SPC_Seed.dwg",
        "ColorToLayer LISP": script_dir / "ColortoLayer.lsp",
        "accoreconsole.exe": acad_console,
        "Parts List CSV": parts_list,
    }
    missing = [(label, path) for label, path in required.items() if not path.is_file()]
    if missing:
        for label, path in missing:
            print(f"ERROR: {label} not found at {path}", file=sys.stderr)
        print("One or more required paths are missing. Script aborted.", file=sys.stderr)
        return 1

    try:
        parts = load_parts(parts_list)
    except (OSError, csv.Error, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    archive_dir = workspace / "_PROCESSED_DXF_ARCHIVE"
    log_dir = workspace / "_ORCHESTRATOR_LOGS"
    archive_dir.mkdir(exist_ok=True)
    log_dir.mkdir(exist_ok=True)
    counts = {"clean": 0, "bevel": 0, "failed": 0}

    print("=== STARTING CAD WORKFLOW AUTOMATION ORCHESTRATOR (PYTHON) ===")
    print(f"AutoCAD Core Console: {acad_console}")
    with tempfile.TemporaryDirectory(prefix="dxf-orchestrator-") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        h2d_path = temp_dir / "HashToDash.lsp"
        h2d_path.write_text(HASH_TO_DASH_LISP, encoding="ascii")
        input_files: list[Path] = []
        folders = sorted(
            (
                path for path in workspace.iterdir()
                if path.is_dir()
                and re.match(r"^\d+", path.name)
                and not re.match(r"^\d+-[A-Za-z]", path.name)
            ),
            key=lambda path: path.name,
        )
        for folder in folders:
            print(f"\nScanning Folder: {folder.name}")
            files = sorted(folder.glob("*.dxf"), key=lambda path: path.name.lower())
            if not files:
                print("  -> No DXF files found in this folder.")
            for dxf in files:
                print(f"  Processing: {dxf.name}...")
                input_files.append(dxf)

        processed_counts = process_files(
            input_files,
            parts=parts,
            workspace=workspace,
            archive_dir=archive_dir,
            log_dir=log_dir,
            acad_console=acad_console,
            lsp_path=required["ColorToLayer LISP"],
            seed_path=required["Seed DWG"],
            h2d_path=h2d_path,
            temp_dir=temp_dir,
            workers=args.workers,
            console_timeout=args.console_timeout,
        )
        for result, count in processed_counts.items():
            counts[result] += count

    if counts["failed"]:
        print("\n=== PROCESSING COMPLETE WITH FAILURES; ORIGINAL DXFS WERE KEPT ===")
    else:
        print("\n=== ALL PARTS PROCESSED, SORTED, AND SAVED! ===")
    print(
        f"=== Standard: {counts['clean']}  Bevel-marked: {counts['bevel']}  "
        f"Failed: {counts['failed']} ==="
    )
    print("=== Originals archived in _PROCESSED_DXF_ARCHIVE ===")
    return 1 if counts["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
