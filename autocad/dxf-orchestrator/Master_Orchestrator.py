#!/usr/bin/env python3
"""Python implementation of the AutoCAD DXF master orchestrator.

This is an alternative entry point to Master_Orchestrator.ps1.  It intentionally
uses only the Python standard library and preserves the PowerShell workflow's
input, output, archive, logging, and manual bevel-review behavior.
"""

from __future__ import annotations

import argparse
import base64
import csv
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import NamedTuple


DEFAULT_ACAD_CONSOLE = Path(
    r"C:\Program Files\Autodesk\AutoCAD 2026\accoreconsole.exe"
)
DEFAULT_ACAD_GUI = Path(r"C:\Program Files\Autodesk\AutoCAD 2026\acad.exe")
CONSOLE_TIMEOUT_SECONDS = 180
BEVEL_REVIEW_TIMEOUT_SECONDS = 3600
MINIMUM_DWG_BYTES = 1024
AUTOMATION_RETRY_ATTEMPTS = 600

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

    Missing a manual gate is more costly than opening one unnecessarily, so
    this intentionally accepts punctuation, spacing, common shop synonyms,
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
            # Route them through manual review even when their annotation text
            # is stored in a referenced block rather than inline in the DXF.
            if entity_type in {"LEADER", "MLEADER"}:
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


def output_stamp(path: Path) -> tuple[int, int] | None:
    try:
        stat = path.stat()
    except OSError:
        return None
    if stat.st_size <= MINIMUM_DWG_BYTES:
        return None
    return stat.st_size, stat.st_mtime_ns


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


def open_review_drawing(acad_gui: Path, dxf: Path, script_path: Path) -> bool:
    """Open a review in the running AutoCAD session, starting it if necessary.

    Calling acad.exe for every drawing can create another application process,
    even though FINISH deliberately leaves the current process alive.  AutoCAD's
    COM automation object lets us add the next document to that existing session.
    PowerShell supplies the COM bridge while keeping this module dependency-free.
    """
    dxf_value = str(dxf.resolve()).replace("'", "''")
    script_value = autocad_path(script_path).replace("'", "''")
    acad_value = str(acad_gui.resolve()).replace("'", "''")
    powershell = f"""
$ErrorActionPreference = 'Stop'
$acad = $null
for ($attempt = 0; $attempt -lt 40; $attempt++) {{
    try {{
        $acad = [Runtime.InteropServices.Marshal]::GetActiveObject('AutoCAD.Application')
        break
    }} catch {{
        # AutoCAD can have a process before it has registered its COM object.
        # Give that process time to finish starting rather than launching a
        # second instance immediately.
        if (-not (Get-Process -Name acad -ErrorAction SilentlyContinue)) {{ break }}
        Start-Sleep -Milliseconds 500
    }}
}}
if ($null -eq $acad) {{
    if (Get-Process -Name acad -ErrorAction SilentlyContinue) {{ exit 2 }}
    # Start an empty application and open the document through that process's
    # COM object. Passing the DXF to acad.exe can both forward the file to an
    # existing session and leave the newly started process as a blank second
    # instance, which then displays a misleading read-only prompt.
    Start-Process -FilePath '{acad_value}'
    for ($attempt = 0; $attempt -lt 120; $attempt++) {{
        Start-Sleep -Milliseconds 500
        try {{
            $acad = [Runtime.InteropServices.Marshal]::GetActiveObject('AutoCAD.Application')
            break
        }} catch {{ }}
    }}
    if ($null -eq $acad) {{ exit 3 }}
}}
$acad.Visible = $true
# Opening a document and sending its script are deliberately separate retry
# phases. Documents.Open can succeed before AutoCAD becomes busy running the
# drawing's startup work. Retrying the *whole* operation at that point opens a
# second, read-only tab for the same DXF.
$document = $null
for ($attempt = 0; $attempt -lt {AUTOMATION_RETRY_ATTEMPTS}; $attempt++) {{
    try {{
        # A previous automation attempt may have opened the drawing and then
        # lost its COM call while AutoCAD was busy. Reuse that tab rather than
        # opening a duplicate/read-only copy when this operation is retried.
        foreach ($candidate in $acad.Documents) {{
            if ($candidate.FullName -eq '{dxf_value}') {{
                $document = $candidate
                break
            }}
        }}
        if ($null -eq $document) {{
            $document = $acad.Documents.Open('{dxf_value}')
        }}
        break
    }} catch {{
        Start-Sleep -Milliseconds 500
    }}
}}
if ($null -eq $document) {{ exit 4 }}

# AutoCAD may reject automation calls temporarily while the document finishes
# opening. Keep the document reference and retry only activation/script setup;
# never call Documents.Open again for this review.
for ($attempt = 0; $attempt -lt {AUTOMATION_RETRY_ATTEMPTS}; $attempt++) {{
    try {{
        $document.Activate()
        $document.SetVariable('FILEDIA', 0)
        $document.SendCommand("_.SCRIPT`n`\"{script_value}`\"`n")
        exit 0
    }} catch {{
        Start-Sleep -Milliseconds 500
    }}
}}
exit 4
"""
    encoded = base64.b64encode(powershell.encode("utf-16le")).decode("ascii")
    attached = subprocess.run(
        [
            "powershell.exe", "-NoProfile", "-NonInteractive",
            "-EncodedCommand", encoded,
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        check=False,
    )
    if attached.returncode == 0:
        print("    --> Review opened through the AutoCAD automation session.")
        return True

    if attached.returncode == 2:
        print(
            "    [X] AutoCAD is open but its automation interface is not "
            "available. Close that session or run this tool at the same "
            "privilege level; a second AutoCAD instance was not opened."
        )
        return False

    print(
        "    [X] AutoCAD did not publish its automation interface in time. "
        "The review drawing was not opened."
    )
    return False


class ReviewJob(NamedTuple):
    dxf: Path
    final_dwg: Path
    target_name: str
    previous_stamp: tuple[int, int] | None
    started_at: float


def output_details(
    dxf: Path, parts: dict[str, dict[str, str]], workspace: Path
) -> tuple[str, Path]:
    """Return the sorting directory name and final DWG path for a DXF."""
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
    target_dir = workspace / target_name
    target_dir.mkdir(exist_ok=True)
    return target_name, target_dir / f"{working_name}.dwg"


def archive_original(source: Path, archive_dir: Path) -> None:
    destination = archive_dir / source.name
    if destination.exists():
        destination.unlink()
    shutil.move(str(source), str(destination))


def process_file(
    dxf: Path,
    parts: dict[str, dict[str, str]],
    workspace: Path,
    archive_dir: Path,
    log_dir: Path,
    acad_console: Path,
    acad_gui: Path,
    lsp_path: Path,
    seed_path: Path,
    h2d_path: Path,
    script_path: Path,
    console_timeout: int,
    review_timeout: int,
) -> str:
    target_name, final_dwg = output_details(dxf, parts, workspace)
    escaped_dwg = autocad_path(final_dwg)
    preamble = [
        "FILEDIA", "0", "SECURELOAD", "0", "-INSERT",
        f'"{autocad_path(seed_path)}"', "0,0", "1", "1", "0",
        "ERASE", "L", "", f'(load "{autocad_path(lsp_path)}")',
        "ColorToLayer", f'(load "{autocad_path(h2d_path)}")', "HashToDash",
    ]

    if is_beveled_dxf(dxf):
        print("    [!] BEVEL DETECTED. Loading into AutoCAD for manual check...")
        finish_lisp = (
            '(defun c:SPCFINISH ( / oldFileDia) '
            '(setq oldFileDia (getvar "FILEDIA")) (setvar "FILEDIA" 0) '
            f'(if (findfile "{escaped_dwg}") '
            f'(command "_.SAVEAS" "2018" "{escaped_dwg}" "Y") '
            f'(command "_.SAVEAS" "2018" "{escaped_dwg}")) '
            '(setvar "FILEDIA" oldFileDia) (command "_.CLOSE") (princ))'
        )
        write_script(script_path, preamble + [
            finish_lisp, "SECURELOAD", "1", "FILEDIA", "1",
        ])
        previous = output_stamp(final_dwg)
        if not open_review_drawing(acad_gui, dxf, script_path):
            return "failed"
        print("    --> Review in AutoCAD, then type SPCFINISH (Enter) to save & sort...")
        deadline = time.monotonic() + review_timeout
        saved = False
        while time.monotonic() < deadline:
            current = output_stamp(final_dwg)
            if current is not None and current != previous:
                saved = True
                break
            time.sleep(0.5)
        if saved and wait_file_stable(final_dwg):
            archive_original(dxf, archive_dir)
            print(f"    [SUCCESS] Saved and sorted to {target_name}")
            return "bevel"
        print(f"    [X] SPCFINISH did not produce a valid DWG within {review_timeout}s. Original DXF kept.")
        return "failed"

    print("    [+] Clean Part. Running in background core console...")
    write_script(script_path, preamble + [
        "_SAVEAS", "2018", f'"{escaped_dwg}"', "SECURELOAD", "1",
        "FILEDIA", "1", "_QUIT", "Y",
    ])
    log_path = log_dir / f"{dxf.stem}.log"
    error_path = log_dir / f"{dxf.stem}.err.log"
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
        print(f"    [SUCCESS] Converted and sorted to {target_name}")
        return "clean"
    print(f"    [X] ERROR: exit={exit_code}; DWG missing or too small. Original DXF kept. Log: {log_path}")
    return "failed"


def launch_review(
    dxf: Path,
    parts: dict[str, dict[str, str]],
    workspace: Path,
    acad_gui: Path,
    lsp_path: Path,
    seed_path: Path,
    h2d_path: Path,
    script_path: Path,
) -> ReviewJob | None:
    """Prepare and open one beveled drawing without waiting for its review."""
    target_name, final_dwg = output_details(dxf, parts, workspace)
    escaped_dwg = autocad_path(final_dwg)
    finish_lisp = (
        '(defun c:SPCFINISH ( / oldFileDia) '
        '(setq oldFileDia (getvar "FILEDIA")) (setvar "FILEDIA" 0) '
        f'(if (findfile "{escaped_dwg}") '
        f'(command "_.SAVEAS" "2018" "{escaped_dwg}" "Y") '
        f'(command "_.SAVEAS" "2018" "{escaped_dwg}")) '
        '(setvar "FILEDIA" oldFileDia) (command "_.CLOSE") (princ))'
    )
    preamble = [
        "FILEDIA", "0", "SECURELOAD", "0", "-INSERT",
        f'"{autocad_path(seed_path)}"', "0,0", "1", "1", "0",
        "ERASE", "L", "", f'(load "{autocad_path(lsp_path)}")',
        "ColorToLayer", f'(load "{autocad_path(h2d_path)}")', "HashToDash",
    ]
    write_script(
        script_path,
        preamble + [finish_lisp, "SECURELOAD", "1", "FILEDIA", "1"],
    )
    previous = output_stamp(final_dwg)
    if not open_review_drawing(acad_gui, dxf, script_path):
        return None
    return ReviewJob(dxf, final_dwg, target_name, previous, time.monotonic())


def wait_for_reviews(
    jobs: list[ReviewJob], archive_dir: Path, review_timeout: int
) -> tuple[int, int]:
    """Watch all open review tabs, allowing the operator to finish any order."""
    pending = list(jobs)
    completed = failed = 0
    while pending:
        now = time.monotonic()
        for job in pending[:]:
            current = output_stamp(job.final_dwg)
            if current is not None and current != job.previous_stamp:
                if wait_file_stable(job.final_dwg):
                    archive_original(job.dxf, archive_dir)
                    print(f"    [SUCCESS] {job.dxf.name} saved and sorted to {job.target_name}")
                    completed += 1
                    pending.remove(job)
            elif now - job.started_at >= review_timeout:
                print(
                    f"    [X] {job.dxf.name}: SPCFINISH did not produce a valid "
                    f"DWG within {review_timeout}s. Original DXF kept."
                )
                failed += 1
                pending.remove(job)
        if pending:
            time.sleep(0.5)
    return completed, failed


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--acad-console-path", type=Path, default=DEFAULT_ACAD_CONSOLE)
    parser.add_argument("--acad-gui-path", type=Path, default=DEFAULT_ACAD_GUI)
    parser.add_argument("--workspace", type=Path, default=Path.cwd())
    parser.add_argument(
        "--parts-list",
        type=Path,
        default=None,
        help="Required CSV Parts List (defaults to 'Parts List.csv' in the workspace).",
    )
    parser.add_argument("--console-timeout", type=int, default=CONSOLE_TIMEOUT_SECONDS)
    parser.add_argument("--review-timeout", type=int, default=BEVEL_REVIEW_TIMEOUT_SECONDS)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    script_dir = Path(__file__).resolve().parent
    workspace = args.workspace.resolve()
    parts_list = (args.parts_list or workspace / "Parts List.csv").resolve()
    required = {
        "Seed DWG": script_dir / "SPC_Seed.dwg",
        "ColorToLayer LISP": script_dir / "ColorToLayer.lsp",
        "accoreconsole.exe": args.acad_console_path,
        "acad.exe": args.acad_gui_path,
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
    with tempfile.TemporaryDirectory(prefix="dxf-orchestrator-") as temp_dir_name:
        temp_dir = Path(temp_dir_name)
        h2d_path = temp_dir / "HashToDash.lsp"
        h2d_path.write_text(HASH_TO_DASH_LISP, encoding="ascii")
        script_path = temp_dir / "automation_job.scr"
        review_files: list[Path] = []
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
                if is_beveled_dxf(dxf):
                    print("    [!] BEVEL DETECTED. Queued for manual check...")
                    review_files.append(dxf)
                    continue
                result = process_file(
                    dxf, parts, workspace, archive_dir, log_dir,
                    args.acad_console_path, args.acad_gui_path,
                    required["ColorToLayer LISP"], required["Seed DWG"],
                    h2d_path, script_path, args.console_timeout, args.review_timeout,
                )
                counts[result] += 1

        if review_files:
            print(
                f"\n=== OPENING {len(review_files)} BEVEL REVIEWS IN AUTOCAD ==="
            )
            jobs: list[ReviewJob] = []
            for index, dxf in enumerate(review_files):
                review_script = temp_dir / f"review_{index}.scr"
                job = launch_review(
                    dxf, parts, workspace, args.acad_gui_path,
                    required["ColorToLayer LISP"], required["Seed DWG"],
                    h2d_path, review_script,
                )
                if job is None:
                    counts["failed"] += 1
                else:
                    jobs.append(job)
            if jobs:
                print(
                    "\n--> All available reviews are open. Review each tab, then "
                    "type SPCFINISH (Enter) in that tab. You may finish them in any order."
                )
                completed, failed = wait_for_reviews(
                    jobs, archive_dir, args.review_timeout
                )
                counts["bevel"] += completed
                counts["failed"] += failed

    if counts["failed"]:
        print("\n=== PROCESSING COMPLETE WITH FAILURES; ORIGINAL DXFS WERE KEPT ===")
    else:
        print("\n=== ALL PARTS PROCESSED, SORTED, AND SAVED! ===")
    print(f"=== Clean: {counts['clean']}  Bevel: {counts['bevel']}  Failed: {counts['failed']} ===")
    print("=== Originals archived in _PROCESSED_DXF_ARCHIVE ===")
    return 1 if counts["failed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
