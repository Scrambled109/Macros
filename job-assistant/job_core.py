"""Core, UI-independent services for the Engineering Job Assistant."""

from __future__ import annotations

import csv
import json
import re
import shutil
from dataclasses import dataclass, asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


MANIFEST_NAME = "job_manifest.json"
MANIFEST_VERSION = 1

FOLDERS = {
    "source_boms": "01_Source_BOMs",
    "cut_files": "02_Cut_Files",
    "cad_models": "03_CAD_Models",
    "nests": "04_Nests",
    "exports": "05_Exports",
    "logs": "06_Logs",
    "reports": "07_Reports",
    "final_records": "08_Final_Records",
}

STAGES = [
    ("bom", "A-BOM conversion"),
    ("dxf", "DXF preparation"),
    ("plate_model", "Plate modeling"),
    ("manual_model", "Manual modeling"),
    ("assembly", "Assembly"),
    ("autobom", "AutoBOM"),
    ("model_review", "Model/Parts List review"),
    ("nesting", "Linear nesting"),
    ("exports", "Final exports"),
    ("comparison", "Production comparison"),
    ("final", "Final acceptance"),
]

REQUIRED_COLUMNS = {
    "parts": {
        "PART NUMBER",
        "DESCRIPTION",
        "TOTAL QUANTITY",
        "THICKNESS/SHAPE",
        "LENGTH",
        "MATERIAL",
    },
    "solidworks": {"FILE NAME", "QUANTITY", "SHAPE", "LENGTH", "MATERIAL"},
}
BOM_COLUMNS = {"ENG MAT ID", "DESCRIPTION", "QTY", "WIDTH", "LENGTH", "MTL TYPE"}


class JobError(RuntimeError):
    """A user-correctable job or manifest problem."""


@dataclass(frozen=True)
class Check:
    level: str
    code: str
    message: str
    correction: str = ""


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def safe_name(value: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", value.strip())
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    if not cleaned:
        raise JobError("Job number and name must contain usable characters.")
    return cleaned


def new_manifest(job_number: str, job_name: str, revision: str, root: Path) -> dict[str, Any]:
    job_number = safe_name(job_number)
    job_name = safe_name(job_name)
    revision = safe_name(revision)
    return {
        "manifest_version": MANIFEST_VERSION,
        "job": {
            "number": job_number,
            "name": job_name,
            "revision": revision,
            "created_at": utc_now(),
            "updated_at": utc_now(),
        },
        "root": str(root.resolve()),
        "paths": {key: str((root / folder).resolve()) for key, folder in FOLDERS.items()},
        "stages": {
            key: {
                "label": label,
                "status": "not_started",
                "started_at": None,
                "completed_at": None,
                "reviewed": False,
                "notes": "",
                "warnings_acknowledged": [],
                "artifacts": [],
            }
            for key, label in STAGES
        },
        "events": [],
    }


def save_manifest(manifest: dict[str, Any], path: Path | None = None) -> Path:
    target = path or Path(manifest["root"]) / MANIFEST_NAME
    manifest["job"]["updated_at"] = utc_now()
    temporary = target.with_suffix(".tmp")
    temporary.write_text(json.dumps(manifest, indent=2) + "\n", encoding="utf-8")
    temporary.replace(target)
    return target


def load_manifest(path: Path) -> dict[str, Any]:
    try:
        manifest = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise JobError(f"Could not read manifest: {exc}") from exc
    if manifest.get("manifest_version") != MANIFEST_VERSION:
        raise JobError("This job manifest uses an unsupported format version.")
    if not isinstance(manifest.get("stages"), dict) or not manifest.get("root"):
        raise JobError("The job manifest is incomplete.")
    return manifest


def create_job(
    parent: Path,
    job_number: str,
    job_name: str,
    revision: str,
    templates: Path | None = None,
) -> Path:
    root = parent / f"{safe_name(job_number)} - {safe_name(job_name)}"
    if root.exists() and any(root.iterdir()):
        raise JobError(f"The job folder already exists and is not empty: {root}")
    root.mkdir(parents=True, exist_ok=True)
    for folder in FOLDERS.values():
        (root / folder).mkdir(exist_ok=True)

    if templates and templates.exists():
        destination = root / FOLDERS["source_boms"]
        for source in templates.iterdir():
            if source.is_file():
                target = destination / source.name
                if target.exists():
                    raise JobError(f"Template copy would overwrite {target.name}.")
                shutil.copy2(source, target)

    manifest = new_manifest(job_number, job_name, revision, root)
    manifest["events"].append({"at": utc_now(), "type": "job_created"})
    return save_manifest(manifest)


def record_artifact(manifest: dict[str, Any], stage: str, path: Path) -> dict[str, Any]:
    if stage not in manifest["stages"]:
        raise JobError(f"Unknown workflow stage: {stage}")
    if not path.is_file():
        raise JobError(f"Artifact does not exist: {path}")
    stat = path.stat()
    artifact = {
        "path": str(path.resolve()),
        "name": path.name,
        "size": stat.st_size,
        "modified_at": datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat(timespec="seconds"),
        "recorded_at": utc_now(),
        "revision": manifest["job"]["revision"],
    }
    existing = manifest["stages"][stage]["artifacts"]
    existing[:] = [item for item in existing if item.get("path") != artifact["path"]]
    existing.append(artifact)
    return artifact


def start_stage(manifest: dict[str, Any], stage: str, warnings: Iterable[Check] = ()) -> None:
    item = manifest["stages"][stage]
    if item["status"] == "complete":
        raise JobError("A completed stage must be deliberately reopened before rerunning it.")
    if not item["started_at"]:
        item["started_at"] = utc_now()
    item["status"] = "in_progress"
    item["warnings_acknowledged"] = [asdict(check) for check in warnings]
    manifest["events"].append({"at": utc_now(), "type": "stage_started", "stage": stage})


def complete_stage(manifest: dict[str, Any], stage: str, notes: str) -> None:
    item = manifest["stages"][stage]
    if item["status"] != "in_progress":
        raise JobError("Start the stage before completing its review.")
    if not notes.strip():
        raise JobError("Enter a review note before marking the stage complete.")
    item["status"] = "complete"
    item["reviewed"] = True
    item["completed_at"] = utc_now()
    item["notes"] = notes.strip()
    manifest["events"].append({"at": utc_now(), "type": "stage_completed", "stage": stage})


def reopen_stage(manifest: dict[str, Any], stage: str, reason: str) -> None:
    if not reason.strip():
        raise JobError("A reason is required when reopening a completed stage.")
    item = manifest["stages"][stage]
    item["status"] = "in_progress"
    item["reviewed"] = False
    item["completed_at"] = None
    item["notes"] = reason.strip()
    manifest["events"].append(
        {"at": utc_now(), "type": "stage_reopened", "stage": stage, "reason": reason.strip()}
    )


def change_revision(manifest: dict[str, Any], revision: str) -> None:
    revision = safe_name(revision)
    old_revision = manifest["job"]["revision"]
    if revision == old_revision:
        return
    manifest["job"]["revision"] = revision
    manifest["events"].append(
        {
            "at": utc_now(),
            "type": "revision_changed",
            "from": old_revision,
            "to": revision,
        }
    )


def _files(path: str | Path, patterns: tuple[str, ...]) -> list[Path]:
    folder = Path(path)
    if not folder.is_dir():
        return []
    found: list[Path] = []
    for pattern in patterns:
        found.extend(folder.rglob(pattern))
    return sorted({item.resolve() for item in found if item.is_file()})


def _header(path: Path) -> set[str]:
    with path.open("r", encoding="utf-8-sig", newline="", errors="replace") as handle:
        sample = handle.read(8192)
        handle.seek(0)
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
        except csv.Error:
            dialect = csv.excel
        row = next(csv.reader(handle, dialect), [])
    return {str(value).strip().upper() for value in row if str(value).strip()}


def _duplicate_stems(paths: Iterable[Path]) -> list[str]:
    seen: dict[str, int] = {}
    for path in paths:
        key = path.stem.upper()
        seen[key] = seen.get(key, 0) + 1
    return sorted(key for key, count in seen.items() if count > 1)


def _bom_checks(paths: Iterable[Path]) -> list[Check]:
    try:
        import openpyxl
    except ImportError:
        return [Check("block", "openpyxl_missing", "The Excel validation package is not installed.", "Run the BOM converter requirements installation from SETUP.md.")]

    found_lofting = False
    invalid_messages: list[str] = []
    for path in paths:
        try:
            workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
        except (OSError, ValueError, KeyError) as exc:
            invalid_messages.append(f"{path.name}: {exc}")
            continue
        try:
            if "Lofting" not in workbook.sheetnames:
                continue
            found_lofting = True
            sheet = workbook["Lofting"]
            headers = {
                str(cell.value).strip().upper()
                for cell in next(sheet.iter_rows(min_row=1, max_row=1), ())
                if cell.value is not None
            }
            missing = sorted(BOM_COLUMNS - headers)
            if missing:
                invalid_messages.append(f"{path.name} Lofting sheet is missing: {', '.join(missing)}")
        finally:
            workbook.close()
    checks: list[Check] = []
    if not found_lofting:
        checks.append(Check("block", "lofting_sheet_missing", "No source workbook contains a Lofting sheet.", "Select the engineering A-BOM workbook, not a renamed CSV or template."))
    if invalid_messages:
        checks.append(Check("block", "invalid_bom_workbook", " | ".join(invalid_messages), "Correct the workbook before conversion."))
    return checks


def _part_number_checks(parts_csv: Path, model_files: Iterable[Path]) -> list[Check]:
    try:
        with parts_csv.open("r", encoding="utf-8-sig", newline="", errors="replace") as handle:
            reader = csv.DictReader(handle)
            field_map = {str(name).strip().upper(): name for name in (reader.fieldnames or [])}
            source_name = field_map.get("PART NUMBER")
            if not source_name:
                return []
            expected = {
                str(row.get(source_name, "")).strip().upper()
                for row in reader
                if str(row.get(source_name, "")).strip()
            }
    except OSError:
        return []
    actual = {path.stem.strip().upper() for path in model_files}
    if not expected or not actual:
        return []
    missing = sorted(expected - actual)
    extra = sorted(actual - expected)
    messages = []
    if missing:
        messages.append("Parts List numbers without matching model filenames: " + ", ".join(missing[:20]))
    if extra:
        messages.append("Model filenames absent from the Parts List: " + ", ".join(extra[:20]))
    return [Check("warning", "part_filename_mismatch", " | ".join(messages), "Confirm hardware/exclusions, then correct production filenames or the Parts List.")] if messages else []


def preflight(manifest: dict[str, Any], stage: str, config: dict[str, Any] | None = None) -> list[Check]:
    config = config or {}
    paths = manifest["paths"]
    checks: list[Check] = []

    for key, value in paths.items():
        if not Path(value).is_dir():
            checks.append(Check("block", f"missing_folder_{key}", f"Required folder is missing: {value}", "Restore or recreate the standard job folder."))

    stage_keys = [key for key, _ in STAGES]
    position = stage_keys.index(stage)
    if position > 0:
        previous = stage_keys[position - 1]
        if manifest["stages"][previous]["status"] != "complete":
            checks.append(
                Check(
                    "block",
                    "previous_stage_incomplete",
                    f"The previous stage, {manifest['stages'][previous]['label']}, is not complete.",
                    "Finish and record its required review checkpoint first.",
                )
            )

    requirements = {
        "bom": ("source_boms", ("*.xls", "*.xlsx", "*.xlsm"), "an A-BOM and Parts List template"),
        "dxf": ("cut_files", ("*.dxf",), "DXF input files"),
        "plate_model": ("cut_files", ("*.dwg",), "prepared DWG files"),
        "assembly": ("cad_models", ("*.SLDPRT", "*.sldprt"), "validated part models"),
        "autobom": ("cad_models", ("*.SLDASM", "*.sldasm"), "a SolidWorks assembly"),
        "nesting": ("exports", ("*.csv",), "a post-model Parts List CSV"),
        "exports": ("cad_models", ("*.SLDASM", "*.sldasm"), "the accepted assembly"),
        "comparison": ("exports", ("*.csv",), "Parts List and SolidWorks CSV exports"),
    }
    if stage in requirements:
        folder_key, patterns, description = requirements[stage]
        if not _files(paths[folder_key], patterns):
            checks.append(Check("block", f"missing_input_{stage}", f"No {description} found in {paths[folder_key]}.", "Place the accepted input in the standard folder."))

    if stage == "bom":
        workbooks = _files(paths["source_boms"], ("*.xlsx", "*.xlsm"))
        if len(workbooks) < 2:
            checks.append(Check("block", "missing_bom_or_template", "A-BOM conversion requires both a source workbook and a Parts List template.", "Place both workbooks in 01_Source_BOMs."))
        if workbooks:
            checks.extend(_bom_checks(workbooks))

    if stage in {"dxf", "plate_model"}:
        executable = config.get("autocad_gui" if stage == "dxf" else "autocad_console")
        if executable and not Path(executable).is_file():
            checks.append(Check("block", "autocad_missing", f"Configured AutoCAD executable was not found: {executable}", "Update local_config.json."))
        elif not executable:
            checks.append(Check("warning", "autocad_unconfigured", "AutoCAD is not configured in local_config.json.", "Copy local_config.example.json and set the installed path."))

    if stage in {"plate_model", "assembly", "autobom", "exports"}:
        executable = config.get("solidworks")
        if executable and not Path(executable).is_file():
            checks.append(Check("block", "solidworks_missing", f"Configured SolidWorks executable was not found: {executable}", "Update local_config.json."))
        elif not executable:
            checks.append(Check("warning", "solidworks_unconfigured", "SolidWorks is not configured in local_config.json.", "Copy local_config.example.json and set the installed path."))
    if stage == "plate_model" and not config.get("cad_batch_macro_rebuilt_for_environment_paths", False):
        checks.append(
            Check(
                "block",
                "cad_macro_not_rebuilt",
                "The compiled CAD batch macro is not confirmed to use Job Assistant folder configuration.",
                "Import the updated .bas modules, rebuild Main.RunBatch.swp, test it on disposable drawings, then set cad_batch_macro_rebuilt_for_environment_paths to true.",
            )
        )

    model_files = _files(paths["cad_models"], ("*.SLDPRT", "*.sldprt"))
    duplicates = _duplicate_stems(model_files)
    if duplicates:
        checks.append(Check("block", "duplicate_part_models", "Duplicate model part numbers: " + ", ".join(duplicates), "Remove or rename duplicate production models."))
    if stage in {"assembly", "autobom", "model_review", "nesting", "exports"}:
        part_exports = [path for path in _files(paths["exports"], ("*.csv",)) if "part" in path.name.lower()]
        if part_exports:
            checks.extend(_part_number_checks(part_exports[0], model_files))

    if stage in {"comparison", "final"}:
        export_files = _files(paths["exports"], ("*.csv",))
        parts = next((p for p in export_files if "part" in p.name.lower()), None)
        solidworks = next((p for p in export_files if any(word in p.name.lower() for word in ("solidworks", "assembly", "visualization"))), None)
        nests = _files(paths["nests"], ("*.csv", "*.txt"))
        for label, path, required in (("Parts List", parts, REQUIRED_COLUMNS["parts"]), ("SolidWorks", solidworks, REQUIRED_COLUMNS["solidworks"])):
            if not path:
                checks.append(Check("block", f"missing_{label.lower().replace(' ', '_')}_export", f"{label} CSV could not be identified.", "Include 'parts' or 'solidworks/assembly' in the accepted export filename."))
            elif path.stat().st_size == 0:
                checks.append(Check("block", "empty_export", f"{path.name} is empty.", "Create a fresh export."))
            else:
                missing = sorted(required - _header(path))
                if missing:
                    checks.append(Check("block", "missing_columns", f"{path.name} is missing columns: {', '.join(missing)}.", "Export the documented columns."))
        if not nests:
            checks.append(Check("block", "missing_nests", "No linear nest CSV/TXT exports were found.", "Place final nest exports in 04_Nests."))
        if parts and solidworks and nests:
            timestamps = [parts.stat().st_mtime, solidworks.stat().st_mtime, *[p.stat().st_mtime for p in nests]]
            if max(timestamps) - min(timestamps) > 7 * 24 * 3600:
                checks.append(Check("warning", "mixed_export_dates", "Final inputs span more than seven days and may represent mixed revisions.", "Confirm and record that all inputs belong to the manifest revision."))
        old_artifacts = [
            item
            for stage_item in manifest["stages"].values()
            for item in stage_item["artifacts"]
            if item.get("revision") != manifest["job"]["revision"]
        ]
        if old_artifacts:
            checks.append(Check("block", "stale_revision_artifacts", "Recorded artifacts from an older revision are still active.", "Re-export and record the affected artifacts for the current revision."))

    log_text = "\n".join(
        path.read_text(encoding="utf-8", errors="ignore")[-20000:]
        for path in _files(paths["logs"], ("*.log", "*.txt"))
    ).upper()
    if stage not in {"bom", "dxf"} and re.search(r"\b(FATAL|FAILED|ERROR)\b", log_text):
        checks.append(Check("warning", "unresolved_log_errors", "Job logs contain error or failure markers.", "Review and resolve or explicitly approve every logged exception."))

    if not checks:
        checks.append(Check("pass", "ready", "Preflight passed."))
    return checks


def load_local_config(repo_root: Path) -> dict[str, Any]:
    path = repo_root / "local_config.json"
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise JobError(f"local_config.json is invalid: {exc}") from exc
