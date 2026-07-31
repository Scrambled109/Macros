from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

from .engine import ConversionError, convert_epls, load_metadata


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="EPLConverter",
        description="Convert explicitly supplied EPL workbooks into Plates and Shapes workbooks.",
    )
    parser.add_argument("--epl", nargs="+", help="One or more local EPL .xlsx files.")
    parser.add_argument(
        "--bop",
        nargs="+",
        help="One or more local BOP .xlsx/.xlsm files that define the in-scope parts.",
    )
    parser.add_argument("--output-dir", help="Folder for the new output files.")
    parser.add_argument("--name", help="Output base name.")
    parser.add_argument("--metadata", help="Optional part metadata .json or .csv.")
    parser.add_argument(
        "--plate-template",
        help="Optional Plates workbook whose first visible sheet supplies the header/layout style.",
    )
    parser.add_argument(
        "--request",
        help="Orchestrator request JSON file, or - to read the request from standard input.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Print a machine-readable JSON result (human-readable text is the default).",
    )
    return parser


def _load_request(value: str) -> dict[str, Any]:
    try:
        text = sys.stdin.read() if value == "-" else Path(value).read_text(encoding="utf-8-sig")
        payload = json.loads(text)
    except (OSError, json.JSONDecodeError) as exc:
        raise ConversionError(f"Could not read orchestrator request JSON: {exc}") from exc
    if not isinstance(payload, dict):
        raise ConversionError("The orchestrator request must be a JSON object.")
    return payload


def _request_metadata(payload: dict[str, Any]) -> dict[str, Any] | None:
    metadata = payload.get("metadata_by_part")
    if metadata is None:
        metadata = payload.get("metadata")
    if isinstance(metadata, str):
        return load_metadata(metadata)
    if metadata is not None and not isinstance(metadata, dict):
        raise ConversionError("metadata_by_part must be an object keyed by EPL Part No.")
    return metadata


def run_cli(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)
    try:
        if args.request:
            request = _load_request(args.request)
            epl_files = request.get("epl_files")
            bop_files = request.get("bop_files")
            output_dir = request.get("output_dir")
            output_base = request.get("output_base")
            plate_template = request.get("plate_template")
            metadata = _request_metadata(request)
        else:
            epl_files = args.epl
            bop_files = args.bop
            output_dir = args.output_dir
            output_base = args.name
            plate_template = args.plate_template
            metadata = load_metadata(args.metadata) if args.metadata else None
        if not epl_files or not bop_files or not output_dir or not output_base:
            raise ConversionError(
                "Supply --epl, --bop, --output-dir, and --name, "
                "or use --request with all required fields."
            )
        result = convert_epls(
            epl_files,
            output_dir,
            output_base,
            metadata,
            bop_paths=bop_files,
            plate_template_path=plate_template,
        )
        payload = result.to_dict()
        if args.json or args.request:
            print(json.dumps(payload, ensure_ascii=False))
        else:
            summary = payload["summary"]
            print(
                f"Converted {summary['input_files']} EPL file(s): "
                f"{summary['plates_exported']} plates, {summary['shapes_exported']} shapes."
            )
            for label, path in payload["outputs"].items():
                print(f"{label.replace('_', ' ').title()}: {path}")
        return 0
    except ConversionError as exc:
        error = {"ok": False, "error": str(exc)}
        if args.json or args.request:
            print(json.dumps(error, ensure_ascii=False))
        else:
            print(f"ERROR: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(run_cli())
