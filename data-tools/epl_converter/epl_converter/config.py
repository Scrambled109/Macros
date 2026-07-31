from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any


DEFAULT_CONFIG: dict[str, Any] = {
    "material_translations": [
        {
            "name": "HY-80",
            "patterns": [
                r"\bT9074\b.*\bHY[\s-]*80\b",
                r"\bGR[\s-]*HY[\s-]*80\b",
                r"\bHY[\s-]*80\b",
            ],
        },
        {
            "name": "MIL-S-22698_EH-36T",
            "patterns": [
                r"\bMIL[\s-]*S[\s-]*22698\b.*\bEH[\s-]*36T\b",
                r"\bGR[\s-]*EH[\s-]*36T\b",
            ],
        },
        {
            "name": "CRES 316L",
            "patterns": [
                r"\bCRES[\s-]*316L\b",
                r"\b316L\b",
                r"\bQQ[\s-]*S[\s-]*763\b.*\b316L\b",
            ],
        },
    ],
    "plate_keywords": ["PLATE", "SHEET", "SHIM", "WEDGE"],
    "shape_keywords": [
        "ANGLE",
        "BAR",
        "BEAM",
        "CHANNEL",
        "FLAT",
        "HOUSING",
        "PIPE",
        "ROD",
        "STUFFING TUBE",
        "TEE",
        "TUBE",
        "TUBING",
    ],
}


def application_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def load_config() -> dict[str, Any]:
    """Load editable settings beside the EXE/project, falling back to safe defaults."""
    path = application_dir() / "material_translations.json"
    if not path.exists():
        return DEFAULT_CONFIG
    try:
        supplied = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"Could not read {path.name}: {exc}") from exc
    merged = dict(DEFAULT_CONFIG)
    merged.update(supplied)
    return merged
