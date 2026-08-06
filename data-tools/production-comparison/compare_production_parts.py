"""
Production Part Reconciliation
==============================

Compares:
    1. SolidWorks Assembly Visualization export (engineering baseline)
    2. Drawing Parts List export
    3. Every linear-material nesting CSV in a selected folder

SolidWorks is treated as the required quantity and geometry baseline.

Classification rules:
    - PLATE: identified from the Parts List DESCRIPTION field.
      Compared between SolidWorks and the Parts List. Linear nesting is
      intentionally not required yet.
    - EXCLUDED HARDWARE: bolts, nuts, washers, screws, etc.
      Not expected in linear nesting.
    - LINEAR: every other classified part.
      Required in SolidWorks, Parts List, and linear nesting.
    - UNKNOWN: cannot be classified because description data is unavailable.

The nesting files may display in one Excel column. That is okay: they are
semicolon-delimited and this script parses them directly.

No third-party Python packages are required.

OUTPUT
------
A timestamped folder containing:
    comparison_summary.json
    comparison_report.html
    production_part_comparison.xlsx
    all_comparisons.csv
    errors_requiring_action.csv
    technical_details.csv
    missing_from_solidworks_or_parts_list.csv
    not_checked.csv
    exact_matches.csv
    source_data_issues.csv
    source_rows.csv
    run_summary.txt

NORMAL USE
----------
Double-click this file or run:

    py compare_production_parts.py

COMMAND-LINE USE
----------------
    py compare_production_parts.py ^
        --nests "C:\\Job\\Nests" ^
        --parts "C:\\Job\\parts_list.csv" ^
        --solidworks "C:\\Job\\solidworks.csv" ^
        --output "C:\\Job\\Comparison Reports"

INPUT REQUIREMENTS
------------------
Parts List:
    PART NUMBER
    DESCRIPTION
    TOTAL QUANTITY
    THICKNESS/SHAPE
    LENGTH
    MATERIAL
    WIDTH is optional

SolidWorks Assembly Visualization:
    FILE NAME
    QUANTITY
    SHAPE
    LENGTH
    MATERIAL
    DESCRIPTION and WIDTH are optional

The program will continue if a requested comparison column is missing, but
affected rows will be placed under NOT CHECKED and the missing column will be
listed under SOURCE DATA ISSUES. This prevents incomplete data from being
reported as an exact match.
"""

from __future__ import annotations

import argparse
import csv
import html
import json
import math
import re
import sys
import webbrowser
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import datetime, timezone
from fractions import Fraction
from pathlib import Path
from typing import Any, Iterable, Optional
from xml.sax.saxutils import escape as xml_escape


# =============================================================================
# DEFAULT RULES
# =============================================================================

DEFAULT_RULES: dict[str, Any] = {
    "length_tolerance_inches": 0.02,
    "width_tolerance_inches": 0.02,
    "quantity_tolerance": 0.0,

    "plate_description_keywords": [
        "PLATE"
    ],

    "excluded_hardware_description_keywords": [
        "BOLT",
        "NUT",
        "WASHER",
        "SCREW",
        "HARDWARE",
        "FASTENER",
        "THREADED ROD"
    ],

    "solidworks_configuration_suffixes": [
        "DEFAULT<AS MACHINED>",
        "DEFAULT OF DEFAULT MACHINED",
        "DEFAULT OF DEFAULT",
        "DEFAULT MACHINED",
        "DEFAULT"
    ],

    "part_number_aliases": {},

    "shape_aliases": {
        "ANGLE2X2X.25": "AL2X2X.25",
        "L2X2X.25": "AL2X2X.25"
    },

    "material_aliases": {
        "A500GRADEB": "A500GRB",
        "A500GRADEBSTEEL": "A500GRB",
        "A500GR.B": "A500GRB",
        "A500GRBSTEEL": "A500GRB",
        "A36STEEL": "A36"
    },

    "header_aliases": {
        "part_number": [
            "PART NUMBER",
            "PART NO",
            "PART NUM",
            "FILE NAME",
            "FILENAME"
        ],
        "quantity": [
            "TOTAL QUANTITY",
            "TOTAL QTY",
            "QUANTITY",
            "QTY"
        ],
        "description": [
            "DESCRIPTION",
            "PART DESCRIPTION",
            "DESC"
        ],
        "shape": [
            "THICKNESS SHAPE",
            "THICKNESS/SHAPE",
            "SHAPE",
            "PROFILE",
            "SECTION"
        ],
        "length": [
            "LENGTH",
            "PART LENGTH",
            "CUT LENGTH"
        ],
        "width": [
            "WIDTH",
            "PART WIDTH"
        ],
        "material": [
            "MATERIAL",
            "MATERIAL TYPE",
            "MATERIAL SPEC",
            "MATERIAL GRADE",
            "GRADE"
        ],
        "configuration": [
            "CONFIGURATION",
            "CONFIGURATION NAME",
            "CONFIG"
        ]
    },

    "checks": {
        "quantity": True,
        "length": True,
        "shape": True,
        "material": True,
        "width_when_available": True
    },

    "output": {
        "create_timestamped_subfolder": True,
        "open_html_report_when_finished": True,
        "create_excel_workbook": True
    }
}


# =============================================================================
# DATA TYPES
# =============================================================================

@dataclass
class Issue:
    severity: str
    issue: str
    source: str
    file: str = ""
    line: str = ""
    part_number: str = ""
    details: str = ""

    def as_dict(self) -> dict[str, str]:
        return {
            "Severity": self.severity,
            "Issue": self.issue,
            "Source": self.source,
            "File": self.file,
            "Line": self.line,
            "Part Number": self.part_number,
            "Details": self.details,
        }


@dataclass
class SourceRow:
    source: str
    source_file: str
    source_line: int

    part_number_raw: str
    part_number: str

    quantity_raw: str = ""
    quantity: Optional[float] = None

    description_raw: str = ""

    length_raw: str = ""
    length: Optional[float] = None

    width_raw: str = ""
    width: Optional[float] = None

    shape_raw: str = ""
    shape_key: str = ""

    material_raw: str = ""
    material_key: str = ""

    configuration_raw: str = ""
    stock_description_raw: str = ""


@dataclass
class ParsedSource:
    name: str
    path: Path
    rows: list[SourceRow]
    available_fields: set[str]
    selected_headers: dict[str, str]
    issues: list[Issue]


@dataclass
class Aggregate:
    source: str
    part_number: str
    rows: list[SourceRow] = field(default_factory=list)

    quantity: float = 0.0

    descriptions: list[str] = field(default_factory=list)
    lengths: list[float] = field(default_factory=list)
    length_raw_values: list[str] = field(default_factory=list)

    widths: list[float] = field(default_factory=list)
    width_raw_values: list[str] = field(default_factory=list)

    shape_raw_values: list[str] = field(default_factory=list)
    shape_keys: list[str] = field(default_factory=list)

    material_raw_values: list[str] = field(default_factory=list)
    material_keys: list[str] = field(default_factory=list)

    source_locations: list[str] = field(default_factory=list)
    stock_descriptions: list[str] = field(default_factory=list)

    @property
    def representative_length(self) -> Optional[float]:
        return self.lengths[0] if self.lengths else None

    @property
    def representative_width(self) -> Optional[float]:
        return self.widths[0] if self.widths else None

    @property
    def representative_shape_key(self) -> str:
        return self.shape_keys[0] if self.shape_keys else ""

    @property
    def representative_material_key(self) -> str:
        return self.material_keys[0] if self.material_keys else ""

    @property
    def descriptions_text(self) -> str:
        return " | ".join(unique_preserve_order(v for v in self.descriptions if v))

    @property
    def raw_lengths(self) -> str:
        return " | ".join(unique_preserve_order(v for v in self.length_raw_values if v))

    @property
    def raw_widths(self) -> str:
        return " | ".join(unique_preserve_order(v for v in self.width_raw_values if v))

    @property
    def raw_shapes(self) -> str:
        return " | ".join(unique_preserve_order(v for v in self.shape_raw_values if v))

    @property
    def raw_materials(self) -> str:
        return " | ".join(unique_preserve_order(v for v in self.material_raw_values if v))

    @property
    def locations_text(self) -> str:
        return " | ".join(unique_preserve_order(self.source_locations))

    @property
    def stock_text(self) -> str:
        return " | ".join(unique_preserve_order(v for v in self.stock_descriptions if v))


# =============================================================================
# RULE LOADING
# =============================================================================

def deep_merge(base: dict[str, Any], updates: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in base.items():
        if isinstance(value, dict):
            result[key] = deep_merge(value, {})
        elif isinstance(value, list):
            result[key] = list(value)
        else:
            result[key] = value

    for key, value in updates.items():
        if (
            key in result
            and isinstance(result[key], dict)
            and isinstance(value, dict)
        ):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value

    return result


def load_rules(path: Path) -> tuple[dict[str, Any], list[Issue]]:
    issues: list[Issue] = []

    if not path.exists():
        issues.append(
            Issue(
                severity="WARNING",
                issue="Rules file not found",
                source="Configuration",
                file=str(path),
                details="Built-in default rules were used.",
            )
        )
        return deep_merge(DEFAULT_RULES, {}), issues

    try:
        loaded = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        issues.append(
            Issue(
                severity="ERROR",
                issue="Rules file could not be read",
                source="Configuration",
                file=str(path),
                details=f"{exc}. Built-in default rules were used.",
            )
        )
        return deep_merge(DEFAULT_RULES, {}), issues

    if not isinstance(loaded, dict):
        issues.append(
            Issue(
                severity="ERROR",
                issue="Rules file root must be a JSON object",
                source="Configuration",
                file=str(path),
                details="Built-in default rules were used.",
            )
        )
        return deep_merge(DEFAULT_RULES, {}), issues

    return deep_merge(DEFAULT_RULES, loaded), issues


# =============================================================================
# GENERAL UTILITIES
# =============================================================================

def unique_preserve_order(values: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []

    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)

    return result


def clean_header(value: str) -> str:
    text = str(value or "").replace("\ufeff", "").strip().upper()
    text = re.sub(r"[\s_/\\-]+", " ", text)
    return re.sub(r"\s+", " ", text)


def canonical_header_aliases(rules: dict[str, Any]) -> dict[str, set[str]]:
    aliases: dict[str, set[str]] = {}

    for field_name, values in rules["header_aliases"].items():
        aliases[field_name] = {clean_header(value) for value in values}

    return aliases


def read_text(path: Path) -> str:
    if path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        raise ValueError(
            f"{path.name} is an Excel workbook. Export that sheet as CSV UTF-8 "
            "before running this script."
        )

    errors: list[str] = []

    for encoding in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=encoding)
        except (UnicodeError, UnicodeDecodeError) as exc:
            errors.append(f"{encoding}: {exc}")

    raise UnicodeError(
        f"Unable to decode {path}. Attempts: {'; '.join(errors)}"
    )


def detect_delimiter(text: str) -> str:
    if re.search(
        r"(?im)^\s*Parts\s*;\s*Length\s*;\s*(Quantity|Qty)",
        text,
    ):
        return ";"

    sample_lines = [line for line in text.splitlines() if line.strip()][:40]
    sample = "\n".join(sample_lines)

    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;\t|")
        return dialect.delimiter
    except csv.Error:
        counts = {delimiter: sample.count(delimiter) for delimiter in ",;\t|"}
        return max(counts, key=counts.get)


def read_delimited_rows(path: Path) -> list[list[str]]:
    text = read_text(path)
    delimiter = detect_delimiter(text)

    return [
        [cell.strip() for cell in row]
        for row in csv.reader(text.splitlines(), delimiter=delimiter)
    ]


def get_cell(row: list[str], index: Optional[int]) -> str:
    if index is None or index >= len(row):
        return ""
    return row[index].strip()


def parse_number(value: Any) -> Optional[float]:
    if value is None:
        return None

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    text = (
        text.upper()
        .replace(",", "")
        .replace('"', "")
        .replace("″", "")
        .replace("INCHES", "")
        .replace("INCH", "")
        .replace(" IN", "")
        .strip()
    )

    try:
        return float(text)
    except ValueError:
        pass

    mixed = re.fullmatch(r"([+-]?\d+)\s+(\d+)\s*/\s*(\d+)", text)
    if mixed:
        whole, numerator, denominator = map(int, mixed.groups())
        if denominator == 0:
            return None
        fraction = numerator / denominator
        return whole + fraction if whole >= 0 else whole - fraction

    fraction = re.fullmatch(r"([+-]?\d+)\s*/\s*(\d+)", text)
    if fraction:
        numerator, denominator = map(int, fraction.groups())
        if denominator == 0:
            return None
        return numerator / denominator

    feet_inches = re.fullmatch(
        r"([+-]?\d+)\s*'\s*(\d+(?:\.\d+)?(?:\s+\d+/\d+)?)\s*",
        text,
    )
    if feet_inches:
        feet = int(feet_inches.group(1))
        inches = parse_number(feet_inches.group(2))
        if inches is None:
            return None
        return feet * 12 + inches

    return None


def format_quantity(value: Optional[float]) -> str:
    if value is None:
        return ""

    if abs(value - round(value)) < 1e-9:
        return str(int(round(value)))

    return f"{value:.6f}".rstrip("0").rstrip(".")


def format_decimal(value: Optional[float]) -> str:
    if value is None:
        return ""

    return f"{value:.6f}".rstrip("0").rstrip(".")


def format_fraction(value: Optional[float], denominator: int = 64) -> str:
    if value is None:
        return ""

    sign = "-" if value < 0 else ""
    absolute = abs(value)
    whole = int(math.floor(absolute))
    fractional = Fraction(absolute - whole).limit_denominator(denominator)

    if fractional.numerator == 0:
        return f"{sign}{whole}"

    if fractional.numerator == fractional.denominator:
        return f"{sign}{whole + 1}"

    if whole:
        return (
            f"{sign}{whole} "
            f"{fractional.numerator}/{fractional.denominator}"
        )

    return f"{sign}{fractional.numerator}/{fractional.denominator}"


def difference(actual: Optional[float], baseline: Optional[float]) -> str:
    if actual is None or baseline is None:
        return ""

    return format_decimal(actual - baseline)


def number_matches(
    actual: Optional[float],
    baseline: Optional[float],
    tolerance: float,
) -> Optional[bool]:
    if actual is None or baseline is None:
        return None

    return abs(actual - baseline) <= tolerance


# =============================================================================
# PART NUMBER NORMALIZATION
# =============================================================================

def normalize_part_number(
    value: str,
    source: str,
    rules: dict[str, Any],
) -> str:
    text = str(value or "").replace("\ufeff", "").strip()

    if not text:
        return ""

    # Remove folder paths if a full SolidWorks file path was exported.
    text = re.split(r"[\\/]", text)[-1]

    text = re.sub(
        r"\.(SLDPRT|SLDASM)$",
        "",
        text,
        flags=re.IGNORECASE,
    )

    if source == "SolidWorks":
        suffixes = [
            str(value).strip()
            for value in rules["solidworks_configuration_suffixes"]
        ]
        canonical_suffixes = {
            re.sub(r"[^A-Z0-9]+", "", suffix.upper()) for suffix in suffixes
        }

        # SolidWorks frequently exports configurations in parentheses using
        # punctuation that varies by version, for example
        # ``(Default<As Machined>)``. Compare punctuation-free keys rather than
        # requiring one exact spelling from the rules file.
        bracketed = re.search(r"\s*[\(\[]([^\(\)\[\]]+)[\)\]]\s*$", text)
        if bracketed:
            key = re.sub(r"[^A-Z0-9]+", "", bracketed.group(1).upper())
            if key in canonical_suffixes:
                text = text[: bracketed.start()].rstrip()

        for suffix in sorted(suffixes, key=len, reverse=True):
            escaped = re.escape(suffix)
            for pattern in (
                rf"\s*@\s*{escaped}\s*$",
                rf"\s*[-_]\s*{escaped}\s*$",
                rf"\s+{escaped}\s*$",
            ):
                text = re.sub(pattern, "", text, flags=re.IGNORECASE)

        # Remove SolidWorks component instance markers such as <1> after the
        # optional configuration suffix has been removed.
        text = re.sub(r"<\d+>$", "", text.strip())

    if re.fullmatch(r"[+-]?\d+\.0", text):
        text = text[:-2]

    # The Parts List template represents the drawing/part separator as ``#``;
    # nesting and SolidWorks exports use ``-`` for the same identifier.
    text = re.sub(r"\s+", "", text).upper().replace("#", "-")

    aliases = {
        str(key).strip().upper().replace("#", "-"):
        str(value).strip().upper().replace("#", "-")
        for key, value in rules["part_number_aliases"].items()
    }

    return aliases.get(text, text)


# =============================================================================
# SHAPE AND MATERIAL NORMALIZATION
# =============================================================================

def replace_fraction_tokens(text: str) -> str:
    mixed_pattern = re.compile(r"(?<![\d.])(\d+)\s+(\d+)\s*/\s*(\d+)")
    simple_pattern = re.compile(r"(?<![\d.])(\d+)\s*/\s*(\d+)")

    def mixed_replacement(match: re.Match[str]) -> str:
        whole, numerator, denominator = map(int, match.groups())
        if denominator == 0:
            return match.group(0)
        return normalize_numeric_token(whole + numerator / denominator)

    def simple_replacement(match: re.Match[str]) -> str:
        numerator, denominator = map(int, match.groups())
        if denominator == 0:
            return match.group(0)
        return normalize_numeric_token(numerator / denominator)

    text = mixed_pattern.sub(mixed_replacement, text)
    text = simple_pattern.sub(simple_replacement, text)
    return text


def normalize_numeric_token(value: float) -> str:
    rounded = round(value, 8)

    if abs(rounded - round(rounded)) < 1e-9:
        return str(int(round(rounded)))

    text = f"{rounded:.8f}".rstrip("0").rstrip(".")

    if text.startswith("0."):
        return text[1:]

    if text.startswith("-0."):
        return "-" + text[2:]

    return text


def normalize_shape(value: str, rules: dict[str, Any]) -> str:
    text = str(value or "").upper().strip()

    if not text:
        return ""

    text = text.replace("×", "X")
    text = text.replace("*", "X")
    text = replace_fraction_tokens(text)

    replacements = [
        (r"\bANGLE IRON\b", "AL"),
        (r"\bANGLE\b", "AL"),
        (r"\bSQ(?:UARE)?\s*TUBE\b", "SQTUBE"),
        (r"\bRECT(?:ANGULAR)?\s*TUBE\b", "RECTTUBE"),
        (r"\bROUND\s*TUBE\b", "ROUNDTUBE"),
        (r"\bROUND\s*BAR\b", "ROUNDBAR"),
        (r"\bFLAT\s*BAR\b", "FLATBAR"),
        (r"\bT[\s-]*BAR\b", "TBAR"),
        (r"\bI[\s-]*BEAM\b", "IBEAM"),
    ]

    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)

    # Normalize every decimal token independently.
    def decimal_replacement(match: re.Match[str]) -> str:
        try:
            return normalize_numeric_token(float(match.group(0)))
        except ValueError:
            return match.group(0)

    text = re.sub(r"(?<![A-Z])\d*\.\d+|\d+\.(?![A-Z])", decimal_replacement, text)

    text = re.sub(r"[\s,_\-]+", "", text)
    text = text.replace('"', "").replace("'", "")

    aliases = {
        normalize_shape_alias_key(key): normalize_shape_alias_key(value)
        for key, value in rules["shape_aliases"].items()
    }

    return aliases.get(text, text)


def normalize_shape_alias_key(value: str) -> str:
    text = str(value or "").upper().replace("×", "X").replace("*", "X")
    text = replace_fraction_tokens(text)
    text = re.sub(r"[\s,_\-]+", "", text)
    return text.replace('"', "").replace("'", "")


def normalize_material(value: str, rules: dict[str, Any]) -> str:
    text = str(value or "").upper().strip()

    if not text:
        return ""

    text = text.replace("GRADE", "GR")
    text = re.sub(r"\bASTM\b", "", text)
    text = re.sub(r"\bMATERIAL\b", "", text)
    text = re.sub(r"\bSTEEL\b", "", text)
    text = re.sub(r"\bALUMINUM\b", "", text)

    text = re.sub(r"[\s,_\-.]+", "", text)

    aliases = {
        re.sub(r"[\s,_\-.]+", "", str(key).upper()): re.sub(
            r"[\s,_\-.]+",
            "",
            str(value).upper(),
        )
        for key, value in rules["material_aliases"].items()
    }

    return aliases.get(text, text)


def split_nest_stock(
    stock_description: str,
    rules: dict[str, Any],
) -> tuple[str, str]:
    raw = str(stock_description or "").strip()

    if not raw:
        return "", ""

    # The normal nesting format is SHAPE_MATERIAL.
    if "_" in raw:
        left, right = raw.split("_", 1)
        if normalize_material(right, rules):
            return left.strip(), right.strip()

    material_pattern = re.compile(
        r"(?i)"
        r"("
        r"ASTM[\s_-]*[A-Z]{0,2}\d{2,4}"
        r"(?:[\s_.-]*(?:GR(?:ADE)?[\s_.-]*[A-Z0-9]+|T\d+))?"
        r"|AH36"
        r"|A36"
        r"|A500(?:[\s_.-]*(?:GR(?:ADE)?[\s_.-]*[A-Z0-9]+))?"
        r"|A572(?:[\s_.-]*(?:GR(?:ADE)?[\s_.-]*[A-Z0-9]+))?"
        r"|A588"
        r"|A992"
        r"|6061(?:[\s_.-]*T6)?"
        r"|5083"
        r"|304L?"
        r"|316L?"
        r")\s*$"
    )

    match = material_pattern.search(raw)
    if match:
        shape = raw[:match.start()].rstrip(" _-|")
        material = match.group(1)
        return shape, material

    return raw, ""


# =============================================================================
# TABLE HEADER DETECTION
# =============================================================================

def find_header_index(
    headers: list[str],
    accepted: set[str],
) -> Optional[int]:
    for index, header in enumerate(headers):
        if header in accepted:
            return index

    return None


def find_prioritized_header_index(
    headers: list[str],
    priority_order: Iterable[str],
) -> Optional[int]:
    """Find the first available quantity header by configured priority."""
    for requested in priority_order:
        normalized_requested = clean_header(requested)
        for index, header in enumerate(headers):
            if header == normalized_requested:
                return index

    return None


def choose_populated_quantity_column(
    rows: list[list[str]],
    header_row: int,
    headers: list[str],
    part_number_index: Optional[int],
    priority_order: Iterable[str],
) -> tuple[Optional[int], str, list[tuple[str, int]]]:
    """
    Choose the quantity column containing the most usable numeric values.

    Parts List exports often include both QTY and TOTAL QUANTITY. Some exports
    leave TOTAL QUANTITY completely blank while QTY contains every value. The
    previous priority-only selection treated every part as missing in that
    situation. Population now wins first; configured priority breaks ties so a
    populated TOTAL QUANTITY still wins over an equally populated QTY column.
    """
    priority = [clean_header(value) for value in priority_order]
    priority_rank = {
        header: rank
        for rank, header in enumerate(priority)
    }

    candidates: list[tuple[int, str]] = []
    seen_indexes: set[int] = set()

    for requested in priority:
        for index, header in enumerate(headers):
            if header == requested and index not in seen_indexes:
                candidates.append((index, header))
                seen_indexes.add(index)

    if not candidates:
        return None, "", []

    counts: dict[int, int] = {index: 0 for index, _ in candidates}

    for row in rows[header_row + 1:]:
        part_raw = get_cell(row, part_number_index)
        if not part_raw:
            continue

        # Ignore repeated table headers inside the export.
        if clean_header(part_raw) in {"PART NUMBER", "PART NO", "PART NUM", "FILE NAME", "FILENAME"}:
            continue

        for index, _ in candidates:
            if parse_number(get_cell(row, index)) is not None:
                counts[index] += 1

    selected_index, selected_header = max(
        candidates,
        key=lambda candidate: (
            counts[candidate[0]],
            -priority_rank.get(candidate[1], len(priority_rank)),
        ),
    )

    population = [
        (header, counts[index])
        for index, header in candidates
    ]

    return selected_index, selected_header, population


def locate_table_header(
    rows: list[list[str]],
    source_name: str,
    aliases: dict[str, set[str]],
    quantity_priority: list[str],
    require_file_name: bool,
) -> tuple[int, dict[str, Optional[int]], dict[str, str]]:
    candidates: list[
        tuple[int, dict[str, Optional[int]], dict[str, str]]
    ] = []

    for row_index, row in enumerate(rows[:100]):
        headers = [clean_header(cell) for cell in row]

        part_aliases = aliases["part_number"]
        if require_file_name:
            part_aliases = {"FILE NAME", "FILENAME"}

        indexes = {
            field_name: find_header_index(headers, accepted)
            for field_name, accepted in aliases.items()
        }

        indexes["part_number"] = find_header_index(
            headers,
            part_aliases,
        )

        # Do not use a set for quantity selection because sets discard the
        # configured preference. TOTAL QUANTITY must be preferred over
        # QUANTITY and QTY.
        indexes["quantity"] = find_prioritized_header_index(
            headers,
            quantity_priority,
        )

        if (
            indexes["part_number"] is not None
            and indexes["quantity"] is not None
        ):
            selected_headers = {
                field_name: (
                    headers[index]
                    if index is not None and index < len(headers)
                    else ""
                )
                for field_name, index in indexes.items()
            }

            candidates.append(
                (row_index, indexes, selected_headers)
            )

    if not candidates:
        preview = "\n".join(
            f"{index + 1}: {row}"
            for index, row in enumerate(rows[:20])
        )

        raise ValueError(
            f"{source_name}: the table header could not be found.\n"
            f"Expected a part/file-name column and quantity column.\n"
            f"First rows:\n{preview}"
        )

    # Parts List exports commonly repeat title, units, and filter header rows.
    return candidates[-1]


def required_comparison_fields(source_name: str) -> list[str]:
    if source_name == "Parts List":
        return [
            "description",
            "shape",
            "length",
            "material",
        ]

    if source_name == "SolidWorks":
        return [
            "shape",
            "length",
            "material",
        ]

    return []


# =============================================================================
# PARTS LIST / SOLIDWORKS PARSING
# =============================================================================

def parse_standard_source(
    path: Path,
    source_name: str,
    require_file_name: bool,
    rules: dict[str, Any],
) -> ParsedSource:
    rows = read_delimited_rows(path)
    aliases = canonical_header_aliases(rules)

    quantity_priority = [
        clean_header(value)
        for value in rules["header_aliases"]["quantity"]
    ]

    header_row, indexes, selected_headers = locate_table_header(
        rows,
        source_name,
        aliases,
        quantity_priority=quantity_priority,
        require_file_name=require_file_name,
    )

    header_cells = [clean_header(cell) for cell in rows[header_row]]
    quantity_index, quantity_header, quantity_population = (
        choose_populated_quantity_column(
            rows,
            header_row,
            header_cells,
            indexes.get("part_number"),
            quantity_priority,
        )
    )

    if quantity_index is not None:
        indexes["quantity"] = quantity_index
        selected_headers["quantity"] = quantity_header

    available_fields = {
        field_name
        for field_name, index in indexes.items()
        if index is not None
    }

    issues: list[Issue] = []

    population_text = ", ".join(
        f'{header}: {count} numeric row(s)'
        for header, count in quantity_population
    )

    issues.append(
        Issue(
            severity="INFO",
            issue="Selected quantity column",
            source=source_name,
            file=str(path),
            details=(
                f'Using "{selected_headers.get("quantity", "")}" '
                f"for quantity values. Available quantity columns: "
                f"{population_text or 'none'}."
            ),
        )
    )

    for required_field in required_comparison_fields(source_name):
        if required_field not in available_fields:
            issues.append(
                Issue(
                    severity="ERROR",
                    issue="Comparison column missing",
                    source=source_name,
                    file=str(path),
                    details=(
                        f'The "{required_field}" column was not found. '
                        "Affected checks will be reported as NOT CHECKED."
                    ),
                )
            )

    parsed_rows: list[SourceRow] = []

    for row_index in range(header_row + 1, len(rows)):
        row = rows[row_index]

        part_raw = get_cell(row, indexes["part_number"])
        part_number = normalize_part_number(
            part_raw,
            source_name,
            rules,
        )

        if not part_number:
            continue

        if clean_header(part_raw) in aliases["part_number"]:
            continue

        quantity_raw = get_cell(row, indexes["quantity"])
        quantity = parse_number(quantity_raw)

        if quantity is None:
            issues.append(
                Issue(
                    severity="ERROR",
                    issue="Quantity could not be parsed",
                    source=source_name,
                    file=str(path),
                    line=str(row_index + 1),
                    part_number=part_number,
                    details=f'Raw quantity: "{quantity_raw}"',
                )
            )
            continue

        description_raw = get_cell(row, indexes.get("description"))

        length_raw = get_cell(row, indexes.get("length"))
        length = parse_number(length_raw)
        if length_raw and length is None:
            issues.append(
                Issue(
                    severity="ERROR",
                    issue="Length could not be parsed",
                    source=source_name,
                    file=str(path),
                    line=str(row_index + 1),
                    part_number=part_number,
                    details=f'Raw length: "{length_raw}"',
                )
            )

        width_raw = get_cell(row, indexes.get("width"))
        width = parse_number(width_raw)
        if width_raw and width is None:
            issues.append(
                Issue(
                    severity="ERROR",
                    issue="Width could not be parsed",
                    source=source_name,
                    file=str(path),
                    line=str(row_index + 1),
                    part_number=part_number,
                    details=f'Raw width: "{width_raw}"',
                )
            )

        shape_raw = get_cell(row, indexes.get("shape"))
        material_raw = get_cell(row, indexes.get("material"))
        configuration_raw = get_cell(row, indexes.get("configuration"))

        parsed_rows.append(
            SourceRow(
                source=source_name,
                source_file=str(path),
                source_line=row_index + 1,
                part_number_raw=part_raw,
                part_number=part_number,
                quantity_raw=quantity_raw,
                quantity=quantity,
                description_raw=description_raw,
                length_raw=length_raw,
                length=length,
                width_raw=width_raw,
                width=width,
                shape_raw=shape_raw,
                shape_key=normalize_shape(shape_raw, rules),
                material_raw=material_raw,
                material_key=normalize_material(material_raw, rules),
                configuration_raw=configuration_raw,
            )
        )

    if not parsed_rows:
        issues.append(
            Issue(
                severity="ERROR",
                issue="No usable part rows found",
                source=source_name,
                file=str(path),
                details="Check the selected file and its column headers.",
            )
        )

    return ParsedSource(
        name=source_name,
        path=path,
        rows=parsed_rows,
        available_fields=available_fields,
        selected_headers=selected_headers,
        issues=issues,
    )


# =============================================================================
# NEST PARSING
# =============================================================================

def row_headers(row: list[str]) -> list[str]:
    return [clean_header(cell) for cell in row]


def is_parts_header(row: list[str]) -> bool:
    headers = row_headers(row)

    return (
        len(headers) >= 3
        and headers[0] == "PARTS"
        and headers[1] == "LENGTH"
        and headers[2] in {"QUANTITY", "QTY"}
    )


def is_bar_header(row: list[str]) -> bool:
    headers = row_headers(row)

    return (
        len(headers) >= 3
        and headers[0] == "BAR"
        and headers[1] == "LENGTH"
        and headers[2] in {"QUANTITY", "QTY"}
    )


def is_end_of_parts_section(row: list[str]) -> bool:
    first = clean_header(row[0]) if row else ""

    return first in {
        "CUT THICKNESS",
        "BAR ENDS TRIM",
        "BAR END TRIM",
        "REMNANTS",
        "SCRAP",
        "SUMMARY",
    }


def previous_nonempty_first_cell(
    rows: list[list[str]],
    start_index: int,
) -> str:
    for index in range(start_index - 1, -1, -1):
        nonempty = [cell.strip() for cell in rows[index] if cell.strip()]
        if nonempty:
            return nonempty[0]

    return ""


def parse_nest_file(
    path: Path,
    rules: dict[str, Any],
) -> ParsedSource:
    rows = read_delimited_rows(path)
    issues: list[Issue] = []

    parts_header_index = next(
        (
            index
            for index, row in enumerate(rows)
            if is_parts_header(row)
        ),
        None,
    )

    if parts_header_index is None:
        issues.append(
            Issue(
                severity="WARNING",
                issue="No Parts;Length;Quantity section found",
                source="Nest",
                file=str(path),
                details="The file was skipped.",
            )
        )
        return ParsedSource(
            name="Nest",
            path=path,
            rows=[],
            available_fields=set(),
            selected_headers={
                "part_number": "PARTS",
                "length": "LENGTH",
                "quantity": "QUANTITY",
            },
            issues=issues,
        )

    bar_header_index = next(
        (
            index
            for index, row in enumerate(rows[:parts_header_index])
            if is_bar_header(row)
        ),
        None,
    )

    stock_description = ""
    if bar_header_index is not None:
        stock_description = previous_nonempty_first_cell(
            rows,
            bar_header_index,
        )

    nest_shape_raw, nest_material_raw = split_nest_stock(
        stock_description,
        rules,
    )

    if stock_description and not nest_material_raw:
        issues.append(
            Issue(
                severity="ERROR",
                issue="Nest material could not be extracted",
                source="Nest",
                file=str(path),
                details=(
                    f'Stock description: "{stock_description}". '
                    "Update material aliases/patterns if this is a valid format."
                ),
            )
        )

    parsed_rows: list[SourceRow] = []

    for row_index in range(parts_header_index + 1, len(rows)):
        row = rows[row_index]

        if not any(cell.strip() for cell in row):
            continue

        if is_end_of_parts_section(row):
            break

        part_raw = get_cell(row, 0)
        part_number = normalize_part_number(part_raw, "Nest", rules)

        if not part_number:
            continue

        length_raw = get_cell(row, 1)
        quantity_raw = get_cell(row, 2)

        quantity = parse_number(quantity_raw)
        length = parse_number(length_raw)

        if quantity is None:
            issues.append(
                Issue(
                    severity="ERROR",
                    issue="Nest quantity could not be parsed",
                    source="Nest",
                    file=str(path),
                    line=str(row_index + 1),
                    part_number=part_number,
                    details=f'Raw row: {";".join(row)}',
                )
            )
            continue

        if length is None:
            issues.append(
                Issue(
                    severity="ERROR",
                    issue="Nest length could not be parsed",
                    source="Nest",
                    file=str(path),
                    line=str(row_index + 1),
                    part_number=part_number,
                    details=f'Raw length: "{length_raw}"',
                )
            )

        parsed_rows.append(
            SourceRow(
                source="Nest",
                source_file=str(path),
                source_line=row_index + 1,
                part_number_raw=part_raw,
                part_number=part_number,
                quantity_raw=quantity_raw,
                quantity=quantity,
                length_raw=length_raw,
                length=length,
                shape_raw=nest_shape_raw,
                shape_key=normalize_shape(nest_shape_raw, rules),
                material_raw=nest_material_raw,
                material_key=normalize_material(nest_material_raw, rules),
                stock_description_raw=stock_description,
            )
        )

    return ParsedSource(
        name="Nest",
        path=path,
        rows=parsed_rows,
        available_fields={
            "part_number",
            "quantity",
            "length",
            "shape",
            "material",
        },
        selected_headers={
            "part_number": "PARTS",
            "length": "LENGTH",
            "quantity": "QUANTITY",
        },
        issues=issues,
    )


def parse_nest_folder(
    folder: Path,
    rules: dict[str, Any],
) -> tuple[list[SourceRow], list[Issue], int]:
    files = sorted(
        set(folder.rglob("*.csv")) | set(folder.rglob("*.txt"))
    )

    if not files:
        raise FileNotFoundError(
            f"No CSV or TXT nesting files were found in {folder}"
        )

    all_rows: list[SourceRow] = []
    all_issues: list[Issue] = []
    usable_file_count = 0

    for file_path in files:
        parsed = parse_nest_file(file_path, rules)
        all_rows.extend(parsed.rows)
        all_issues.extend(parsed.issues)

        if parsed.rows:
            usable_file_count += 1

    return all_rows, all_issues, usable_file_count


# =============================================================================
# AGGREGATION AND SOURCE VALIDATION
# =============================================================================

def distinct_numeric_values(
    values: list[float],
    tolerance: float,
) -> list[float]:
    distinct: list[float] = []

    for value in values:
        if not any(abs(value - existing) <= tolerance for existing in distinct):
            distinct.append(value)

    return distinct


def aggregate_rows(
    rows: list[SourceRow],
    rules: dict[str, Any],
    quantity_mode: str = "sum",
) -> tuple[dict[str, Aggregate], list[Issue]]:
    aggregates: dict[str, Aggregate] = {}
    issues: list[Issue] = []

    for row in rows:
        aggregate = aggregates.setdefault(
            row.part_number,
            Aggregate(
                source=row.source,
                part_number=row.part_number,
            ),
        )

        aggregate.rows.append(row)
        aggregate.quantity += float(row.quantity or 0.0)

        if row.description_raw:
            aggregate.descriptions.append(row.description_raw)

        if row.length is not None:
            aggregate.lengths.append(row.length)
        if row.length_raw:
            aggregate.length_raw_values.append(row.length_raw)

        if row.width is not None:
            aggregate.widths.append(row.width)
        if row.width_raw:
            aggregate.width_raw_values.append(row.width_raw)

        if row.shape_raw:
            aggregate.shape_raw_values.append(row.shape_raw)
        if row.shape_key:
            aggregate.shape_keys.append(row.shape_key)

        if row.material_raw:
            aggregate.material_raw_values.append(row.material_raw)
        if row.material_key:
            aggregate.material_keys.append(row.material_key)

        aggregate.source_locations.append(
            f"{Path(row.source_file).name}:{row.source_line}"
        )

        if row.stock_description_raw:
            aggregate.stock_descriptions.append(row.stock_description_raw)

    if quantity_mode not in {"sum", "preaggregated"}:
        raise ValueError(
            f'Unsupported quantity aggregation mode: "{quantity_mode}"'
        )

    length_tolerance = float(rules["length_tolerance_inches"])
    width_tolerance = float(rules["width_tolerance_inches"])
    quantity_tolerance = float(rules["quantity_tolerance"])

    for part_number, aggregate in aggregates.items():
        if quantity_mode == "preaggregated":
            quantity_values = [
                float(row.quantity)
                for row in aggregate.rows
                if row.quantity is not None
            ]

            if quantity_values:
                # TOTAL QUANTITY is already the full part total. Repeated rows
                # must not be added together.
                aggregate.quantity = quantity_values[0]

                distinct_quantities = distinct_numeric_values(
                    quantity_values,
                    quantity_tolerance,
                )

                if len(distinct_quantities) > 1:
                    issues.append(
                        Issue(
                            severity="ERROR",
                            issue=(
                                "Duplicate rows have conflicting "
                                "TOTAL QUANTITY values"
                            ),
                            source=aggregate.source,
                            part_number=part_number,
                            details=(
                                "Values: "
                                + ", ".join(
                                    format_quantity(value)
                                    for value in distinct_quantities
                                )
                                + f". Rows: {aggregate.locations_text}"
                            ),
                        )
                    )

        distinct_lengths = distinct_numeric_values(
            aggregate.lengths,
            length_tolerance,
        )
        if len(distinct_lengths) > 1:
            issues.append(
                Issue(
                    severity="ERROR",
                    issue="One part number has multiple lengths",
                    source=aggregate.source,
                    part_number=part_number,
                    details=(
                        f"Lengths: {', '.join(format_decimal(v) for v in distinct_lengths)}. "
                        f"Rows: {aggregate.locations_text}"
                    ),
                )
            )

        distinct_widths = distinct_numeric_values(
            aggregate.widths,
            width_tolerance,
        )
        if len(distinct_widths) > 1:
            issues.append(
                Issue(
                    severity="ERROR",
                    issue="One part number has multiple widths",
                    source=aggregate.source,
                    part_number=part_number,
                    details=(
                        f"Widths: {', '.join(format_decimal(v) for v in distinct_widths)}. "
                        f"Rows: {aggregate.locations_text}"
                    ),
                )
            )

        distinct_shapes = unique_preserve_order(aggregate.shape_keys)
        if len(distinct_shapes) > 1:
            issues.append(
                Issue(
                    severity="ERROR",
                    issue="One part number has multiple shapes",
                    source=aggregate.source,
                    part_number=part_number,
                    details=(
                        f"Shape keys: {', '.join(distinct_shapes)}. "
                        f"Raw shapes: {aggregate.raw_shapes}. "
                        f"Rows: {aggregate.locations_text}"
                    ),
                )
            )

        distinct_materials = unique_preserve_order(aggregate.material_keys)
        if len(distinct_materials) > 1:
            issues.append(
                Issue(
                    severity="ERROR",
                    issue="One part number has multiple materials",
                    source=aggregate.source,
                    part_number=part_number,
                    details=(
                        f"Material keys: {', '.join(distinct_materials)}. "
                        f"Raw materials: {aggregate.raw_materials}. "
                        f"Rows: {aggregate.locations_text}"
                    ),
                )
            )

        if aggregate.source in {"Parts List", "SolidWorks"} and len(aggregate.rows) > 1:
            if quantity_mode == "preaggregated":
                duplicate_details = (
                    f"{len(aggregate.rows)} rows were found. "
                    "The repeated TOTAL QUANTITY was used once and was not summed. "
                    f"Rows: {aggregate.locations_text}"
                )
            else:
                duplicate_details = (
                    f"{len(aggregate.rows)} rows were summed because the "
                    "selected column is a per-row QUANTITY field. "
                    f"Rows: {aggregate.locations_text}"
                )

            issues.append(
                Issue(
                    severity="WARNING",
                    issue="Duplicate part rows were aggregated",
                    source=aggregate.source,
                    part_number=part_number,
                    details=duplicate_details,
                )
            )

    return aggregates, issues


# =============================================================================
# CLASSIFICATION
# =============================================================================

def contains_keyword(
    description: str,
    keywords: Iterable[str],
) -> bool:
    normalized = str(description or "").upper()

    for keyword in keywords:
        keyword_text = str(keyword).upper().strip()

        if not keyword_text:
            continue

        if re.search(
            rf"(?<![A-Z0-9]){re.escape(keyword_text)}(?![A-Z0-9])",
            normalized,
        ):
            return True

    return False


def classify_part(
    parts_list: Optional[Aggregate],
    solidworks: Optional[Aggregate],
    rules: dict[str, Any],
) -> tuple[str, str]:
    parts_description = (
        parts_list.descriptions_text
        if parts_list is not None
        else ""
    )
    solidworks_description = (
        solidworks.descriptions_text
        if solidworks is not None
        else ""
    )

    description = parts_description or solidworks_description

    if not description:
        return "UNKNOWN", ""

    if contains_keyword(
        description,
        rules["plate_description_keywords"],
    ):
        return "PLATE", description

    if contains_keyword(
        description,
        rules["excluded_hardware_description_keywords"],
    ):
        return "EXCLUDED HARDWARE", description

    return "LINEAR", description


# =============================================================================
# FIELD COMPARISON
# =============================================================================

def compare_numeric_field(
    baseline_value: Optional[float],
    comparison_value: Optional[float],
    tolerance: float,
    baseline_label: str,
    comparison_label: str,
) -> tuple[str, Optional[str], Optional[str]]:
    if baseline_value is None and comparison_value is None:
        return (
            "NOT CHECKED",
            None,
            f"{baseline_label} and {comparison_label} values are blank",
        )

    if baseline_value is None:
        return (
            "NOT CHECKED",
            None,
            f"{baseline_label} value is blank",
        )

    if comparison_value is None:
        return (
            "NOT CHECKED",
            None,
            f"{comparison_label} value is blank",
        )

    if abs(comparison_value - baseline_value) <= tolerance:
        return "MATCH", None, None

    return (
        "MISMATCH",
        (
            f"{comparison_label} differs from {baseline_label}: "
            f"{format_decimal(comparison_value)} versus "
            f"{format_decimal(baseline_value)}"
        ),
        None,
    )


def compare_text_field(
    baseline_key: str,
    comparison_key: str,
    baseline_raw: str,
    comparison_raw: str,
    baseline_label: str,
    comparison_label: str,
) -> tuple[str, Optional[str], Optional[str]]:
    if not baseline_key and not comparison_key:
        return (
            "NOT CHECKED",
            None,
            f"{baseline_label} and {comparison_label} values are blank",
        )

    if not baseline_key:
        return (
            "NOT CHECKED",
            None,
            f"{baseline_label} value is blank",
        )

    if not comparison_key:
        return (
            "NOT CHECKED",
            None,
            f"{comparison_label} value is blank",
        )

    if baseline_key == comparison_key:
        return "MATCH", None, None

    return (
        "MISMATCH",
        (
            f"{comparison_label} differs from {baseline_label}: "
            f'"{comparison_raw}" versus "{baseline_raw}"'
        ),
        None,
    )


def add_check_result(
    row: dict[str, str],
    column_name: str,
    result: tuple[str, Optional[str], Optional[str]],
    errors: list[str],
    incomplete: list[str],
) -> None:
    status, error_message, incomplete_message = result
    row[column_name] = status

    if error_message:
        errors.append(error_message)

    if incomplete_message:
        incomplete.append(incomplete_message)


# =============================================================================
# FULL RECONCILIATION
# =============================================================================

def reconcile(
    nest: dict[str, Aggregate],
    parts: dict[str, Aggregate],
    solidworks: dict[str, Aggregate],
    source_fields: dict[str, set[str]],
    rules: dict[str, Any],
    issues: list[Issue],
) -> list[dict[str, str]]:
    all_part_numbers = sorted(set(nest) | set(parts) | set(solidworks))
    part_issue_severity: dict[str, set[str]] = defaultdict(set)

    for issue in issues:
        if issue.part_number:
            part_issue_severity[issue.part_number].add(issue.severity)

    rows: list[dict[str, str]] = []

    quantity_tolerance = float(rules["quantity_tolerance"])
    length_tolerance = float(rules["length_tolerance_inches"])
    width_tolerance = float(rules["width_tolerance_inches"])
    checks = rules["checks"]

    for part_number in all_part_numbers:
        nest_part = nest.get(part_number)
        parts_part = parts.get(part_number)
        sw_part = solidworks.get(part_number)

        category, description = classify_part(
            parts_part,
            sw_part,
            rules,
        )

        errors: list[str] = []
        incomplete: list[str] = []

        row: dict[str, str] = {
            "Review Status": "",
            "Category": category,
            "Part Number": part_number,
            "Description": description,
            "Nest Expected": "YES" if category == "LINEAR" else "NO",

            "SolidWorks Quantity": format_quantity(
                sw_part.quantity if sw_part else None
            ),
            "Parts List Quantity": format_quantity(
                parts_part.quantity if parts_part else None
            ),
            "Parts List Quantity Delta": difference(
                parts_part.quantity if parts_part else None,
                sw_part.quantity if sw_part else None,
            ),
            "Parts List Quantity Check": "",

            "Nest Quantity": format_quantity(
                nest_part.quantity if nest_part else None
            ),
            "Nest Quantity Delta": difference(
                nest_part.quantity if nest_part else None,
                sw_part.quantity if sw_part else None,
            ),
            "Nest Quantity Check": "",

            "SolidWorks Length Raw": (
                sw_part.raw_lengths if sw_part else ""
            ),
            "SolidWorks Length Decimal": format_decimal(
                sw_part.representative_length if sw_part else None
            ),
            "Parts List Length Raw": (
                parts_part.raw_lengths if parts_part else ""
            ),
            "Parts List Length Decimal": format_decimal(
                parts_part.representative_length if parts_part else None
            ),
            "Parts List Length Check": "",
            "Nest Length Raw": (
                nest_part.raw_lengths if nest_part else ""
            ),
            "Nest Length Decimal": format_decimal(
                nest_part.representative_length if nest_part else None
            ),
            "Nest Length Fraction": format_fraction(
                nest_part.representative_length if nest_part else None
            ),
            "Nest Length Check": "",

            "SolidWorks Width Raw": (
                sw_part.raw_widths if sw_part else ""
            ),
            "SolidWorks Width Decimal": format_decimal(
                sw_part.representative_width if sw_part else None
            ),
            "Parts List Width Raw": (
                parts_part.raw_widths if parts_part else ""
            ),
            "Parts List Width Decimal": format_decimal(
                parts_part.representative_width if parts_part else None
            ),
            "Parts List Width Check": "",

            "SolidWorks Shape": (
                sw_part.raw_shapes if sw_part else ""
            ),
            "SolidWorks Shape Key": (
                sw_part.representative_shape_key if sw_part else ""
            ),
            "Parts List Shape": (
                parts_part.raw_shapes if parts_part else ""
            ),
            "Parts List Shape Key": (
                parts_part.representative_shape_key if parts_part else ""
            ),
            "Parts List Shape Check": "",
            "Nest Shape": (
                nest_part.raw_shapes if nest_part else ""
            ),
            "Nest Shape Key": (
                nest_part.representative_shape_key if nest_part else ""
            ),
            "Nest Shape Check": "",

            "SolidWorks Material": (
                sw_part.raw_materials if sw_part else ""
            ),
            "SolidWorks Material Key": (
                sw_part.representative_material_key if sw_part else ""
            ),
            "Parts List Material": (
                parts_part.raw_materials if parts_part else ""
            ),
            "Parts List Material Key": (
                parts_part.representative_material_key if parts_part else ""
            ),
            "Parts List Material Check": "",
            "Nest Material": (
                nest_part.raw_materials if nest_part else ""
            ),
            "Nest Material Key": (
                nest_part.representative_material_key if nest_part else ""
            ),
            "Nest Material Check": "",

            "SolidWorks Source Rows": (
                sw_part.locations_text if sw_part else ""
            ),
            "Parts List Source Rows": (
                parts_part.locations_text if parts_part else ""
            ),
            "Nest Source Rows": (
                nest_part.locations_text if nest_part else ""
            ),
            "Nest Stock Description": (
                nest_part.stock_text if nest_part else ""
            ),

            "Errors": "",
            "Not Checked Reasons": "",
            "Source Issue Flag": (
                ", ".join(sorted(part_issue_severity.get(part_number, set())))
            ),
        }

        missing_core_sources: list[str] = []

        if sw_part is None:
            missing_core_sources.append("SolidWorks")

        if parts_part is None:
            missing_core_sources.append("Parts List")

        if missing_core_sources:
            row["Review Status"] = "MISSING FROM SOLIDWORKS OR PARTS LIST"
            errors.append(
                "Missing from: " + ", ".join(missing_core_sources)
            )

            if category == "LINEAR" and nest_part is None:
                errors.append("Also missing from linear nesting")

            row["Errors"] = "; ".join(errors)
            row["Not Checked Reasons"] = "; ".join(incomplete)
            rows.append(row)
            continue

        # Parts List versus SolidWorks.
        if checks.get("quantity", True):
            add_check_result(
                row,
                "Parts List Quantity Check",
                compare_numeric_field(
                    sw_part.quantity,
                    parts_part.quantity,
                    quantity_tolerance,
                    "SolidWorks quantity",
                    "Parts List quantity",
                ),
                errors,
                incomplete,
            )
        else:
            row["Parts List Quantity Check"] = "DISABLED"

        if checks.get("length", True):
            add_check_result(
                row,
                "Parts List Length Check",
                compare_numeric_field(
                    sw_part.representative_length,
                    parts_part.representative_length,
                    length_tolerance,
                    "SolidWorks length",
                    "Parts List length",
                ),
                errors,
                incomplete,
            )
        else:
            row["Parts List Length Check"] = "DISABLED"

        if checks.get("shape", True):
            add_check_result(
                row,
                "Parts List Shape Check",
                compare_text_field(
                    sw_part.representative_shape_key,
                    parts_part.representative_shape_key,
                    sw_part.raw_shapes,
                    parts_part.raw_shapes,
                    "SolidWorks shape",
                    "Parts List shape",
                ),
                errors,
                incomplete,
            )
        else:
            row["Parts List Shape Check"] = "DISABLED"

        if checks.get("material", True):
            add_check_result(
                row,
                "Parts List Material Check",
                compare_text_field(
                    sw_part.representative_material_key,
                    parts_part.representative_material_key,
                    sw_part.raw_materials,
                    parts_part.raw_materials,
                    "SolidWorks material",
                    "Parts List material",
                ),
                errors,
                incomplete,
            )
        else:
            row["Parts List Material Check"] = "DISABLED"

        if checks.get("width_when_available", True):
            sw_has_width_column = "width" in source_fields["SolidWorks"]
            parts_has_width_column = "width" in source_fields["Parts List"]
            sw_width = sw_part.representative_width
            parts_width = parts_part.representative_width

            # Width is normally meaningful for plates and some rectangular
            # profiles. Blank width cells on ordinary beams/tubes should not
            # prevent an otherwise exact match.
            if sw_width is not None or parts_width is not None:
                add_check_result(
                    row,
                    "Parts List Width Check",
                    compare_numeric_field(
                        sw_width,
                        parts_width,
                        width_tolerance,
                        "SolidWorks width",
                        "Parts List width",
                    ),
                    errors,
                    incomplete,
                )
            elif sw_has_width_column or parts_has_width_column:
                row["Parts List Width Check"] = "NOT APPLICABLE"
            else:
                row["Parts List Width Check"] = "NOT AVAILABLE"
        else:
            row["Parts List Width Check"] = "DISABLED"

        # Category-specific nesting rules.
        if category == "LINEAR":
            if nest_part is None:
                errors.append("Part is missing from linear nesting")
                row["Nest Quantity Check"] = "MISSING"
                row["Nest Length Check"] = "MISSING"
                row["Nest Shape Check"] = "MISSING"
                row["Nest Material Check"] = "MISSING"
            else:
                if checks.get("quantity", True):
                    add_check_result(
                        row,
                        "Nest Quantity Check",
                        compare_numeric_field(
                            sw_part.quantity,
                            nest_part.quantity,
                            quantity_tolerance,
                            "SolidWorks quantity",
                            "Nest quantity",
                        ),
                        errors,
                        incomplete,
                    )
                else:
                    row["Nest Quantity Check"] = "DISABLED"

                if checks.get("length", True):
                    add_check_result(
                        row,
                        "Nest Length Check",
                        compare_numeric_field(
                            sw_part.representative_length,
                            nest_part.representative_length,
                            length_tolerance,
                            "SolidWorks length",
                            "Nest length",
                        ),
                        errors,
                        incomplete,
                    )
                else:
                    row["Nest Length Check"] = "DISABLED"

                if checks.get("shape", True):
                    add_check_result(
                        row,
                        "Nest Shape Check",
                        compare_text_field(
                            sw_part.representative_shape_key,
                            nest_part.representative_shape_key,
                            sw_part.raw_shapes,
                            nest_part.raw_shapes,
                            "SolidWorks shape",
                            "Nest shape",
                        ),
                        errors,
                        incomplete,
                    )
                else:
                    row["Nest Shape Check"] = "DISABLED"

                if checks.get("material", True):
                    add_check_result(
                        row,
                        "Nest Material Check",
                        compare_text_field(
                            sw_part.representative_material_key,
                            nest_part.representative_material_key,
                            sw_part.raw_materials,
                            nest_part.raw_materials,
                            "SolidWorks material",
                            "Nest material",
                        ),
                        errors,
                        incomplete,
                    )
                else:
                    row["Nest Material Check"] = "DISABLED"

        elif category == "PLATE":
            row["Nest Quantity Check"] = "NOT REQUIRED"
            row["Nest Length Check"] = "NOT REQUIRED"
            row["Nest Shape Check"] = "NOT REQUIRED"
            row["Nest Material Check"] = "NOT REQUIRED"

            incomplete.append(
                "Plate nesting is intentionally not checked yet"
            )

            if nest_part is not None:
                errors.append(
                    "Plate was unexpectedly found in the linear-nesting folder"
                )

        elif category == "EXCLUDED HARDWARE":
            row["Nest Quantity Check"] = "NOT REQUIRED"
            row["Nest Length Check"] = "NOT REQUIRED"
            row["Nest Shape Check"] = "NOT REQUIRED"
            row["Nest Material Check"] = "NOT REQUIRED"

            incomplete.append(
                "Hardware is excluded from linear-material nesting"
            )

            if nest_part is not None:
                errors.append(
                    "Excluded hardware was unexpectedly found in linear nesting"
                )

        else:
            row["Nest Quantity Check"] = "NOT CHECKED"
            row["Nest Length Check"] = "NOT CHECKED"
            row["Nest Shape Check"] = "NOT CHECKED"
            row["Nest Material Check"] = "NOT CHECKED"

            incomplete.append(
                "Part category is unknown because description data is unavailable"
            )

        if "ERROR" in part_issue_severity.get(part_number, set()):
            errors.append("Part has one or more source-data errors")

        if errors:
            row["Review Status"] = "ERROR REQUIRING ACTION"
        elif incomplete:
            row["Review Status"] = "NOT CHECKED"
        else:
            row["Review Status"] = "EXACT MATCH"

        row["Errors"] = "; ".join(unique_preserve_order(errors))
        row["Not Checked Reasons"] = "; ".join(
            unique_preserve_order(incomplete)
        )

        rows.append(row)

    status_order = {
        "ERROR REQUIRING ACTION": 0,
        "MISSING FROM SOLIDWORKS OR PARTS LIST": 1,
        "NOT CHECKED": 2,
        "EXACT MATCH": 3,
    }

    rows.sort(
        key=lambda row: (
            status_order.get(row["Review Status"], 9),
            row["Category"],
            row["Part Number"],
        )
    )

    return rows


# =============================================================================
# OUTPUT
# =============================================================================

REPORT_HEADERS = [
    "Part Number",
    "Review Status",
    "Category",
    "Description",
    "Nest Expected",

    "SolidWorks Quantity",
    "Parts List Quantity",
    "Parts List Quantity Delta",
    "Parts List Quantity Check",
    "Nest Quantity",
    "Nest Quantity Delta",
    "Nest Quantity Check",

    "SolidWorks Length Raw",
    "SolidWorks Length Decimal",
    "Parts List Length Raw",
    "Parts List Length Decimal",
    "Parts List Length Check",
    "Nest Length Raw",
    "Nest Length Decimal",
    "Nest Length Fraction",
    "Nest Length Check",

    "SolidWorks Width Raw",
    "SolidWorks Width Decimal",
    "Parts List Width Raw",
    "Parts List Width Decimal",
    "Parts List Width Check",

    "SolidWorks Shape",
    "SolidWorks Shape Key",
    "Parts List Shape",
    "Parts List Shape Key",
    "Parts List Shape Check",
    "Nest Shape",
    "Nest Shape Key",
    "Nest Shape Check",

    "SolidWorks Material",
    "SolidWorks Material Key",
    "Parts List Material",
    "Parts List Material Key",
    "Parts List Material Check",
    "Nest Material",
    "Nest Material Key",
    "Nest Material Check",

    "Errors",
    "Not Checked Reasons",
    "Source Issue Flag",

    "SolidWorks Source Rows",
    "Parts List Source Rows",
    "Nest Source Rows",
    "Nest Stock Description",
]

COMPACT_REPORT_HEADERS = [
    "Part Number",
    "Review Status",
    "Category",
    "Description",
    "Problem",
    "SolidWorks Value",
    "Parts List Value",
    "Nest Value",
    "Required Action",
]

SOURCE_ROW_HEADERS = [
    "Part Number Normalized",
    "Part Number Raw",
    "Source",
    "Source File",
    "Source Line",
    "Quantity Raw",
    "Quantity Parsed",
    "Description",
    "Length Raw",
    "Length Decimal",
    "Width Raw",
    "Width Decimal",
    "Shape Raw",
    "Shape Key",
    "Material Raw",
    "Material Key",
    "Configuration",
    "Nest Stock Description",
]

ISSUE_HEADERS = [
    "Part Number",
    "Severity",
    "Issue",
    "Source",
    "File",
    "Line",
    "Details",
]

ACTION_STATUSES = {
    "ERROR REQUIRING ACTION",
    "MISSING FROM SOLIDWORKS OR PARTS LIST",
}


def write_dict_csv(
    path: Path,
    rows: list[dict[str, str]],
    headers: list[str],
) -> None:
    with path.open(
        "w",
        newline="",
        encoding="utf-8-sig",
    ) as file:
        writer = csv.DictWriter(
            file,
            fieldnames=headers,
            extrasaction="ignore",
        )
        writer.writeheader()
        writer.writerows(rows)


def source_row_to_dict(row: SourceRow) -> dict[str, str]:
    return {
        "Part Number Normalized": row.part_number,
        "Part Number Raw": row.part_number_raw,
        "Source": row.source,
        "Source File": row.source_file,
        "Source Line": str(row.source_line),
        "Quantity Raw": row.quantity_raw,
        "Quantity Parsed": format_quantity(row.quantity),
        "Description": row.description_raw,
        "Length Raw": row.length_raw,
        "Length Decimal": format_decimal(row.length),
        "Width Raw": row.width_raw,
        "Width Decimal": format_decimal(row.width),
        "Shape Raw": row.shape_raw,
        "Shape Key": row.shape_key,
        "Material Raw": row.material_raw,
        "Material Key": row.material_key,
        "Configuration": row.configuration_raw,
        "Nest Stock Description": row.stock_description_raw,
    }


def compact_value(
    row: dict[str, str],
    source: str,
    field_name: str,
) -> str:
    candidates: dict[tuple[str, str], list[str]] = {
        ("SolidWorks", "Quantity"): ["SolidWorks Quantity"],
        ("Parts List", "Quantity"): ["Parts List Quantity"],
        ("Nest", "Quantity"): ["Nest Quantity"],
        ("SolidWorks", "Length"): [
            "SolidWorks Length Raw",
            "SolidWorks Length Decimal",
        ],
        ("Parts List", "Length"): [
            "Parts List Length Raw",
            "Parts List Length Decimal",
        ],
        ("Nest", "Length"): [
            "Nest Length Raw",
            "Nest Length Fraction",
            "Nest Length Decimal",
        ],
        ("SolidWorks", "Width"): [
            "SolidWorks Width Raw",
            "SolidWorks Width Decimal",
        ],
        ("Parts List", "Width"): [
            "Parts List Width Raw",
            "Parts List Width Decimal",
        ],
        ("SolidWorks", "Shape"): ["SolidWorks Shape"],
        ("Parts List", "Shape"): ["Parts List Shape"],
        ("Nest", "Shape"): ["Nest Shape"],
        ("SolidWorks", "Material"): ["SolidWorks Material"],
        ("Parts List", "Material"): ["Parts List Material"],
        ("Nest", "Material"): ["Nest Material"],
    }

    for column in candidates.get((source, field_name), []):
        value = str(row.get(column, "")).strip()
        if value:
            return value

    return ""


def add_compact_line(
    target: list[str],
    label: str,
    value: str,
) -> None:
    if value:
        target.append(f"{label}: {value}")


def build_compact_rows(
    rows: list[dict[str, str]],
    issues: list[Issue],
) -> list[dict[str, str]]:
    issues_by_part: dict[str, list[str]] = defaultdict(list)

    for issue in issues:
        if not issue.part_number:
            continue

        detail = f"{issue.source}: {issue.issue}"
        if issue.details:
            detail += f" — {issue.details}"
        issues_by_part[issue.part_number].append(detail)

    compact_rows: list[dict[str, str]] = []

    field_specs = [
        (
            "Quantity",
            "Parts List Quantity Check",
            "Nest Quantity Check",
        ),
        (
            "Length",
            "Parts List Length Check",
            "Nest Length Check",
        ),
        (
            "Width",
            "Parts List Width Check",
            "",
        ),
        (
            "Shape",
            "Parts List Shape Check",
            "Nest Shape Check",
        ),
        (
            "Material",
            "Parts List Material Check",
            "Nest Material Check",
        ),
    ]

    for row in rows:
        part_number = row["Part Number"]
        status = row["Review Status"]

        problems: list[str] = []
        solidworks_values: list[str] = []
        parts_values: list[str] = []
        nest_values: list[str] = []
        actions: list[str] = []

        error_text = str(row.get("Errors", ""))
        not_checked_text = str(row.get("Not Checked Reasons", ""))

        missing_core_match = re.search(
            r"Missing from:\s*([^;]+)",
            error_text,
            flags=re.IGNORECASE,
        )
        if missing_core_match:
            missing_sources = missing_core_match.group(1).strip()
            problems.append(f"Missing from {missing_sources}")
            actions.append(
                f"Add or correct the part in {missing_sources}, then rerun."
            )

            if row.get("SolidWorks Source Rows"):
                add_compact_line(
                    solidworks_values,
                    "Present quantity",
                    compact_value(row, "SolidWorks", "Quantity"),
                )
            if row.get("Parts List Source Rows"):
                add_compact_line(
                    parts_values,
                    "Present quantity",
                    compact_value(row, "Parts List", "Quantity"),
                )
            if row.get("Nest Source Rows"):
                add_compact_line(
                    nest_values,
                    "Present quantity",
                    compact_value(row, "Nest", "Quantity"),
                )

        nest_missing = (
            not row.get("Nest Source Rows")
            and any(
                row.get(column) == "MISSING"
                for column in (
                    "Nest Quantity Check",
                    "Nest Length Check",
                    "Nest Shape Check",
                    "Nest Material Check",
                )
            )
        )
        if nest_missing:
            problems.append("Missing from linear nesting")
            nest_values.append("Part not found")
            actions.append("Create or update the nesting entry, then rerun.")

        for field_name, parts_check_column, nest_check_column in field_specs:
            parts_status = row.get(parts_check_column, "")
            nest_status = (
                row.get(nest_check_column, "")
                if nest_check_column
                else ""
            )

            mismatched_sources: list[str] = []

            if parts_status == "MISMATCH":
                mismatched_sources.append("Parts List")
                add_compact_line(
                    parts_values,
                    field_name,
                    compact_value(row, "Parts List", field_name),
                )

            if nest_status == "MISMATCH":
                mismatched_sources.append("Nest")
                add_compact_line(
                    nest_values,
                    field_name,
                    compact_value(row, "Nest", field_name),
                )

            if mismatched_sources:
                problems.append(
                    f"{field_name} differs in "
                    + " and ".join(mismatched_sources)
                )
                add_compact_line(
                    solidworks_values,
                    field_name,
                    compact_value(row, "SolidWorks", field_name),
                )

                if mismatched_sources == ["Parts List"]:
                    actions.append(
                        f"Correct the Parts List {field_name.lower()} "
                        "or confirm the SolidWorks baseline."
                    )
                elif mismatched_sources == ["Nest"]:
                    actions.append(
                        f"Update the nest {field_name.lower()} "
                        "to match SolidWorks."
                    )
                else:
                    actions.append(
                        f"Correct the Parts List and nest {field_name.lower()} "
                        "or confirm the SolidWorks baseline."
                    )

        for issue_detail in issues_by_part.get(part_number, []):
            if issue_detail not in problems:
                problems.append(issue_detail)
        if issues_by_part.get(part_number):
            actions.append("Correct the indicated source-data row(s), then rerun.")

        if "unexpectedly found in linear nesting" in error_text.lower():
            if row["Category"] == "PLATE":
                problems.append("Plate unexpectedly found in linear nesting")
            elif row["Category"] == "EXCLUDED HARDWARE":
                problems.append(
                    "Excluded hardware unexpectedly found in linear nesting"
                )
            actions.append("Remove or reclassify the unexpected nesting entry.")

        if status == "NOT CHECKED" and not_checked_text:
            problems.extend(
                part.strip()
                for part in not_checked_text.split(";")
                if part.strip()
            )
            if row["Category"] == "UNKNOWN":
                actions.append(
                    "Supply a usable description so the part can be classified."
                )
            elif "blank" in not_checked_text.lower():
                actions.append("Supply the missing comparison value, then rerun.")

        if status == "EXACT MATCH":
            problems.append("All required checks match")

        if not problems:
            fallback = error_text or not_checked_text or status
            problems.append(fallback)

        compact_rows.append(
            {
                "Part Number": part_number,
                "Review Status": status,
                "Category": row["Category"],
                "Description": row["Description"],
                "Problem": "\n".join(unique_preserve_order(problems)),
                "SolidWorks Value": "\n".join(
                    unique_preserve_order(solidworks_values)
                ),
                "Parts List Value": "\n".join(
                    unique_preserve_order(parts_values)
                ),
                "Nest Value": "\n".join(
                    unique_preserve_order(nest_values)
                ),
                "Required Action": "\n".join(
                    unique_preserve_order(actions)
                ),
            }
        )

    return compact_rows


def html_cell(value: Any) -> str:
    escaped = html.escape(str(value or ""))
    return escaped.replace("\n", "<br>")


def html_table(
    table_id: str,
    rows: list[dict[str, str]],
    headers: list[str],
    empty_message: str,
    sticky_columns: int = 0,
    compact: bool = False,
) -> str:
    if not rows:
        return f'<p class="empty">{html.escape(empty_message)}</p>'

    header_html = "".join(
        f'<th class="column-{index + 1}">{html.escape(header)}</th>'
        for index, header in enumerate(headers)
    )

    body_rows: list[str] = []

    for row in rows:
        status_class = re.sub(
            r"[^a-z0-9]+",
            "-",
            row.get("Review Status", "").lower(),
        ).strip("-")

        cells = "".join(
            (
                f'<td class="column-{index + 1}" '
                f'data-header="{html.escape(header)}">'
                f'{html_cell(row.get(header, ""))}</td>'
            )
            for index, header in enumerate(headers)
        )

        body_rows.append(f'<tr class="{status_class}">{cells}</tr>')

    table_classes = ["report-table"]
    if sticky_columns:
        table_classes.append(f"sticky-{sticky_columns}")
    if compact:
        table_classes.append("compact-table")

    return f"""
    <input
        class="filter"
        type="search"
        placeholder="Filter by part number, problem, or value..."
        oninput="filterTable('{table_id}', this.value)"
    >
    <div class="table-wrap">
        <table id="{table_id}" class="{' '.join(table_classes)}">
            <thead><tr>{header_html}</tr></thead>
            <tbody>{''.join(body_rows)}</tbody>
        </table>
    </div>
    """


def write_html_report(
    path: Path,
    rows: list[dict[str, str]],
    issues: list[Issue],
    metadata: dict[str, str],
) -> None:
    compact_rows = build_compact_rows(rows, issues)
    action_rows = [
        row for row in compact_rows
        if row["Review Status"] in ACTION_STATUSES
    ]

    status_counts: dict[str, int] = defaultdict(int)
    for row in rows:
        status_counts[row["Review Status"]] += 1

    issue_rows = [issue.as_dict() for issue in issues]

    metadata_html = "".join(
        f"<li><strong>{html.escape(key)}:</strong> "
        f"{html.escape(value)}</li>"
        for key, value in metadata.items()
    )

    document = f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Production Part Reconciliation</title>
<style>
    :root {{
        color-scheme: light;
        font-family: Arial, Helvetica, sans-serif;
        --part-width: 165px;
        --status-width: 205px;
        --category-width: 145px;
        --description-width: 330px;
    }}

    * {{ box-sizing: border-box; }}

    body {{
        margin: 24px;
        color: #172033;
        background: #f4f7fb;
    }}

    h1 {{ margin-bottom: 4px; }}
    h2 {{ margin-top: 34px; margin-bottom: 10px; }}

    .subtitle {{ margin-top: 0; color: #58657a; }}

    .summary {{
        display: flex;
        flex-wrap: wrap;
        gap: 12px;
        margin: 22px 0;
    }}

    .card {{
        min-width: 170px;
        padding: 14px 18px;
        border: 1px solid #c8d1df;
        border-radius: 9px;
        background: white;
        box-shadow: 0 1px 3px rgba(0, 0, 0, .05);
    }}

    .card strong {{
        display: block;
        margin-top: 5px;
        font-size: 25px;
    }}

    .error-card strong {{ color: #b42318; }}
    .missing-card strong {{ color: #9a6700; }}
    .unchecked-card strong {{ color: #175cd3; }}
    .match-card strong {{ color: #067647; }}

    .panel {{
        margin: 18px 0;
        padding: 16px;
        border: 1px solid #c8d1df;
        border-radius: 9px;
        background: white;
    }}

    .links a {{
        display: inline-block;
        margin: 4px 18px 4px 0;
    }}

    .filter {{
        width: min(520px, 100%);
        margin: 4px 0 10px;
        padding: 8px 10px;
        border: 1px solid #98a2b3;
        border-radius: 6px;
    }}

    .table-wrap {{
        max-height: 70vh;
        overflow: auto;
        border: 1px solid #98a2b3;
        background: white;
    }}

    table {{
        width: max-content;
        min-width: 100%;
        border-collapse: separate;
        border-spacing: 0;
        font-size: 12.5px;
    }}

    th {{
        position: sticky;
        top: 0;
        z-index: 4;
        padding: 8px;
        border-right: 1px solid #667085;
        border-bottom: 1px solid #667085;
        color: white;
        background: #344054;
        white-space: nowrap;
        text-align: left;
    }}

    td {{
        padding: 7px 8px;
        border-right: 1px solid #d0d5dd;
        border-bottom: 1px solid #d0d5dd;
        white-space: nowrap;
        vertical-align: top;
        background: white;
    }}

    .compact-table td:nth-child(n+4) {{
        max-width: 430px;
        white-space: normal;
        line-height: 1.35;
    }}

    .compact-table .column-1 {{ width: var(--part-width); min-width: var(--part-width); }}
    .compact-table .column-2 {{ width: var(--status-width); min-width: var(--status-width); }}
    .compact-table .column-3 {{ width: var(--category-width); min-width: var(--category-width); }}
    .compact-table .column-4 {{ width: var(--description-width); min-width: var(--description-width); }}

    .sticky-4 .column-1 {{
        position: sticky;
        left: 0;
        z-index: 3;
        font-weight: bold;
        background: #eaf2f8;
    }}

    .sticky-4 .column-2 {{
        position: sticky;
        left: var(--part-width);
        z-index: 3;
    }}

    .sticky-4 .column-3 {{
        position: sticky;
        left: calc(var(--part-width) + var(--status-width));
        z-index: 3;
        background: #fff;
    }}

    .sticky-4 .column-4 {{
        position: sticky;
        left: calc(var(--part-width) + var(--status-width) + var(--category-width));
        z-index: 3;
        background: #fff;
        box-shadow: 4px 0 5px -4px rgba(0, 0, 0, .35);
    }}

    .sticky-4 thead .column-1,
    .sticky-4 thead .column-2,
    .sticky-4 thead .column-3,
    .sticky-4 thead .column-4 {{
        z-index: 6;
        background: #344054;
    }}

    tbody tr:hover td {{ background: #eef4ff; }}
    tbody tr:hover td.column-1,
    tbody tr:hover td.column-3,
    tbody tr:hover td.column-4 {{ background: #eef4ff; }}

    tr.error-requiring-action td.column-2 {{
        color: #7a271a;
        background: #fee4e2;
        font-weight: bold;
    }}

    tr.missing-from-solidworks-or-parts-list td.column-2 {{
        color: #7a2e0e;
        background: #fef0c7;
        font-weight: bold;
    }}

    tr.not-checked td.column-2 {{
        color: #1849a9;
        background: #d1e9ff;
        font-weight: bold;
    }}

    tr.exact-match td.column-2 {{
        color: #05603a;
        background: #d1fadf;
        font-weight: bold;
    }}

    .empty {{
        padding: 12px;
        border-radius: 6px;
        color: #05603a;
        background: #ecfdf3;
    }}

    details {{ margin: 22px 0; }}
    summary {{ cursor: pointer; font-size: 20px; font-weight: bold; }}

    ul.metadata {{ columns: 2; padding-left: 24px; }}

    @media (max-width: 1100px) {{
        :root {{
            --part-width: 145px;
            --status-width: 170px;
            --category-width: 115px;
            --description-width: 250px;
        }}
    }}

    @media (max-width: 900px) {{
        ul.metadata {{ columns: 1; }}
    }}
</style>
<script>
function filterTable(tableId, query) {{
    const normalized = query.toLowerCase();
    const rows = document.querySelectorAll(`#${{tableId}} tbody tr`);

    rows.forEach(row => {{
        row.style.display = row.innerText.toLowerCase().includes(normalized)
            ? ""
            : "none";
    }});
}}
</script>
</head>
<body>
<h1>Production Part Reconciliation</h1>
<p class="subtitle">
SolidWorks is the engineering baseline. Only mismatched values are shown in the
compact reports; complete source values remain in Technical Details.
</p>

<div class="summary">
    <div class="card">Parts checked<strong>{len(rows)}</strong></div>
    <div class="card error-card">
        Items requiring action
        <strong>{len(action_rows)}</strong>
    </div>
    <div class="card missing-card">
        Missing core parts
        <strong>{status_counts["MISSING FROM SOLIDWORKS OR PARTS LIST"]}</strong>
    </div>
    <div class="card unchecked-card">
        Not checked
        <strong>{status_counts["NOT CHECKED"]}</strong>
    </div>
    <div class="card match-card">
        Exact matches
        <strong>{status_counts["EXACT MATCH"]}</strong>
    </div>
    <div class="card">
        Source data issues
        <strong>{len(issues)}</strong>
    </div>
</div>

<div class="panel">
    <h2>Run information</h2>
    <ul class="metadata">{metadata_html}</ul>

    <p class="links">
        <a href="production_part_comparison.xlsx">Formatted Excel report</a>
        <a href="errors_requiring_action.csv">Errors requiring action CSV</a>
        <a href="all_comparisons.csv">All comparisons CSV</a>
        <a href="technical_details.csv">Technical details CSV</a>
        <a href="source_data_issues.csv">Source data issues</a>
        <a href="run_summary.txt">Run summary</a>
    </p>
</div>

<h2>Items requiring action</h2>
{html_table(
    "errors-table",
    action_rows,
    COMPACT_REPORT_HEADERS,
    "No comparison errors or missing core parts were found.",
    sticky_columns=4,
    compact=True,
)}

<h2>All comparisons</h2>
{html_table(
    "all-table",
    compact_rows,
    COMPACT_REPORT_HEADERS,
    "No parts were found.",
    sticky_columns=4,
    compact=True,
)}

<details>
<summary>Technical details ({len(rows)})</summary>
{html_table(
    "technical-table",
    rows,
    REPORT_HEADERS,
    "No technical detail rows were found.",
    sticky_columns=0,
    compact=False,
)}
</details>

<details>
<summary>Source data issues ({len(issue_rows)})</summary>
{html_table(
    "issues-table",
    issue_rows,
    ISSUE_HEADERS,
    "No source-data issues were found.",
    sticky_columns=0,
    compact=False,
)}
</details>

</body>
</html>
"""

    path.write_text(document, encoding="utf-8")


def excel_column_name(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def clean_excel_text(value: Any) -> str:
    text = str(value or "")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    if len(text) > 32767:
        text = text[:32766] + "…"
    return text


def xlsx_cell(reference: str, value: Any, style_index: int) -> str:
    text = clean_excel_text(value)
    escaped = xml_escape(text, {'"': "&quot;"})
    preserve = ' xml:space="preserve"' if text != text.strip() or "\n" in text else ""
    return (
        f'<c r="{reference}" s="{style_index}" t="inlineStr">'
        f'<is><t{preserve}>{escaped}</t></is></c>'
    )


def xlsx_status_style(status: str) -> int:
    return {
        "ERROR REQUIRING ACTION": 3,
        "MISSING FROM SOLIDWORKS OR PARTS LIST": 4,
        "NOT CHECKED": 5,
        "EXACT MATCH": 6,
    }.get(status, 8)


def xlsx_sheet_xml(
    rows: list[dict[str, str]],
    headers: list[str],
    widths: list[float],
    freeze_columns: int,
    status_column: str = "Review Status",
) -> str:
    max_row = max(1, len(rows) + 1)
    max_col = max(1, len(headers))
    dimension = f"A1:{excel_column_name(max_col)}{max_row}"

    columns_xml = "".join(
        f'<col min="{index}" max="{index}" width="{width}" customWidth="1"/>'
        for index, width in enumerate(widths, start=1)
    )

    if freeze_columns:
        top_left = f"{excel_column_name(freeze_columns + 1)}2"
        pane_xml = (
            f'<pane xSplit="{freeze_columns}" ySplit="1" '
            f'topLeftCell="{top_left}" activePane="bottomRight" state="frozen"/>'
            f'<selection pane="bottomRight" activeCell="{top_left}" sqref="{top_left}"/>'
        )
    else:
        pane_xml = (
            '<pane ySplit="1" topLeftCell="A2" '
            'activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
        )

    header_cells = "".join(
        xlsx_cell(f"{excel_column_name(index)}1", header, 1)
        for index, header in enumerate(headers, start=1)
    )
    sheet_rows = [
        f'<row r="1" ht="26" customHeight="1">{header_cells}</row>'
    ]

    status_index = (
        headers.index(status_column) + 1
        if status_column in headers
        else None
    )

    for row_number, row in enumerate(rows, start=2):
        estimated_line_counts: list[int] = []
        for header, width in zip(headers, widths):
            text_value = clean_excel_text(row.get(header, ""))
            character_capacity = max(10, int(width * 1.15))
            estimated_lines = sum(
                max(1, math.ceil(len(line) / character_capacity))
                for line in (text_value.split("\n") or [""])
            )
            estimated_line_counts.append(estimated_lines)

        line_count = max(1, *estimated_line_counts)
        row_height = min(120, max(22, 8 + line_count * 17))
        cells: list[str] = []

        for column_number, header in enumerate(headers, start=1):
            value = row.get(header, "")

            if column_number == 1:
                style_index = 2
            elif status_index == column_number:
                style_index = xlsx_status_style(str(value))
            elif header in {
                "Description",
                "Problem",
                "SolidWorks Value",
                "Parts List Value",
                "Nest Value",
                "Required Action",
                "Errors",
                "Not Checked Reasons",
                "Details",
            }:
                style_index = 7
            else:
                style_index = 8

            reference = f"{excel_column_name(column_number)}{row_number}"
            cells.append(xlsx_cell(reference, value, style_index))

        sheet_rows.append(
            f'<row r="{row_number}" ht="{row_height}" customHeight="1">'
            f'{"".join(cells)}</row>'
        )

    auto_filter = f"A1:{excel_column_name(max_col)}{max_row}"

    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <dimension ref="{dimension}"/>
    <sheetViews>
        <sheetView tabSelected="0" workbookViewId="0">{pane_xml}</sheetView>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="18"/>
    <cols>{columns_xml}</cols>
    <sheetData>{''.join(sheet_rows)}</sheetData>
    <autoFilter ref="{auto_filter}"/>
</worksheet>'''


def xlsx_styles_xml() -> str:
    return '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <fonts count="3">
        <font><sz val="11"/><name val="Calibri"/><family val="2"/></font>
        <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>
        <font><b/><color rgb="FF172033"/><sz val="11"/><name val="Calibri"/></font>
    </fonts>
    <fills count="8">
        <fill><patternFill patternType="none"/></fill>
        <fill><patternFill patternType="gray125"/></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FF344054"/><bgColor indexed="64"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFEAF2F8"/><bgColor indexed="64"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFFEE4E2"/><bgColor indexed="64"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFFEF0C7"/><bgColor indexed="64"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFD1E9FF"/><bgColor indexed="64"/></patternFill></fill>
        <fill><patternFill patternType="solid"><fgColor rgb="FFD1FADF"/><bgColor indexed="64"/></patternFill></fill>
    </fills>
    <borders count="2">
        <border><left/><right/><top/><bottom/><diagonal/></border>
        <border>
            <left style="thin"><color rgb="FFD0D5DD"/></left>
            <right style="thin"><color rgb="FFD0D5DD"/></right>
            <top style="thin"><color rgb="FFD0D5DD"/></top>
            <bottom style="thin"><color rgb="FFD0D5DD"/></bottom>
            <diagonal/>
        </border>
    </borders>
    <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
    <cellXfs count="9">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
        <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>
        <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf>
        <xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
        <xf numFmtId="0" fontId="2" fillId="5" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
        <xf numFmtId="0" fontId="2" fillId="6" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
        <xf numFmtId="0" fontId="2" fillId="7" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top" wrapText="1"/></xf>
        <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyAlignment="1"><alignment vertical="top"/></xf>
    </cellXfs>
    <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''


def write_xlsx_report(
    path: Path,
    compact_rows: list[dict[str, str]],
    action_rows: list[dict[str, str]],
    technical_rows: list[dict[str, str]],
    issues: list[Issue],
) -> None:
    issue_rows = [issue.as_dict() for issue in issues]

    compact_widths = [
        20, 30, 20, 42, 48, 32, 32, 32, 48,
    ]
    technical_widths = [
        20, 30, 20, 42,
        *([18] * (len(REPORT_HEADERS) - 4)),
    ]
    issue_widths = [20, 12, 34, 18, 48, 10, 70]

    sheet_specs = [
        (
            "Errors Requiring Action",
            action_rows,
            COMPACT_REPORT_HEADERS,
            compact_widths,
            4,
            False,
        ),
        (
            "All Comparisons",
            compact_rows,
            COMPACT_REPORT_HEADERS,
            compact_widths,
            4,
            False,
        ),
        (
            "Technical Details",
            technical_rows,
            REPORT_HEADERS,
            technical_widths,
            4,
            True,
        ),
        (
            "Source Data Issues",
            issue_rows,
            ISSUE_HEADERS,
            issue_widths,
            1,
            True,
        ),
    ]

    content_overrides = "".join(
        (
            f'<Override PartName="/xl/worksheets/sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.'
            'spreadsheetml.worksheet+xml"/>'
        )
        for index in range(1, len(sheet_specs) + 1)
    )

    content_types = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
    {content_overrides}
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>'''

    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>'''

    sheet_entries = []
    workbook_rels = []
    for index, (name, *_rest, hidden) in enumerate(sheet_specs, start=1):
        hidden_attribute = ' state="hidden"' if hidden else ""
        sheet_entries.append(
            f'<sheet name="{xml_escape(name)}" sheetId="{index}" '
            f'r:id="rId{index}"{hidden_attribute}/>'
        )
        workbook_rels.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )

    styles_rel_id = len(sheet_specs) + 1
    workbook_rels.append(
        f'<Relationship Id="rId{styles_rel_id}" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )

    workbook_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <bookViews><workbookView activeTab="0"/></bookViews>
    <sheets>{''.join(sheet_entries)}</sheets>
</workbook>'''

    workbook_rels_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    {''.join(workbook_rels)}
</Relationships>'''

    now_utc = (
        datetime.now(timezone.utc)
        .replace(microsecond=0)
        .isoformat()
        .replace("+00:00", "Z")
    )
    core_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>Production Part Reconciliation</dc:creator>
    <cp:lastModifiedBy>Production Part Reconciliation</cp:lastModifiedBy>
    <dcterms:created xsi:type="dcterms:W3CDTF">{now_utc}</dcterms:created>
    <dcterms:modified xsi:type="dcterms:W3CDTF">{now_utc}</dcterms:modified>
</cp:coreProperties>'''

    sheet_names = [name for name, *_ in sheet_specs]
    app_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
    <Application>Microsoft Excel Compatible</Application>
    <DocSecurity>0</DocSecurity>
    <ScaleCrop>false</ScaleCrop>
    <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{len(sheet_specs)}</vt:i4></vt:variant></vt:vector></HeadingPairs>
    <TitlesOfParts><vt:vector size="{len(sheet_specs)}" baseType="lpstr">{''.join(f'<vt:lpstr>{xml_escape(name)}</vt:lpstr>' for name in sheet_names)}</vt:vector></TitlesOfParts>
    <Company></Company>
    <LinksUpToDate>false</LinksUpToDate>
    <SharedDoc>false</SharedDoc>
    <HyperlinksChanged>false</HyperlinksChanged>
    <AppVersion>16.0300</AppVersion>
</Properties>'''

    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", workbook_rels_xml)
        archive.writestr("xl/styles.xml", xlsx_styles_xml())
        archive.writestr("docProps/core.xml", core_xml)
        archive.writestr("docProps/app.xml", app_xml)

        for index, (_name, sheet_rows, headers, widths, freeze, _hidden) in enumerate(
            sheet_specs,
            start=1,
        ):
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                xlsx_sheet_xml(
                    sheet_rows,
                    headers,
                    widths,
                    freeze_columns=freeze,
                ),
            )


def create_output_folder(
    base_folder: Path,
    rules: dict[str, Any],
) -> Path:
    if rules["output"].get("create_timestamped_subfolder", True):
        timestamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
        folder = base_folder / f"Part_Comparison_{timestamp}"
    else:
        folder = base_folder

    folder.mkdir(parents=True, exist_ok=True)
    return folder


def write_summary(
    path: Path,
    rows: list[dict[str, str]],
    issues: list[Issue],
    metadata: dict[str, str],
) -> None:
    counts: dict[str, int] = defaultdict(int)

    for row in rows:
        counts[row["Review Status"]] += 1

    issue_counts: dict[str, int] = defaultdict(int)
    for issue in issues:
        issue_counts[issue.severity] += 1

    lines = [
        "PRODUCTION PART RECONCILIATION",
        "=" * 34,
        "",
        *[f"{key}: {value}" for key, value in metadata.items()],
        "",
        f"Parts checked: {len(rows)}",
        (
            "Errors requiring action: "
            f'{counts["ERROR REQUIRING ACTION"]}'
        ),
        (
            "Missing from SolidWorks or Parts List: "
            f'{counts["MISSING FROM SOLIDWORKS OR PARTS LIST"]}'
        ),
        f'Not checked: {counts["NOT CHECKED"]}',
        f'Exact matches: {counts["EXACT MATCH"]}',
        "",
        f"Source issues: {len(issues)}",
        f'  Errors: {issue_counts["ERROR"]}',
        f'  Warnings: {issue_counts["WARNING"]}',
        "",
        "Review comparison_report.html first.",
    ]

    path.write_text("\n".join(lines), encoding="utf-8")


# =============================================================================
# FILE SELECTION
# =============================================================================

def select_paths_with_dialogs() -> tuple[Path, Path, Path, Path]:
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk
    except ImportError as exc:
        raise RuntimeError(
            "Tkinter is unavailable. Run the script with command-line arguments."
        ) from exc

    steel_blue = "#57a0d3"
    dark_bg = "#101820"
    dark_surface = "#18232d"
    dark_input = "#111a22"
    text_primary = "#eef5f9"
    root = tk.Tk()
    root.title("Production Part Reconciliation")
    root.geometry("860x360")
    root.minsize(720, 330)
    root.configure(background=dark_bg)
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass
    style.configure("TFrame", background=dark_bg)
    style.configure("TLabel", background=dark_bg, foreground=text_primary)
    style.configure(
        "Heading.TLabel",
        background=dark_bg,
        foreground=text_primary,
        font=("Segoe UI Semibold", 18),
    )
    style.configure(
        "TEntry",
        fieldbackground=dark_input,
        foreground=text_primary,
        insertcolor=text_primary,
        bordercolor="#40515f",
    )
    style.configure(
        "TButton",
        background=dark_surface,
        foreground=text_primary,
        bordercolor="#40515f",
        padding=(10, 6),
    )
    style.map("TButton", background=[("active", "#2b3c49")])
    style.configure(
        "Accent.TButton",
        background=steel_blue,
        foreground="#ffffff",
        bordercolor=steel_blue,
        padding=(12, 7),
    )
    style.map("Accent.TButton", background=[("active", "#6eb1df")])

    values = {
        "nests": tk.StringVar(),
        "parts": tk.StringVar(),
        "solidworks": tk.StringVar(),
        "output": tk.StringVar(),
    }
    result: list[tuple[Path, Path, Path, Path]] = []
    frame = ttk.Frame(root, padding=20)
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(1, weight=1)
    ttk.Label(
        frame,
        text="Production Part Reconciliation",
        style="Heading.TLabel",
    ).grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 5))
    ttk.Label(
        frame,
        text="Choose the four production inputs, then run the comparison.",
        foreground="#a9bac6",
    ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 16))

    file_types = [("CSV or text files", "*.csv *.txt"), ("All files", "*.*")]

    def browse_folder(key: str, title: str) -> None:
        selected = filedialog.askdirectory(
            title=title,
            initialdir=values["nests"].get() or None,
            parent=root,
        )
        if selected:
            values[key].set(selected)

    def browse_file(key: str, title: str) -> None:
        selected = filedialog.askopenfilename(
            title=title,
            filetypes=file_types,
            parent=root,
        )
        if selected:
            values[key].set(selected)

    rows = (
        ("Nesting CSV folder", "nests", lambda: browse_folder("nests", "Select nesting folder")),
        ("Parts List CSV", "parts", lambda: browse_file("parts", "Select Parts List CSV")),
        ("SolidWorks CSV", "solidworks", lambda: browse_file("solidworks", "Select SolidWorks CSV")),
        ("Output folder", "output", lambda: browse_folder("output", "Select output folder")),
    )
    for row, (label, key, command) in enumerate(rows, start=2):
        ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=5)
        ttk.Entry(frame, textvariable=values[key]).grid(
            row=row, column=1, sticky="ew", padx=10, pady=5
        )
        ttk.Button(frame, text="Browse…", command=command).grid(
            row=row, column=2, pady=5
        )

    def accept() -> None:
        missing = [label for label, key, _command in rows if not values[key].get().strip()]
        if missing:
            messagebox.showerror(
                "Required inputs",
                "Select: " + ", ".join(missing),
                parent=root,
            )
            return
        result.append(
            tuple(Path(values[key].get()) for key in ("nests", "parts", "solidworks", "output"))
        )
        root.destroy()

    controls = ttk.Frame(frame)
    controls.grid(row=6, column=0, columnspan=3, sticky="e", pady=(16, 0))
    ttk.Button(controls, text="Cancel", command=root.destroy).pack(side="left", padx=6)
    ttk.Button(
        controls,
        text="Run Comparison",
        command=accept,
        style="Accent.TButton",
    ).pack(side="left")
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()
    if not result:
        raise SystemExit("Cancelled.")
    return result[0]


def show_completion_message(
    output_folder: Path,
    summary: str,
) -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        messagebox.showinfo(
            "Comparison complete",
            f"{summary}\n\nPrimary outputs:\n"
            f"{output_folder / 'production_part_comparison.xlsx'}\n"
            f"{output_folder / 'comparison_report.html'}",
            parent=root,
        )

        root.destroy()
    except Exception:
        pass


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description=(
            "Compare SolidWorks, Parts List, and linear nesting data."
        )
    )

    parser.add_argument(
        "--nests",
        type=Path,
        help="Folder containing linear nesting CSV/TXT files",
    )
    parser.add_argument(
        "--parts",
        type=Path,
        help="Parts List CSV export",
    )
    parser.add_argument(
        "--solidworks",
        type=Path,
        help="SolidWorks Assembly Visualization CSV export",
    )
    parser.add_argument(
        "--output",
        type=Path,
        help="Base output folder",
    )
    parser.add_argument(
        "--rules",
        type=Path,
        help="Optional path to comparison_rules.json",
    )
    parser.add_argument(
        "--no-open",
        action="store_true",
        help="Do not open the HTML report automatically",
    )

    return parser.parse_args()


# =============================================================================
# MAIN
# =============================================================================

def main() -> int:
    args = parse_arguments()

    script_folder = Path(__file__).resolve().parent
    rules_path = (
        args.rules
        if args.rules is not None
        else script_folder / "comparison_rules.json"
    )

    rules, configuration_issues = load_rules(rules_path)

    supplied = [
        args.nests,
        args.parts,
        args.solidworks,
        args.output,
    ]

    if any(value is not None for value in supplied):
        if not all(value is not None for value in supplied):
            raise ValueError(
                "When using command-line arguments, provide --nests, "
                "--parts, --solidworks, and --output."
            )

        nests_folder = args.nests
        parts_file = args.parts
        solidworks_file = args.solidworks
        output_base = args.output
    else:
        (
            nests_folder,
            parts_file,
            solidworks_file,
            output_base,
        ) = select_paths_with_dialogs()

    assert nests_folder is not None
    assert parts_file is not None
    assert solidworks_file is not None
    assert output_base is not None

    if not nests_folder.is_dir():
        raise FileNotFoundError(
            f"Nesting folder does not exist: {nests_folder}"
        )

    for selected_file in (parts_file, solidworks_file):
        if not selected_file.is_file():
            raise FileNotFoundError(
                f"Selected file does not exist: {selected_file}"
            )

    output_folder = create_output_folder(output_base, rules)

    print("Reading linear nesting files...", flush=True)
    nest_rows, nest_issues, usable_nest_files = parse_nest_folder(
        nests_folder,
        rules,
    )

    print("Reading Parts List...", flush=True)
    parts_source = parse_standard_source(
        parts_file,
        source_name="Parts List",
        require_file_name=False,
        rules=rules,
    )

    print("Reading SolidWorks Assembly Visualization...", flush=True)
    sw_source = parse_standard_source(
        solidworks_file,
        source_name="SolidWorks",
        require_file_name=True,
        rules=rules,
    )

    print("Aggregating part numbers...", flush=True)

    parts_quantity_header = parts_source.selected_headers.get(
        "quantity",
        "",
    )
    sw_quantity_header = sw_source.selected_headers.get(
        "quantity",
        "",
    )

    parts_quantity_mode = (
        "preaggregated"
        if parts_quantity_header in {"TOTAL QUANTITY", "TOTAL QTY"}
        else "sum"
    )
    sw_quantity_mode = (
        "preaggregated"
        if sw_quantity_header in {"TOTAL QUANTITY", "TOTAL QTY"}
        else "sum"
    )

    nest_aggregates, nest_aggregate_issues = aggregate_rows(
        nest_rows,
        rules,
        quantity_mode="sum",
    )
    parts_aggregates, parts_aggregate_issues = aggregate_rows(
        parts_source.rows,
        rules,
        quantity_mode=parts_quantity_mode,
    )
    sw_aggregates, sw_aggregate_issues = aggregate_rows(
        sw_source.rows,
        rules,
        quantity_mode=sw_quantity_mode,
    )

    all_issues = (
        configuration_issues
        + nest_issues
        + parts_source.issues
        + sw_source.issues
        + nest_aggregate_issues
        + parts_aggregate_issues
        + sw_aggregate_issues
    )

    source_fields = {
        "Nest": {
            "part_number",
            "quantity",
            "length",
            "shape",
            "material",
        },
        "Parts List": parts_source.available_fields,
        "SolidWorks": sw_source.available_fields,
    }

    print("Comparing against the SolidWorks baseline...", flush=True)
    comparison_rows = reconcile(
        nest_aggregates,
        parts_aggregates,
        sw_aggregates,
        source_fields,
        rules,
        all_issues,
    )

    compact_rows = build_compact_rows(comparison_rows, all_issues)
    compact_by_part = {
        row["Part Number"]: row
        for row in compact_rows
    }

    action_rows = [
        row for row in compact_rows
        if row["Review Status"] in ACTION_STATUSES
    ]

    compact_groups = {
        "missing_from_solidworks_or_parts_list.csv": [
            compact_by_part[row["Part Number"]]
            for row in comparison_rows
            if row["Review Status"]
            == "MISSING FROM SOLIDWORKS OR PARTS LIST"
        ],
        "not_checked.csv": [
            compact_by_part[row["Part Number"]]
            for row in comparison_rows
            if row["Review Status"] == "NOT CHECKED"
        ],
        "exact_matches.csv": [
            compact_by_part[row["Part Number"]]
            for row in comparison_rows
            if row["Review Status"] == "EXACT MATCH"
        ],
    }

    write_dict_csv(
        output_folder / "all_comparisons.csv",
        compact_rows,
        COMPACT_REPORT_HEADERS,
    )

    write_dict_csv(
        output_folder / "errors_requiring_action.csv",
        action_rows,
        COMPACT_REPORT_HEADERS,
    )

    for filename, compact_group_rows in compact_groups.items():
        write_dict_csv(
            output_folder / filename,
            compact_group_rows,
            COMPACT_REPORT_HEADERS,
        )

    write_dict_csv(
        output_folder / "technical_details.csv",
        comparison_rows,
        REPORT_HEADERS,
    )

    write_dict_csv(
        output_folder / "source_data_issues.csv",
        [issue.as_dict() for issue in all_issues],
        ISSUE_HEADERS,
    )

    all_source_rows = (
        nest_rows
        + parts_source.rows
        + sw_source.rows
    )

    write_dict_csv(
        output_folder / "source_rows.csv",
        [source_row_to_dict(row) for row in all_source_rows],
        SOURCE_ROW_HEADERS,
    )

    metadata = {
        "Run time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "SolidWorks baseline file": str(solidworks_file),
        "Parts List file": str(parts_file),
        "Linear nesting folder": str(nests_folder),
        "Usable nesting files": str(usable_nest_files),
        "Rules file": str(rules_path),
        "Length tolerance": (
            f'{rules["length_tolerance_inches"]} in'
        ),
        "Width tolerance": (
            f'{rules["width_tolerance_inches"]} in'
        ),
        "Parts List quantity column": (
            f"{parts_quantity_header} ({parts_quantity_mode})"
        ),
        "SolidWorks quantity column": (
            f"{sw_quantity_header} ({sw_quantity_mode})"
        ),
        "SolidWorks configuration suffixes removed": ", ".join(
            rules["solidworks_configuration_suffixes"]
        ),
    }

    if rules["output"].get("create_excel_workbook", True):
        write_xlsx_report(
            output_folder / "production_part_comparison.xlsx",
            compact_rows,
            action_rows,
            comparison_rows,
            all_issues,
        )

    write_html_report(
        output_folder / "comparison_report.html",
        comparison_rows,
        all_issues,
        metadata,
    )

    write_summary(
        output_folder / "run_summary.txt",
        comparison_rows,
        all_issues,
        metadata,
    )

    counts: dict[str, int] = defaultdict(int)
    for row in comparison_rows:
        counts[row["Review Status"]] += 1

    error_count = counts["ERROR REQUIRING ACTION"]
    missing_core_count = counts["MISSING FROM SOLIDWORKS OR PARTS LIST"]
    not_checked_count = counts["NOT CHECKED"]
    exact_match_count = counts["EXACT MATCH"]
    source_issue_count = len(all_issues)
    outcome = (
        "action_required"
        if error_count or missing_core_count
        else "review_recommended"
        if not_checked_count or source_issue_count
        else "no_discrepancies"
    )

    machine_summary = {
        "schema_version": 1,
        "run_time": datetime.now().astimezone().isoformat(timespec="seconds"),
        "outcome": outcome,
        "counts": {
            "part_numbers": len(comparison_rows),
            "errors": error_count,
            "missing_core": missing_core_count,
            "not_checked": not_checked_count,
            "exact_matches": exact_match_count,
            "source_issues": source_issue_count,
        },
        "inputs": {
            "solidworks": str(solidworks_file.resolve()),
            "parts_list": str(parts_file.resolve()),
            "nesting_folder": str(nests_folder.resolve()),
        },
        "reports": {
            "excel": str(
                (output_folder / "production_part_comparison.xlsx").resolve()
            ),
            "html": str((output_folder / "comparison_report.html").resolve()),
            "folder": str(output_folder.resolve()),
        },
    }
    (output_folder / "comparison_summary.json").write_text(
        json.dumps(machine_summary, indent=2) + "\n",
        encoding="utf-8",
    )

    summary = (
        f"Checked {len(comparison_rows)} part numbers: "
        f"{error_count} errors, "
        f"{missing_core_count} missing core, "
        f"{not_checked_count} not checked, "
        f"{exact_match_count} exact matches, "
        f"{source_issue_count} source issues."
    )

    print("\n" + summary, flush=True)
    print(
        f"Excel report: {output_folder / 'production_part_comparison.xlsx'}",
        flush=True,
    )
    print(
        f"HTML report: {output_folder / 'comparison_report.html'}",
        flush=True,
    )

    should_open = (
        rules["output"].get(
            "open_html_report_when_finished",
            True,
        )
        and not args.no_open
    )

    if should_open:
        try:
            webbrowser.open(
                (output_folder / "comparison_report.html").resolve().as_uri()
            )
        except Exception:
            pass

    # --no-open is also the automation/headless contract. Do not leave the
    # Job Assistant waiting on a second process-owned modal dialog.
    if not args.no_open:
        show_completion_message(output_folder, summary)

    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit:
        raise
    except Exception as exc:
        print(f"\nERROR: {exc}", file=sys.stderr)

        if "--no-open" not in sys.argv:
            try:
                input("\nPress Enter to close...")
            except EOFError:
                pass

        raise SystemExit(1) from None
