from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class Issue:
    severity: str
    code: str
    message: str
    source_file: str = ""
    source_row: int | None = None
    part_no: str = ""
    value: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass
class PartRecord:
    source_file: str
    source_row: int
    loa: str
    drawing_id: str
    hull_no: str
    quantity: int | float | str | None
    mf_part_no: str
    part_no: str
    description: str
    part_description: str
    material_grade: str
    material_specification: str
    material_description: str
    thickness: str = ""
    width: str | float = ""
    length: str | float = ""
    length_inches: float | None = None
    size_thickness: str = ""
    unit_weight: float | None = None
    total_weight: float | None = None
    mdlprt: str = ""
    status: str = "EXPORTED"
    other_info: str = ""
    category: str = ""


@dataclass
class EPLResult:
    source_path: Path
    loa: str
    drawing_id: str
    plates: list[PartRecord] = field(default_factory=list)
    shapes: list[PartRecord] = field(default_factory=list)
    issues: list[Issue] = field(default_factory=list)
    assemblies_omitted: int = 0
    source_rows: int = 0
    bop_selected_rows: int = 0
    out_of_scope_rows: int = 0


@dataclass
class BOPEntry:
    source_file: str
    source_row: int
    hull_no: str
    loa: str
    part_no: str
    mdlprt: str = ""
    mf_part_no: str = ""
    quantity: int | float | str | None = None
    bom_line: str = ""
    revision: str = ""


@dataclass
class ConversionResult:
    plates_path: Path
    shapes_path: Path
    report_path: Path
    report_json_path: Path
    inputs: list[EPLResult]
    issues: list[Issue]
    bop_files: list[Path] = field(default_factory=list)

    @property
    def plate_count(self) -> int:
        return sum(len(x.plates) for x in self.inputs)

    @property
    def shape_count(self) -> int:
        return sum(len(x.shapes) for x in self.inputs)

    @property
    def assembly_count(self) -> int:
        return sum(x.assemblies_omitted for x in self.inputs)

    @property
    def unclassified_count(self) -> int:
        return sum(1 for x in self.issues if x.code == "UNCLASSIFIED_ITEM")

    def to_dict(self) -> dict[str, Any]:
        return {
            "ok": True,
            "outputs": {
                "plates": str(self.plates_path),
                "shapes": str(self.shapes_path),
                "report": str(self.report_path),
                "report_json": str(self.report_json_path),
            },
            "summary": {
                "input_files": len(self.inputs),
                "bop_files": len(self.bop_files),
                "plates_exported": self.plate_count,
                "shapes_exported": self.shape_count,
                "assemblies_omitted": self.assembly_count,
                "bop_selected_rows": sum(x.bop_selected_rows for x in self.inputs),
                "epl_rows_out_of_scope": sum(x.out_of_scope_rows for x in self.inputs),
                "unclassified_items": self.unclassified_count,
                "issues": len(self.issues),
            },
            "inputs": [
                {
                    "source_file": item.source_path.name,
                    "loa": item.loa,
                    "drawing_id": item.drawing_id,
                    "plates_exported": len(item.plates),
                    "shapes_exported": len(item.shapes),
                    "assemblies_omitted": item.assemblies_omitted,
                    "source_rows": item.source_rows,
                    "bop_selected_rows": item.bop_selected_rows,
                    "epl_rows_out_of_scope": item.out_of_scope_rows,
                }
                for item in self.inputs
            ],
            "issues": [issue.to_dict() for issue in self.issues],
        }
