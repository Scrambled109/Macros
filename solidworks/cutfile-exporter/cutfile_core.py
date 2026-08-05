from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Sequence
from uuid import uuid4

import ezdxf
from ezdxf import bbox, edgeminer, edgesmith, path as ezdxf_path


OUTSIDE_LAYER = "CUT - OUTSIDE STRAIGHT"
INSIDE_LAYER = "CUT - INSIDE STRAIGHT"
LINE_MARKING_LAYER = "PIN STAMP LINE MARKING"
TEXT_MARKING_LAYER = "PIN STAMP TEXT"


class CutfileValidationError(RuntimeError):
    """Raised when a DXF cannot be proven safe enough to release as a cut file."""


@dataclass(frozen=True)
class LayeringResult:
    outside_loops: int
    inside_loops: int
    cut_entities: int
    drawing_units: str


@dataclass(frozen=True)
class MarkingResult:
    paths_added: int
    entities_added: int


@dataclass(frozen=True)
class _DetectedLoop:
    entities: tuple[object, ...]
    area: float


_UNIT_NAMES = {
    0: "unitless",
    1: "inches",
    2: "feet",
    4: "millimeters",
    5: "centimeters",
    6: "meters",
}


def load_dxf(path: str | Path):
    try:
        return ezdxf.readfile(str(path))
    except (OSError, ezdxf.DXFError) as exc:
        raise CutfileValidationError(f"Could not read SolidWorks DXF export: {exc}") from exc


def ensure_required_layers(doc) -> None:
    layer_specs = (
        (OUTSIDE_LAYER, 1),
        (INSIDE_LAYER, 5),
        (LINE_MARKING_LAYER, 3),
        (TEXT_MARKING_LAYER, 7),
    )
    for name, color in layer_specs:
        if name not in doc.layers:
            doc.layers.add(name=name, color=color)
        else:
            doc.layers.get(name).color = color


def drawing_unit_name(doc) -> str:
    try:
        code = int(doc.header.get("$INSUNITS", 0))
    except (TypeError, ValueError):
        code = 0
    return _UNIT_NAMES.get(code, f"DXF unit code {code}")


def _payload_key(payload) -> str:
    handle = getattr(getattr(payload, "dxf", None), "handle", None)
    return str(handle or id(payload))


def assign_cut_layers(
    doc,
    *,
    expected_outer_loops: int = 1,
    expected_inner_loops: int | None = None,
    gap_tolerance: float = 1e-6,
) -> LayeringResult:
    """Assign exported face geometry to outside/inside layers.

    SolidWorks supplies the expected topological loop counts.  The independently
    exported DXF must reproduce the same count and every cut entity must belong
    to exactly one closed loop.  Anything ambiguous fails instead of producing a
    potentially incorrect production file.
    """

    if expected_outer_loops != 1:
        raise CutfileValidationError(
            "The selected SolidWorks face must have exactly one outside loop; "
            f"SolidWorks reported {expected_outer_loops}."
        )

    ensure_required_layers(doc)
    modelspace = doc.modelspace()
    open_entities = list(edgesmith.filter_open_edges(modelspace))
    closed_entities = [
        entity
        for entity in modelspace
        if entity.dxftype()
        in {"CIRCLE", "ELLIPSE", "SPLINE", "LWPOLYLINE", "POLYLINE"}
        and edgesmith.is_closed_entity(entity)
    ]
    cut_entities = open_entities + closed_entities
    if not cut_entities:
        raise CutfileValidationError("The SolidWorks face export contains no cut geometry.")

    edges = list(edgesmith.edges_from_entities_2d(open_entities, gap_tol=gap_tolerance))
    if len(edges) != len(open_entities):
        raise CutfileValidationError(
            "One or more exported cut entities could not be interpreted as 2D geometry."
        )

    open_loops: list[Sequence[object]] = []
    if edges:
        try:
            open_loops = list(
                edgeminer.find_all_loops(
                    edgeminer.Deposit(edges, gap_tol=gap_tolerance), timeout=30.0
                )
            )
        except TimeoutError as exc:
            raise CutfileValidationError(
                "DXF loop detection timed out; the exported geometry is branching or ambiguous."
            ) from exc

    edge_use_counts = Counter(edge.id for loop in open_loops for edge in loop)
    used_edge_ids = set(edge_use_counts)
    if len(used_edge_ids) != len(edges):
        unused = len(edges) - len(used_edge_ids)
        raise CutfileValidationError(
            f"The face export contains {unused} open, disconnected, or ambiguous edge(s)."
        )
    multiply_used = sum(count != 1 for count in edge_use_counts.values())
    if multiply_used:
        raise CutfileValidationError(
            "DXF loop detection reused "
            f"{multiply_used} edge(s); the exported geometry is branching or ambiguous."
        )

    loops: list[_DetectedLoop] = []
    for entity in closed_entities:
        try:
            area = _closed_entity_area(entity, gap_tolerance)
        except (ValueError, TypeError) as exc:
            raise CutfileValidationError(
                f"Could not calculate a closed profile area: {exc}"
            ) from exc
        loops.append(_DetectedLoop((entity,), area))

    for loop in open_loops:
        try:
            area = abs(float(edgesmith.loop_area(loop, gap_tol=gap_tolerance)))
        except (ValueError, TypeError) as exc:
            raise CutfileValidationError(
                f"Could not calculate a profile-loop area: {exc}"
            ) from exc
        payloads: list[object] = []
        for edge in loop:
            if edge.payload is None:
                raise CutfileValidationError("A detected DXF edge has no source entity.")
            payloads.append(edge.payload)
        loops.append(_DetectedLoop(tuple(payloads), area))

    if not loops:
        raise CutfileValidationError("No closed profile loop was found in the exported face.")

    inner_count = len(loops) - 1
    if expected_inner_loops is not None and inner_count != expected_inner_loops:
        raise CutfileValidationError(
            "SolidWorks/DXF topology mismatch: SolidWorks reported "
            f"{expected_inner_loops} inside loop(s), but the DXF contains {inner_count}."
        )

    areas = [loop.area for loop in loops]
    if not areas or max(areas) <= gap_tolerance * gap_tolerance:
        raise CutfileValidationError("The exported outside profile has zero or invalid area.")

    outside_index = max(range(len(areas)), key=areas.__getitem__)
    covered_payloads: set[str] = set()
    for index, loop in enumerate(loops):
        layer = OUTSIDE_LAYER if index == outside_index else INSIDE_LAYER
        for entity in loop.entities:
            entity.dxf.layer = layer
            covered_payloads.add(_payload_key(entity))

    unsupported = [
        entity
        for entity in modelspace
        if _payload_key(entity) not in covered_payloads
    ]
    if unsupported:
        kinds = ", ".join(sorted({entity.dxftype() for entity in unsupported}))
        raise CutfileValidationError(
            "The SolidWorks face export contains unsupported/unclassified model-space "
            f"entities: {kinds}."
        )

    return LayeringResult(
        outside_loops=1,
        inside_loops=inner_count,
        cut_entities=len(covered_payloads),
        drawing_units=drawing_unit_name(doc),
    )


def add_marking_paths(
    doc,
    paths: Iterable[Sequence[Sequence[float]]],
    *,
    layer: str = TEXT_MARKING_LAYER,
    close_tolerance: float = 1e-9,
) -> MarkingResult:
    ensure_required_layers(doc)
    modelspace = doc.modelspace()
    path_count = 0
    entity_count = 0

    for raw_path in paths:
        points = [(float(point[0]), float(point[1])) for point in raw_path]
        cleaned: list[tuple[float, float]] = []
        for point in points:
            if not cleaned or _distance_sq(cleaned[-1], point) > close_tolerance**2:
                cleaned.append(point)
        if len(cleaned) < 2:
            continue

        is_closed = (
            len(cleaned) > 2
            and _distance_sq(cleaned[0], cleaned[-1]) <= close_tolerance**2
        )
        if is_closed:
            cleaned.pop()
        if len(cleaned) < 2:
            continue

        if len(cleaned) == 2:
            modelspace.add_line(cleaned[0], cleaned[1], dxfattribs={"layer": layer})
        else:
            modelspace.add_lwpolyline(
                cleaned,
                close=is_closed,
                dxfattribs={"layer": layer},
            )
        path_count += 1
        entity_count += 1

    return MarkingResult(paths_added=path_count, entities_added=entity_count)


def infer_model_to_dxf_scale(
    doc,
    model_points_xy_m: Sequence[Sequence[float]],
    *,
    relative_tolerance: float = 0.03,
) -> tuple[float, str]:
    """Infer the SolidWorks export scale by comparing model and DXF extents."""

    if len(model_points_xy_m) < 2:
        raise CutfileValidationError("Not enough SolidWorks face points to verify DXF scale.")
    model_x = [float(point[0]) for point in model_points_xy_m]
    model_y = [float(point[1]) for point in model_points_xy_m]
    model_span = max(max(model_x) - min(model_x), max(model_y) - min(model_y))
    if model_span <= 1e-12:
        raise CutfileValidationError("The selected SolidWorks face has zero projected size.")

    cut_geometry = [
        entity
        for entity in doc.modelspace()
        if entity.dxf.layer in {OUTSIDE_LAYER, INSIDE_LAYER}
    ]
    extents = bbox.extents(cut_geometry, fast=True)
    if not extents.has_data:
        raise CutfileValidationError("Could not measure the SolidWorks DXF export.")
    dxf_span = max(float(extents.size.x), float(extents.size.y))
    if dxf_span <= 1e-12:
        raise CutfileValidationError("The SolidWorks DXF export has zero size.")

    ratio = dxf_span / model_span
    candidates = (
        (1.0, "meters"),
        (100.0, "centimeters"),
        (1000.0, "millimeters"),
        (39.37007874015748, "inches"),
        (3.280839895013123, "feet"),
    )
    scale, name = min(candidates, key=lambda item: abs(item[0] - ratio) / item[0])
    error = abs(scale - ratio) / scale
    if error > relative_tolerance:
        raise CutfileValidationError(
            "Could not verify SolidWorks-to-DXF units. "
            f"Measured scale {ratio:.6g}; nearest supported scale is {scale:.6g} ({name})."
        )
    return scale, name


def save_dxf(doc, path: str | Path) -> None:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    temporary = target.with_name(f".{target.stem}.{uuid4().hex}.tmp.dxf")
    try:
        doc.saveas(str(temporary))
        temporary.replace(target)
    except (OSError, ezdxf.DXFError) as exc:
        raise CutfileValidationError(f"Could not save final DXF: {exc}") from exc
    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass


def _distance_sq(a: Sequence[float], b: Sequence[float]) -> float:
    return (float(a[0]) - float(b[0])) ** 2 + (float(a[1]) - float(b[1])) ** 2


def _closed_entity_area(entity, gap_tolerance: float) -> float:
    path = ezdxf_path.make_path(entity)
    points = list(
        path.flattening(
            distance=max(float(gap_tolerance) * 10.0, 1e-5),
            segments=16,
        )
    )
    if len(points) < 3:
        raise ValueError(f"{entity.dxftype()} did not produce a closed polygon")
    area_twice = 0.0
    for index, point in enumerate(points):
        next_point = points[(index + 1) % len(points)]
        area_twice += float(point.x) * float(next_point.y)
        area_twice -= float(next_point.x) * float(point.y)
    return abs(area_twice) * 0.5
