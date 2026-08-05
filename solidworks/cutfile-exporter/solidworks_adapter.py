from __future__ import annotations

from dataclasses import dataclass
import math
from pathlib import Path
from typing import Iterable, Iterator, Sequence


SW_DOC_PART = 1
SW_OPEN_SILENT = 1
SW_BODY_SOLID = 0
SW_EXPORT_SELECTED_FACES_OR_LOOPS = 2
SW_SKETCH_TEXT = 4


class SolidWorksExportError(RuntimeError):
    pass


Vector3 = tuple[float, float, float]


@dataclass(frozen=True)
class PlaneFrame:
    origin: Vector3
    x_axis: Vector3
    y_axis: Vector3
    normal: Vector3

    def project_m(self, point: Sequence[float]) -> tuple[float, float]:
        delta = _sub(_vec3(point), self.origin)
        return _dot(delta, self.x_axis), _dot(delta, self.y_axis)

    def shifted(self, x_offset: float, y_offset: float) -> "PlaneFrame":
        origin = _add(
            self.origin,
            _add(_scale(self.x_axis, x_offset), _scale(self.y_axis, y_offset)),
        )
        return PlaneFrame(origin, self.x_axis, self.y_axis, self.normal)

    def alignment(self) -> tuple[float, ...]:
        return (*self.origin, *self.x_axis, *self.y_axis, *self.normal)


@dataclass(frozen=True)
class FaceExportInfo:
    face: object
    frame: PlaneFrame
    projected_points_m: tuple[tuple[float, float], ...]
    outer_loops: int
    inner_loops: int


@dataclass(frozen=True)
class OpenedPart:
    model: object
    path: Path
    opened_by_tool: bool


class SolidWorksSession:
    def __init__(self, app, *, started_by_tool: bool) -> None:
        self.app = app
        self.started_by_tool = started_by_tool

    @classmethod
    def connect(cls, *, visible: bool = True) -> "SolidWorksSession":
        if not _is_windows():
            raise SolidWorksExportError("SolidWorks automation requires Windows.")
        try:
            import win32com.client
        except ImportError as exc:
            raise SolidWorksExportError(
                "pywin32 is not installed. Run: py -m pip install -r requirements.txt"
            ) from exc

        started = False
        try:
            app = win32com.client.GetActiveObject("SldWorks.Application")
        except Exception:
            try:
                app = win32com.client.Dispatch("SldWorks.Application")
                started = True
            except Exception as exc:
                raise SolidWorksExportError(
                    "Could not start or connect to SolidWorks. Open SolidWorks and try again."
                ) from exc
        app = _dynamic_dispatch(app)
        try:
            app.Visible = bool(visible)
        except Exception:
            pass
        return cls(app, started_by_tool=started)

    def open_part(self, path: str | Path) -> OpenedPart:
        source = Path(path).resolve()
        if not source.is_file():
            raise SolidWorksExportError(f"Part file not found: {source}")

        try:
            active = _com_value(self.app, "ActiveDoc")
        except Exception:
            active = None
        if active is not None:
            try:
                if Path(_com_value(active, "GetPathName")).resolve() == source:
                    if int(_com_value(active, "GetType")) != SW_DOC_PART:
                        raise SolidWorksExportError(f"Not a SolidWorks part: {source}")
                    return OpenedPart(active, source, False)
            except SolidWorksExportError:
                raise
            except Exception:
                pass

        try:
            import pythoncom
            from win32com.client import VARIANT

            errors = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            warnings = VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            model = self.app.OpenDoc6(
                str(source), SW_DOC_PART, SW_OPEN_SILENT, "", errors, warnings
            )
            error_code = int(errors.value)
        except Exception as exc:
            raise SolidWorksExportError(f"SolidWorks could not open {source.name}: {exc}") from exc
        if model is None:
            raise SolidWorksExportError(
                f"SolidWorks could not open {source.name} (error code {error_code})."
            )
        return OpenedPart(model, source, True)

    def close_part(self, opened: OpenedPart) -> None:
        if not opened.opened_by_tool:
            return
        try:
            self.app.CloseDoc(_com_value(opened.model, "GetTitle"))
        except Exception:
            pass

    def find_export_face(self, model) -> FaceExportInfo:
        bodies = [body for body in _as_sequence(model.GetBodies2(SW_BODY_SOLID, True)) if body]
        if len(bodies) != 1:
            raise SolidWorksExportError(
                "The cut-file exporter currently requires exactly one visible solid body; "
                f"this part has {len(bodies)}."
            )

        candidates: list[
            tuple[
                float,
                float,
                object,
                PlaneFrame,
                tuple[object, ...],
                tuple[Vector3, ...],
            ]
        ] = []
        face_errors: list[str] = []
        for face in _as_sequence(_com_value(bodies[0], "GetFaces")):
            if face is None:
                continue
            try:
                normal_values = tuple(float(value) for value in _com_value(face, "Normal"))
                if len(normal_values) < 3:
                    raise ValueError("Normal returned fewer than three values")
                normal_length = math.sqrt(sum(value * value for value in normal_values[:3]))
                if normal_length <= 1e-12:
                    continue
                raw_normal = _normalize(_vec3(normal_values[0:3]))
                normal = _canonical_normal(raw_normal)
                loops = tuple(_as_sequence(_com_value(face, "GetLoops")))
                if not loops:
                    raise ValueError("face has no edge loops")
                model_points = tuple(_face_sample_points(face, loops))
                if not model_points:
                    raise ValueError("face returned no tessellation or edge points")
                origin = model_points[0]
                frame = _make_frame(origin, normal)
                area = float(_com_value(face, "GetArea"))
                plane_offset = _dot(origin, normal)
                candidates.append(
                    (area, plane_offset, face, frame, loops, model_points)
                )
            except Exception as exc:
                face_errors.append(f"{type(exc).__name__}: {exc}")
        if not candidates:
            detail = "; ".join(face_errors[:3])
            if len(face_errors) > 3:
                detail += f"; plus {len(face_errors) - 3} more face error(s)"
            suffix = f" SolidWorks details: {detail}" if detail else ""
            raise SolidWorksExportError(
                "No planar face was found in the SolidWorks part." + suffix
            )

        max_area = max(item[0] for item in candidates)
        area_tolerance = max(max_area * 1e-7, 1e-12)
        largest = [item for item in candidates if max_area - item[0] <= area_tolerance]
        selected_area, selected_offset, face, frame, loops, model_points = max(
            largest, key=lambda item: item[1]
        )
        model_span = max(
            max(point[axis] for point in model_points)
            - min(point[axis] for point in model_points)
            for axis in range(3)
        )
        plane_tolerance = max(model_span * 1e-7, 1e-9)
        coplanar_faces = [
            item
            for item in candidates
            if _dot(item[3].normal, frame.normal) >= 1.0 - 1e-8
            and abs(item[1] - selected_offset) <= plane_tolerance
        ]
        coplanar_area = sum(item[0] for item in coplanar_faces)
        if len(coplanar_faces) > 1 and coplanar_area > selected_area + area_tolerance:
            raise SolidWorksExportError(
                "The intended cut side is split into multiple coplanar faces. Exporting "
                "only the largest face would omit part of the perimeter. Merge/remove "
                "the split-face features so the full cut shape is one continuous planar face."
            )
        outer = 0
        for loop in loops:
            try:
                outer += 1 if bool(_com_value(loop, "IsOuter")) else 0
            except Exception as exc:
                raise SolidWorksExportError(
                    f"Could not classify a SolidWorks face loop: {exc}"
                ) from exc
        if outer != 1:
            raise SolidWorksExportError(
                f"The selected face has {outer} outside loops; exactly one is required."
            )

        projected = tuple(frame.project_m(point) for point in model_points)
        min_x = min(point[0] for point in projected)
        min_y = min(point[1] for point in projected)
        shifted_frame = frame.shifted(min_x, min_y)
        shifted_points = tuple(shifted_frame.project_m(point) for point in model_points)
        return FaceExportInfo(
            face=face,
            frame=shifted_frame,
            projected_points_m=shifted_points,
            outer_loops=outer,
            inner_loops=len(loops) - outer,
        )

    def export_face_dxf(
        self,
        opened: OpenedPart,
        face_info: FaceExportInfo,
        output_path: str | Path,
    ) -> None:
        target = Path(output_path).resolve()
        target.parent.mkdir(parents=True, exist_ok=True)
        model = opened.model
        try:
            model.ClearSelection2(True)
            _select_face(face_info.face)
            alignment = _double_array_variant(face_info.frame.alignment())
            ok = bool(
                _export_to_dwg2(
                    model,
                    str(target),
                    str(opened.path),
                    SW_EXPORT_SELECTED_FACES_OR_LOOPS,
                    True,
                    alignment,
                    False,
                    False,
                    0,
                    None,
                )
            )
        except Exception as exc:
            raise SolidWorksExportError(f"SolidWorks face export failed: {exc}") from exc
        finally:
            try:
                model.ClearSelection2(True)
            except Exception:
                pass
        if not ok or not target.is_file() or target.stat().st_size == 0:
            raise SolidWorksExportError("SolidWorks did not create a valid face DXF.")

    def marking_paths_model_space(
        self,
        model,
        sketch_name: str,
        *,
        curve_samples: int = 48,
    ) -> list[list[Vector3]]:
        feature = _find_feature(model, sketch_name)
        if feature is None:
            return []
        try:
            sketch = _com_value(feature, "GetSpecificFeature2")
        except Exception as exc:
            raise SolidWorksExportError(
                f"Could not read marking sketch '{sketch_name}': {exc}"
            ) from exc
        if sketch is None:
            raise SolidWorksExportError(
                f"Feature '{sketch_name}' exists but is not a readable SolidWorks sketch."
            )

        transform = None
        try:
            transform = _com_value(sketch.ModelToSketchTransform, "Inverse")
        except Exception:
            pass

        paths: list[list[Vector3]] = []
        for segment in _as_sequence(_com_value(sketch, "GetSketchSegments")):
            if segment is None:
                continue
            try:
                if bool(segment.ConstructionGeometry):
                    continue
            except Exception:
                pass
            try:
                segment_type = int(_com_value(segment, "GetType"))
            except Exception:
                segment_type = -1
            if segment_type != SW_SKETCH_TEXT:
                points = _sample_sketch_segment(segment, curve_samples)
                points = self._to_model_points(points, transform)
                if len(points) >= 2:
                    paths.append(points)
                continue

            try:
                edges = _as_sequence(_com_value(segment, "GetEdges2"))
            except Exception:
                try:
                    edges = _as_sequence(_com_value(segment, "GetEdges"))
                except Exception as exc:
                    raise SolidWorksExportError(
                        f"Could not render text from marking sketch '{sketch_name}': {exc}"
                    ) from exc
            if not edges:
                raise SolidWorksExportError(
                    f"Sketch text in '{sketch_name}' did not produce any rendered edges."
                )
            for edge in edges:
                points = _sample_edge(edge, curve_samples)
                points = self._to_model_points(points, transform)
                if len(points) >= 2:
                    paths.append(points)
        return paths

    def _to_model_points(self, points, transform) -> list[Vector3]:
        if transform is None:
            return [_vec3(point) for point in points]
        try:
            math_utility = _com_value(self.app, "GetMathUtility")
            converted: list[Vector3] = []
            for point in points:
                math_point = math_utility.CreatePoint(_double_array_variant(_vec3(point)))
                converted_point = math_point.MultiplyTransform(transform)
                converted.append(_vec3(_com_value(converted_point, "ArrayData")))
            return converted
        except Exception as exc:
            raise SolidWorksExportError(
                f"Could not transform marking sketch coordinates to model space: {exc}"
            ) from exc


def project_marking_paths(
    paths_model_m: Iterable[Sequence[Sequence[float]]],
    frame: PlaneFrame,
    scale: float,
    *,
    plane_tolerance_m: float = 1e-4,
) -> list[list[tuple[float, float]]]:
    projected: list[list[tuple[float, float]]] = []
    for path in paths_model_m:
        output_path: list[tuple[float, float]] = []
        for point_value in path:
            point = _vec3(point_value)
            distance = abs(_dot(_sub(point, frame.origin), frame.normal))
            if distance > plane_tolerance_m:
                raise SolidWorksExportError(
                    "The CUTFILE MARKING sketch is not on the exported face plane "
                    f"(offset {distance * 1000.0:.3f} mm)."
                )
            x, y = frame.project_m(point)
            output_path.append((x * scale, y * scale))
        if len(output_path) >= 2:
            projected.append(output_path)
    return projected


def _sample_face_edges(loops: Sequence[object]) -> Iterator[Vector3]:
    seen: set[int] = set()
    for loop in loops:
        for edge in _as_sequence(_com_value(loop, "GetEdges")):
            if edge is None or id(edge) in seen:
                continue
            seen.add(id(edge))
            yield from _sample_edge(edge, 24)


def _face_sample_points(face, loops: Sequence[object]) -> Iterator[Vector3]:
    try:
        values = tuple(float(value) for value in face.GetTessTriangles(True))
        if len(values) >= 9 and len(values) % 3 == 0:
            for index in range(0, len(values), 3):
                yield values[index], values[index + 1], values[index + 2]
            return
    except Exception:
        pass
    yield from _sample_face_edges(loops)


def _sample_edge(edge, samples: int) -> list[Vector3]:
    try:
        values = tuple(float(value) for value in _com_value(edge, "GetCurveParams2"))
        if len(values) >= 8:
            start = _vec3(values[0:3])
            end = _vec3(values[3:6])
            try:
                curve = _com_value(edge, "GetCurve")
                return _evaluate_curve(curve, values[6], values[7], samples)
            except Exception:
                return [start, end]
    except Exception:
        pass
    try:
        curve = _com_value(edge, "GetCurve")
        data = _com_value(edge, "GetCurveParams3")
        u_min = float(data.UMinValue)
        u_max = float(data.UMaxValue)
        return _evaluate_curve(curve, u_min, u_max, samples)
    except Exception:
        try:
            data = _com_value(edge, "GetCurveParams3")
            return [_vec3(data.StartPoint), _vec3(data.EndPoint)]
        except Exception:
            return []


def _sample_sketch_segment(segment, samples: int) -> list[Vector3]:
    try:
        curve = _com_value(segment, "GetCurve")
        values = _com_value(curve, "GetEndParams")
        if isinstance(values, (tuple, list)):
            offset = 1 if len(values) >= 5 and isinstance(values[0], bool) else 0
            if len(values) >= offset + 2:
                return _evaluate_curve(
                    curve,
                    float(values[offset]),
                    float(values[offset + 1]),
                    samples,
                )
    except Exception:
        pass
    try:
        start = _com_value(segment, "GetStartPoint2")
        end = _com_value(segment, "GetEndPoint2")
        return [
            (float(start.X), float(start.Y), float(start.Z)),
            (float(end.X), float(end.Y), float(end.Z)),
        ]
    except Exception:
        return []


def _evaluate_curve(curve, u_min: float, u_max: float, samples: int) -> list[Vector3]:
    count = max(1, int(samples))
    try:
        if bool(_com_value(curve, "IsLine")):
            count = 1
    except Exception:
        pass
    points: list[Vector3] = []
    for index in range(count + 1):
        parameter = u_min + (u_max - u_min) * index / count
        values = curve.Evaluate2(float(parameter), 0)
        points.append(_vec3(values[0:3]))
    return points


def _find_feature(model, requested_name: str):
    target = requested_name.strip().casefold()
    for feature in _walk_features(model):
        try:
            if str(feature.Name).strip().casefold() == target:
                return feature
        except Exception:
            continue
    return None


def _walk_features(model) -> Iterator[object]:
    feature = _com_value(model, "FirstFeature")
    while feature is not None:
        yield feature
        yield from _walk_subfeatures(feature)
        try:
            feature = _com_value(feature, "GetNextFeature")
        except Exception:
            break


def _walk_subfeatures(parent) -> Iterator[object]:
    try:
        feature = _com_value(parent, "GetFirstSubFeature")
    except Exception:
        return
    while feature is not None:
        yield feature
        yield from _walk_subfeatures(feature)
        try:
            feature = _com_value(feature, "GetNextSubFeature")
        except Exception:
            break


def _select_face(face) -> None:
    try:
        if bool(face.Select4(False, None)):
            return
    except Exception:
        pass
    try:
        if bool(face.Select2(False, 0)):
            return
    except Exception as exc:
        raise SolidWorksExportError(f"Could not select the export face: {exc}") from exc
    raise SolidWorksExportError("Could not select the export face.")


def _export_to_dwg2(model, *args):
    """Call the part-only exporter even when pywin32 returned IModelDoc2."""

    failures: list[str] = []
    for candidate in (model, _dynamic_dispatch(model)):
        try:
            method = getattr(candidate, "ExportToDWG2")
            return method(*args)
        except Exception as exc:
            failures.append(f"{type(exc).__name__}: {exc}")

    try:
        import win32com.client

        part_model = win32com.client.CastTo(model, "IPartDoc")
        return part_model.ExportToDWG2(*args)
    except Exception as exc:
        failures.append(f"IPartDoc cast: {type(exc).__name__}: {exc}")

    detail = "; ".join(failures[-3:])
    raise SolidWorksExportError(
        "SolidWorks did not expose the part DXF export interface. "
        f"Details: {detail}"
    )


def _dynamic_dispatch(obj):
    """Ignore stale makepy wrappers while preserving normal COM dispatch."""

    try:
        from win32com.client.dynamic import DumbDispatch

        return DumbDispatch(obj)
    except Exception:
        return obj


def _double_array_variant(values: Sequence[float]):
    try:
        import pythoncom
        from win32com.client import VARIANT

        return VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_R8,
            tuple(float(value) for value in values),
        )
    except ImportError:
        return tuple(float(value) for value in values)


def _make_frame(origin: Vector3, normal: Vector3) -> PlaneFrame:
    reference = (1.0, 0.0, 0.0)
    if abs(_dot(reference, normal)) > 0.95:
        reference = (0.0, 1.0, 0.0)
    x_axis = _normalize(_sub(reference, _scale(normal, _dot(reference, normal))))
    y_axis = _normalize(_cross(normal, x_axis))
    return PlaneFrame(origin, x_axis, y_axis, normal)


def _canonical_normal(normal: Vector3) -> Vector3:
    index = max(range(3), key=lambda i: abs(normal[i]))
    return _scale(normal, -1.0) if normal[index] < 0 else normal


def _as_sequence(value) -> list:
    if value is None:
        return []
    if isinstance(value, (list, tuple)):
        return list(value)
    return [value]


def _com_value(obj, name: str):
    """Read a zero-argument SolidWorks member exposed as a method or property.

    SolidWorks' generated pywin32 wrappers classify several legacy `Get...`
    members as properties, while dynamic dispatch exposes the same members as
    callable methods. Supporting both keeps the tool stable with or without a
    local makepy cache.
    """

    value = getattr(obj, name)
    # pywin32 COM objects implement ``__call__`` for an optional default
    # dispatch member, so ``callable(value)`` alone cannot distinguish a
    # returned COM property object (such as ActiveDoc) from a bound method.
    # Generated and dynamic pywin32 COM wrappers both expose ``_oleobj_``.
    if callable(value) and not hasattr(value, "_oleobj_"):
        return value()
    return value


def _is_windows() -> bool:
    import os

    return os.name == "nt"


def _vec3(value: Sequence[float]) -> Vector3:
    return float(value[0]), float(value[1]), float(value[2])


def _add(a: Vector3, b: Vector3) -> Vector3:
    return a[0] + b[0], a[1] + b[1], a[2] + b[2]


def _sub(a: Vector3, b: Vector3) -> Vector3:
    return a[0] - b[0], a[1] - b[1], a[2] - b[2]


def _scale(a: Vector3, scalar: float) -> Vector3:
    return a[0] * scalar, a[1] * scalar, a[2] * scalar


def _dot(a: Vector3, b: Vector3) -> float:
    return a[0] * b[0] + a[1] * b[1] + a[2] * b[2]


def _cross(a: Vector3, b: Vector3) -> Vector3:
    return (
        a[1] * b[2] - a[2] * b[1],
        a[2] * b[0] - a[0] * b[2],
        a[0] * b[1] - a[1] * b[0],
    )


def _normalize(value: Vector3) -> Vector3:
    length = math.sqrt(_dot(value, value))
    if length <= 1e-15:
        raise SolidWorksExportError("SolidWorks returned a zero-length face normal.")
    return _scale(value, 1.0 / length)
