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
class SolidWorksMarkingPaths:
    line_paths: tuple[tuple[Vector3, ...], ...]
    text_paths: tuple[tuple[Vector3, ...], ...]

    @property
    def total_paths(self) -> int:
        return len(self.line_paths) + len(self.text_paths)


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
            _activate_document(self.app, model)
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
    ) -> SolidWorksMarkingPaths:
        feature = _find_feature(model, sketch_name)
        if feature is None:
            return SolidWorksMarkingPaths((), ())
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

        try:
            sketch_to_model = _sketch_to_model_matrix(sketch)
        except Exception as exc:
            raise SolidWorksExportError(
                f"Could not read marking sketch '{sketch_name}' coordinates: {exc}"
            ) from exc

        line_paths: list[tuple[Vector3, ...]] = []
        text_paths: list[tuple[Vector3, ...]] = []
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
                points = self._to_model_points(points, sketch_to_model)
                if len(points) >= 2:
                    line_paths.append(tuple(points))
                continue

            try:
                edges = _sketch_text_edges(segment)
            except Exception as exc:
                raise SolidWorksExportError(
                    f"Could not render text from marking sketch '{sketch_name}': {exc}"
                ) from exc
            if not edges:
                raise SolidWorksExportError(
                    f"Sketch text in '{sketch_name}' did not produce any rendered edges."
                )
            for edge in edges:
                # GetEdges2 returns transient Edge geometry in model space. Normal
                # sketch segments are sketch-local and require sketch_to_model, but
                # applying that transform to text edges again moves rotated/offset
                # sketch text away from the rest of its marking geometry.
                points = _sample_edge(edge, curve_samples)
                if len(points) >= 2:
                    text_paths.append(tuple(points))
        return SolidWorksMarkingPaths(tuple(line_paths), tuple(text_paths))

    def _to_model_points(self, points, matrix) -> list[Vector3]:
        try:
            return [_transform_point(_vec3(point), matrix) for point in points]
        except Exception as exc:
            raise SolidWorksExportError(
                f"Could not transform marking sketch coordinates to model space: {exc}"
            ) from exc


def project_marking_paths(
    paths_model_m: Iterable[Sequence[Sequence[float]]],
    frame: PlaneFrame,
    scale: float,
) -> list[list[tuple[float, float]]]:
    """Orthographically flatten marking geometry onto the selected cut face."""

    source_paths = [tuple(path) for path in paths_model_m]
    projected: list[list[tuple[float, float]]] = []
    for path in source_paths:
        output_path: list[tuple[float, float]] = []
        for point_value in path:
            point = _vec3(point_value)
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


def _activate_document(app, model) -> None:
    """Make the model active before selecting its face for native DXF export."""

    try:
        active = _com_value(app, "ActiveDoc")
    except Exception:
        active = None
    if _same_document(active, model):
        return

    try:
        title = str(_com_value(model, "GetTitle"))
    except Exception as exc:
        raise SolidWorksExportError(
            f"Could not identify the SolidWorks document before export: {exc}"
        ) from exc

    failures: list[str] = []
    for candidate in (app, _dynamic_dispatch(app)):
        try:
            errors = _int32_byref()
            activated = candidate.ActivateDoc3(title, False, 0, errors)
            if activated is not None:
                return
            failures.append(
                f"ActivateDoc3 returned no document (error {_variant_value(errors)})"
            )
        except Exception as exc:
            failures.append(f"ActivateDoc3: {type(exc).__name__}: {exc}")

    for candidate in (app, _dynamic_dispatch(app)):
        try:
            errors = _int32_byref()
            activated = candidate.ActivateDoc2(title, False, errors)
            if activated is not None:
                return
            failures.append(
                f"ActivateDoc2 returned no document (error {_variant_value(errors)})"
            )
        except Exception as exc:
            failures.append(f"ActivateDoc2: {type(exc).__name__}: {exc}")

    detail = "; ".join(failures[-3:])
    raise SolidWorksExportError(
        f"Could not activate SolidWorks document '{title}' before export. Details: {detail}"
    )


def _same_document(first, second) -> bool:
    if first is None or second is None:
        return False
    if first is second:
        return True
    for member in ("GetPathName", "GetTitle"):
        try:
            first_value = str(_com_value(first, member)).strip().casefold()
            second_value = str(_com_value(second, member)).strip().casefold()
            if first_value and first_value == second_value:
                return True
        except Exception:
            continue
    return False


def _int32_byref():
    try:
        import pythoncom
        from win32com.client import VARIANT

        return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
    except ImportError:
        return 0


def _variant_value(value):
    return getattr(value, "value", value)


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
    candidates = [face, _dynamic_dispatch(face)]
    failures: list[str] = []
    try:
        import win32com.client

        candidates.append(win32com.client.CastTo(face, "IEntity"))
    except Exception as exc:
        failures.append(f"IEntity cast: {type(exc).__name__}: {exc}")

    for candidate in candidates:
        try:
            if bool(candidate.Select4(False, None)):
                return
            failures.append("Select4 returned False")
        except Exception as exc:
            failures.append(f"Select4: {type(exc).__name__}: {exc}")
        try:
            if bool(candidate.Select2(False, 0)):
                return
            failures.append("Select2 returned False")
        except Exception as exc:
            failures.append(f"Select2: {type(exc).__name__}: {exc}")

    detail = "; ".join(failures[-4:])
    raise SolidWorksExportError(
        f"Could not select the export face. SolidWorks details: {detail}"
    )


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


def _sketch_to_model_matrix(sketch) -> tuple[float, ...]:
    """Read and invert a sketch's model-to-sketch transform defensively."""

    failures: list[str] = []
    try:
        transform = _com_value(sketch, "ModelToSketchTransform")
        model_to_sketch = _math_transform_array(transform)
        return _invert_transform_matrix(model_to_sketch)
    except Exception as exc:
        failures.append(f"ModelToSketchTransform: {type(exc).__name__}: {exc}")

    # Older SolidWorks APIs return the same matrix as a raw array. Keeping
    # this fallback mirrors cad-batch-converter's late-bound, version-tolerant
    # COM strategy and avoids relying on a returned IMathTransform wrapper.
    for member in ("ModelToSketchXform", "IModelToSketchXform"):
        try:
            values = tuple(float(value) for value in _com_value(sketch, member))
            return _invert_transform_matrix(values)
        except Exception as exc:
            failures.append(f"{member}: {type(exc).__name__}: {exc}")

    raise SolidWorksExportError(
        "SolidWorks did not expose a readable sketch transform. "
        + "; ".join(failures[-3:])
    )


def _math_transform_array(transform) -> tuple[float, ...]:
    failures: list[str] = []
    candidates: list[tuple[str, object]] = [("direct", transform)]
    dynamic = _dynamic_dispatch(transform)
    if dynamic is not transform:
        candidates.append(("dynamic", dynamic))
    try:
        import win32com.client

        candidates.append(
            ("IMathTransform cast", win32com.client.CastTo(transform, "IMathTransform"))
        )
    except Exception as exc:
        failures.append(f"IMathTransform cast: {type(exc).__name__}: {exc}")

    for label, candidate in candidates:
        for member in ("ArrayData", "IArrayData"):
            try:
                return tuple(
                    float(value) for value in _com_value(candidate, member)
                )
            except Exception as exc:
                failures.append(
                    f"{label}.{member}: {type(exc).__name__}: {exc}"
                )
    raise SolidWorksExportError(
        "SolidWorks returned a math-transform object without readable array data. "
        + "; ".join(failures[-4:])
    )


def _sketch_text_edges(segment) -> list:
    """Read ISketchText edges through whichever COM interface is available."""

    failures: list[str] = []
    candidates: list[tuple[str, object]] = [("direct", segment)]
    dynamic = _dynamic_dispatch(segment)
    if dynamic is not segment:
        candidates.append(("dynamic", dynamic))
    try:
        import win32com.client

        candidates.append(
            ("ISketchText cast", win32com.client.CastTo(segment, "ISketchText"))
        )
    except Exception as exc:
        failures.append(f"ISketchText cast: {type(exc).__name__}: {exc}")

    for label, candidate in candidates:
        for member in ("GetEdges2", "GetEdges"):
            try:
                edges = _as_sequence(_com_value(candidate, member))
                if edges:
                    return edges
            except Exception as exc:
                failures.append(
                    f"{label}.{member}: {type(exc).__name__}: {exc}"
                )
    if failures:
        raise SolidWorksExportError(
            "SolidWorks did not expose rendered sketch-text edges. "
            + "; ".join(failures[-4:])
        )
    return []


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


def _transform_point(point: Vector3, matrix: Sequence[float]) -> Vector3:
    """Apply a SolidWorks IMathTransform ArrayData matrix to a point.

    SolidWorks stores the 3x3 rotation in elements 0..8, translation in
    9..11, and a uniform scale in element 12. Points are rotated, scaled,
    then translated.
    """

    if len(matrix) < 13:
        raise SolidWorksExportError(
            f"SolidWorks returned {len(matrix)} transform values; expected 16."
        )
    scale_factor = float(matrix[12])
    if not math.isfinite(scale_factor) or abs(scale_factor) <= 1e-15:
        raise SolidWorksExportError(
            f"SolidWorks returned an invalid sketch-transform scale: {scale_factor}."
        )
    x, y, z = point
    return (
        scale_factor * (x * matrix[0] + y * matrix[3] + z * matrix[6])
        + matrix[9],
        scale_factor * (x * matrix[1] + y * matrix[4] + z * matrix[7])
        + matrix[10],
        scale_factor * (x * matrix[2] + y * matrix[5] + z * matrix[8])
        + matrix[11],
    )


def _invert_transform_matrix(matrix: Sequence[float]) -> tuple[float, ...]:
    """Invert a SolidWorks affine transform without COM math objects."""

    if len(matrix) < 13:
        raise SolidWorksExportError(
            f"SolidWorks returned {len(matrix)} transform values; expected 16."
        )
    scale_factor = float(matrix[12])
    if not math.isfinite(scale_factor) or abs(scale_factor) <= 1e-15:
        raise SolidWorksExportError(
            f"SolidWorks returned an invalid sketch-transform scale: {scale_factor}."
        )

    a00, a01, a02 = (scale_factor * float(matrix[i]) for i in (0, 1, 2))
    a10, a11, a12 = (scale_factor * float(matrix[i]) for i in (3, 4, 5))
    a20, a21, a22 = (scale_factor * float(matrix[i]) for i in (6, 7, 8))
    determinant = (
        a00 * (a11 * a22 - a12 * a21)
        - a01 * (a10 * a22 - a12 * a20)
        + a02 * (a10 * a21 - a11 * a20)
    )
    if not math.isfinite(determinant) or abs(determinant) <= 1e-15:
        raise SolidWorksExportError("SolidWorks returned a singular sketch transform.")

    inv00 = (a11 * a22 - a12 * a21) / determinant
    inv01 = (a02 * a21 - a01 * a22) / determinant
    inv02 = (a01 * a12 - a02 * a11) / determinant
    inv10 = (a12 * a20 - a10 * a22) / determinant
    inv11 = (a00 * a22 - a02 * a20) / determinant
    inv12 = (a02 * a10 - a00 * a12) / determinant
    inv20 = (a10 * a21 - a11 * a20) / determinant
    inv21 = (a01 * a20 - a00 * a21) / determinant
    inv22 = (a00 * a11 - a01 * a10) / determinant

    tx, ty, tz = (float(matrix[i]) for i in (9, 10, 11))
    inverse_tx = -(tx * inv00 + ty * inv10 + tz * inv20)
    inverse_ty = -(tx * inv01 + ty * inv11 + tz * inv21)
    inverse_tz = -(tx * inv02 + ty * inv12 + tz * inv22)
    return (
        inv00, inv01, inv02,
        inv10, inv11, inv12,
        inv20, inv21, inv22,
        inverse_tx, inverse_ty, inverse_tz,
        1.0, 0.0, 0.0, 0.0,
    )


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
