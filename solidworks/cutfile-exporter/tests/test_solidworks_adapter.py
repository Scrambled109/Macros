from __future__ import annotations

from pathlib import Path
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from solidworks_adapter import (  # noqa: E402
    PlaneFrame,
    SolidWorksExportError,
    SolidWorksSession,
    _invert_transform_matrix,
    _sketch_to_model_matrix,
    _sketch_text_edges,
    _transform_point,
    project_marking_paths,
)


class _Curve:
    def IsLine(self):
        return False

    def Evaluate2(self, parameter, derivative_count):
        return (float(parameter), 2.0 * float(parameter), 0.0)


class _CurveParams:
    UMinValue = 0.0
    UMaxValue = 1.0


class _Edge:
    def GetCurve(self):
        return _Curve()

    def GetCurveParams3(self):
        return _CurveParams()


class _SketchText:
    ConstructionGeometry = False

    def GetType(self):
        return 4

    def GetEdges2(self):
        return [_Edge()]


class _Point:
    def __init__(self, x, y, z=0.0):
        self.X = float(x)
        self.Y = float(y)
        self.Z = float(z)


class _SketchLine:
    ConstructionGeometry = False

    def GetType(self):
        return 0

    def GetStartPoint2(self):
        return _Point(3.0, 4.0)

    def GetEndPoint2(self):
        return _Point(5.0, 6.0)


class _LegacySketchText:
    def GetEdges(self):
        return [_Edge()]


class _MathTransform:
    ArrayData = (
        1.0, 0.0, 0.0,
        0.0, 1.0, 0.0,
        0.0, 0.0, 1.0,
        0.0, 0.0, -0.01,
        1.0, 0.0, 0.0, 0.0,
    )


class _UnreadableMathTransform:
    @property
    def ArrayData(self):
        raise RuntimeError("simulated unreadable returned COM interface")


class _Sketch:
    @property
    def ModelToSketchTransform(self):
        return _MathTransform()

    def GetSketchSegments(self):
        return [_SketchLine(), _SketchText()]


class _LegacyTransformSketch(_Sketch):
    @property
    def ModelToSketchTransform(self):
        return _UnreadableMathTransform()

    ModelToSketchXform = _MathTransform.ArrayData


class _Feature:
    Name = "CUTFILE MARKING"

    def GetSpecificFeature2(self):
        return _Sketch()

    def GetFirstSubFeature(self):
        return None

    def GetNextFeature(self):
        return None


class _Model:
    def FirstFeature(self):
        return _Feature()


class _Loop:
    def __init__(self, outer):
        self._outer = outer

    def IsOuter(self):
        return self._outer


class _PlanarFace:
    Normal = (0.0, 0.0, 1.0)

    def GetLoops(self):
        return [_Loop(True), _Loop(False)]

    def GetTessTriangles(self, no_conversion):
        return (
            0.0, 0.0, 0.01,
            0.1, 0.0, 0.01,
            0.1, 0.06, 0.01,
            0.0, 0.0, 0.01,
            0.1, 0.06, 0.01,
            0.0, 0.06, 0.01,
        )

    def GetArea(self):
        return 0.006


class _SplitPlanarFace(_PlanarFace):
    def __init__(self, x_min, x_max):
        self.x_min = float(x_min)
        self.x_max = float(x_max)

    def GetLoops(self):
        return [_Loop(True)]

    def GetTessTriangles(self, no_conversion):
        return (
            self.x_min, 0.0, 0.01,
            self.x_max, 0.0, 0.01,
            self.x_max, 0.06, 0.01,
            self.x_min, 0.0, 0.01,
            self.x_max, 0.06, 0.01,
            self.x_min, 0.06, 0.01,
        )

    def GetArea(self):
        return (self.x_max - self.x_min) * 0.06


class _Body:
    def GetFaces(self):
        return [_PlanarFace()]


class _PartModel:
    def GetBodies2(self, body_type, visible_only):
        return [_Body()]


class _SplitBody:
    def GetFaces(self):
        return [_SplitPlanarFace(0.0, 0.06), _SplitPlanarFace(0.06, 0.10)]


class _SplitPartModel:
    def GetBodies2(self, body_type, visible_only):
        return [_SplitBody()]


class _ActivePartWithPropertyMembers:
    _oleobj_ = object()

    def __init__(self, path):
        self.GetPathName = str(path)
        self.GetType = 1

    def __call__(self):
        raise AssertionError("A COM property object must not be called")


class _AppWithActivePart:
    def __init__(self, path):
        self.ActiveDoc = _ActivePartWithPropertyMembers(path)


class SolidWorksAdapterTests(unittest.TestCase):
    def test_reuses_active_part_when_getters_are_com_properties(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "test.SLDPRT"
            path.touch()
            session = SolidWorksSession(
                _AppWithActivePart(path), started_by_tool=False
            )

            result = session.open_part(path)

            self.assertFalse(result.opened_by_tool)
            self.assertEqual(result.path, path.resolve())

    def test_finds_planar_face_without_surface_interface_cast(self):
        session = SolidWorksSession(object(), started_by_tool=False)

        result = session.find_export_face(_PartModel())

        self.assertEqual(result.outer_loops, 1)
        self.assertEqual(result.inner_loops, 1)
        self.assertEqual(min(point[0] for point in result.projected_points_m), 0.0)
        self.assertEqual(min(point[1] for point in result.projected_points_m), 0.0)

    def test_rejects_split_coplanar_export_side(self):
        session = SolidWorksSession(object(), started_by_tool=False)

        with self.assertRaisesRegex(SolidWorksExportError, "split into multiple"):
            session.find_export_face(_SplitPartModel())

    def test_separates_line_marking_from_sketch_text(self):
        session = SolidWorksSession(object(), started_by_tool=False)

        paths = session.marking_paths_model_space(
            _Model(), "CUTFILE MARKING", curve_samples=4
        )

        self.assertEqual(len(paths.line_paths), 1)
        self.assertEqual(paths.line_paths[0][0], (3.0, 4.0, 0.01))
        self.assertEqual(paths.line_paths[0][-1], (5.0, 6.0, 0.01))
        self.assertEqual(len(paths.text_paths), 1)
        self.assertEqual(paths.text_paths[0][0], (0.0, 0.0, 0.01))
        self.assertEqual(paths.text_paths[0][-1], (1.0, 2.0, 0.01))

    def test_applies_solidworks_rotation_scale_and_translation_matrix(self):
        matrix = (
            0.0, 1.0, 0.0,
            -1.0, 0.0, 0.0,
            0.0, 0.0, 1.0,
            10.0, 20.0, 30.0,
            2.0, 0.0, 0.0, 0.0,
        )

        result = _transform_point((1.0, 0.0, 3.0), matrix)

        self.assertEqual(result, (10.0, 22.0, 36.0))

    def test_inverts_transform_matrix_without_com_math_objects(self):
        matrix = (
            0.0, 1.0, 0.0,
            -1.0, 0.0, 0.0,
            0.0, 0.0, 1.0,
            10.0, 20.0, 30.0,
            2.0, 0.0, 0.0, 0.0,
        )
        point = (1.0, 2.0, 3.0)

        transformed = _transform_point(point, matrix)
        restored = _transform_point(transformed, _invert_transform_matrix(matrix))

        for actual, expected in zip(restored, point):
            self.assertAlmostEqual(actual, expected)

    def test_uses_legacy_raw_sketch_transform_array_as_fallback(self):
        matrix = _sketch_to_model_matrix(_LegacyTransformSketch())

        self.assertEqual(_transform_point((0.0, 0.0, 0.0), matrix), (0.0, 0.0, 0.01))

    def test_reads_legacy_sketch_text_edges(self):
        self.assertEqual(len(_sketch_text_edges(_LegacySketchText())), 1)

    def test_projects_marking_on_face_plane(self):
        frame = PlaneFrame(
            origin=(0.0, 0.0, 0.01),
            x_axis=(1.0, 0.0, 0.0),
            y_axis=(0.0, 1.0, 0.0),
            normal=(0.0, 0.0, 1.0),
        )

        projected = project_marking_paths(
            [[(0.01, 0.02, 0.01), (0.03, 0.04, 0.01)]], frame, 1000.0
        )

        self.assertEqual(projected, [[(10.0, 20.0), (30.0, 40.0)]])

    def test_rejects_marking_off_face_plane(self):
        frame = PlaneFrame(
            origin=(0.0, 0.0, 0.0),
            x_axis=(1.0, 0.0, 0.0),
            y_axis=(0.0, 1.0, 0.0),
            normal=(0.0, 0.0, 1.0),
        )

        with self.assertRaisesRegex(SolidWorksExportError, "not on the exported face"):
            project_marking_paths(
                [[(0.0, 0.0, 0.001), (1.0, 0.0, 0.001)]], frame, 1.0
            )


if __name__ == "__main__":
    unittest.main()
