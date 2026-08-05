from __future__ import annotations

from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from solidworks_adapter import (  # noqa: E402
    PlaneFrame,
    SolidWorksExportError,
    SolidWorksSession,
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


class _Sketch:
    @property
    def ModelToSketchTransform(self):
        raise RuntimeError("No transform needed in this unit test")

    def GetSketchSegments(self):
        return [_SketchText()]


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


class _Body:
    def GetFaces(self):
        return [_PlanarFace()]


class _PartModel:
    def GetBodies2(self, body_type, visible_only):
        return [_Body()]


class SolidWorksAdapterTests(unittest.TestCase):
    def test_finds_planar_face_without_surface_interface_cast(self):
        session = SolidWorksSession(object(), started_by_tool=False)

        result = session.find_export_face(_PartModel())

        self.assertEqual(result.outer_loops, 1)
        self.assertEqual(result.inner_loops, 1)
        self.assertEqual(min(point[0] for point in result.projected_points_m), 0.0)
        self.assertEqual(min(point[1] for point in result.projected_points_m), 0.0)

    def test_reads_text_from_normal_sketch_segment_enumeration(self):
        session = SolidWorksSession(object(), started_by_tool=False)

        paths = session.marking_paths_model_space(
            _Model(), "CUTFILE MARKING", curve_samples=4
        )

        self.assertEqual(len(paths), 1)
        self.assertEqual(paths[0][0], (0.0, 0.0, 0.0))
        self.assertEqual(paths[0][-1], (1.0, 2.0, 0.0))

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
