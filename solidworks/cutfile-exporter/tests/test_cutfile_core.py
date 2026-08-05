from __future__ import annotations

import tempfile
from pathlib import Path
import sys
import unittest

import ezdxf


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from cutfile_core import (  # noqa: E402
    CutfileValidationError,
    INSIDE_LAYER,
    MARKING_LAYER,
    OUTSIDE_LAYER,
    add_marking_paths,
    assign_cut_layers,
    infer_model_to_dxf_scale,
    save_dxf,
)


class CutfileCoreTests(unittest.TestCase):
    def make_plate(self):
        doc = ezdxf.new("R2013")
        msp = doc.modelspace()
        msp.add_lwpolyline([(0, 0), (10, 0), (10, 6), (0, 6)], close=True)
        msp.add_circle((2, 2), radius=0.5)
        msp.add_circle((8, 2), radius=0.75)
        return doc

    def test_assigns_outside_and_inside_layers(self):
        doc = self.make_plate()
        result = assign_cut_layers(doc, expected_outer_loops=1, expected_inner_loops=2)
        self.assertEqual(result.outside_loops, 1)
        self.assertEqual(result.inside_loops, 2)
        self.assertEqual(len(doc.modelspace().query(f'*[layer=="{OUTSIDE_LAYER}"]')), 1)
        self.assertEqual(len(doc.modelspace().query(f'*[layer=="{INSIDE_LAYER}"]')), 2)

    def test_rejects_topology_mismatch(self):
        doc = self.make_plate()
        with self.assertRaisesRegex(CutfileValidationError, "topology mismatch"):
            assign_cut_layers(doc, expected_outer_loops=1, expected_inner_loops=3)

    def test_rejects_open_cut_geometry(self):
        doc = ezdxf.new("R2013")
        doc.modelspace().add_line((0, 0), (10, 0))
        with self.assertRaises(CutfileValidationError):
            assign_cut_layers(doc, expected_outer_loops=1, expected_inner_loops=0)

    def test_assigns_loose_line_chain_and_closed_circle(self):
        doc = ezdxf.new("R2013")
        msp = doc.modelspace()
        msp.add_line((0, 0), (10, 0))
        msp.add_line((10, 0), (10, 6))
        msp.add_line((10, 6), (0, 6))
        msp.add_line((0, 6), (0, 0))
        msp.add_circle((5, 3), radius=1)

        result = assign_cut_layers(doc, expected_outer_loops=1, expected_inner_loops=1)

        self.assertEqual(result.cut_entities, 5)
        self.assertEqual(len(msp.query(f'*[layer=="{OUTSIDE_LAYER}"]')), 4)
        self.assertEqual(len(msp.query(f'*[layer=="{INSIDE_LAYER}"]')), 1)

    def test_adds_all_marking_to_one_layer(self):
        doc = self.make_plate()
        assign_cut_layers(doc, expected_outer_loops=1, expected_inner_loops=2)
        result = add_marking_paths(
            doc,
            [
                [(1, 1), (3, 1)],
                [(4, 4), (5, 5), (6, 4)],
            ],
        )
        self.assertEqual(result.paths_added, 2)
        self.assertEqual(len(doc.modelspace().query(f'*[layer=="{MARKING_LAYER}"]')), 2)

    def test_infers_inches_from_meter_model_extent(self):
        doc = self.make_plate()
        assign_cut_layers(doc, expected_outer_loops=1, expected_inner_loops=2)
        model_points = [(0.0, 0.0), (0.254, 0.1524)]
        scale, units = infer_model_to_dxf_scale(doc, model_points)
        self.assertAlmostEqual(scale, 39.37007874015748)
        self.assertEqual(units, "inches")

    def test_saves_valid_dxf(self):
        doc = self.make_plate()
        assign_cut_layers(doc, expected_outer_loops=1, expected_inner_loops=2)
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "plate.dxf"
            save_dxf(doc, path)
            loaded = ezdxf.readfile(path)
            self.assertIn(OUTSIDE_LAYER, loaded.layers)
            self.assertIn(INSIDE_LAYER, loaded.layers)


if __name__ == "__main__":
    unittest.main()
