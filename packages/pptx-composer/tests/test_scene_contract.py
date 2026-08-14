import copy
import math
import unittest

from pptx_ai_composer.scene_contract import SceneContractError, validate_scene_spec


def valid_scene_spec():
    return {
        "scene_version": "1.0",
        "deck_id": "demo",
        "canvas": {"width": 1280, "height": 720},
        "slides": [
            {
                "slide_id": "slide-1",
                "archetype": "comparison",
                "variant": "balanced",
                "nodes": [
                    {
                        "node_id": "background",
                        "kind": "shape",
                        "role": "background",
                        "box": {"x": 0, "y": 0, "w": 1280, "h": 720},
                        "shape": "rect",
                        "fill": "F7F1E8",
                    },
                    {
                        "node_id": "title",
                        "kind": "text",
                        "role": "title",
                        "box": {"x": 70, "y": 48, "w": 1140, "h": 72},
                        "text": "Проверяемый заголовок",
                        "style": {"font": "Inter", "size": 34, "color": "2B2926", "bold": True, "align": "left", "valign": "top"},
                    },
                ],
            }
        ],
    }


class SceneContractTests(unittest.TestCase):
    def test_accepts_closed_bounded_scene(self):
        scene = validate_scene_spec(valid_scene_spec())
        self.assertEqual(scene["canvas"], {"width": 1280, "height": 720})
        self.assertEqual(scene["slides"][0]["nodes"][1]["kind"], "text")

    def test_rejects_unknown_node_field(self):
        scene = valid_scene_spec()
        scene["slides"][0]["nodes"][0]["ooxml"] = "<p:sp/>"
        with self.assertRaisesRegex(SceneContractError, "unknown field"):
            validate_scene_spec(scene)

    def test_rejects_off_canvas_box(self):
        scene = valid_scene_spec()
        scene["slides"][0]["nodes"][1]["box"]["x"] = 1270
        with self.assertRaisesRegex(SceneContractError, "outside canvas"):
            validate_scene_spec(scene)

    def test_rejects_excessive_nodes(self):
        scene = valid_scene_spec()
        node = scene["slides"][0]["nodes"][0]
        scene["slides"][0]["nodes"] = [dict(copy.deepcopy(node), node_id=f"node-{index}") for index in range(81)]
        with self.assertRaisesRegex(SceneContractError, "cardinality"):
            validate_scene_spec(scene)

    def test_rejects_duplicate_node_ids(self):
        scene = valid_scene_spec()
        scene["slides"][0]["nodes"].append(copy.deepcopy(scene["slides"][0]["nodes"][1]))
        with self.assertRaisesRegex(SceneContractError, "duplicate node_id"):
            validate_scene_spec(scene)

    def test_rejects_unbounded_nesting_or_callbacks_by_closed_schema(self):
        scene = valid_scene_spec()
        scene["slides"][0]["nodes"][0]["children"] = []
        with self.assertRaisesRegex(SceneContractError, "unknown field"):
            validate_scene_spec(scene)

    def test_rejects_non_finite_chart_values_at_runtime_boundary(self):
        scene = valid_scene_spec()
        scene["slides"][0]["nodes"].append({
            "node_id": "chart", "kind": "chart", "role": "chart.primary",
            "box": {"x": 70, "y": 160, "w": 700, "h": 400},
            "chart": {"type": "bar", "data_source_id": "src_1", "categories": ["A", "B"], "series": [{"name": "Value", "values": [1, 2]}]},
        })
        for value in (math.nan, math.inf, -math.inf):
            candidate = copy.deepcopy(scene)
            candidate["slides"][0]["nodes"][2]["chart"]["series"][0]["values"][0] = value
            with self.subTest(value=value), self.assertRaisesRegex(SceneContractError, "finite"):
                validate_scene_spec(candidate)


if __name__ == "__main__":
    unittest.main()
