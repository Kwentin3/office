import copy
import hashlib
import tempfile
import unittest
from pathlib import Path

from PIL import Image

from pptx_ai_composer.compiler import CompileError, compile_deck
from pptx_ai_composer.scene_contract import validate_scene_spec
from tests.fixtures import expanded_deck_spec, valid_deck_spec


class CompositionCompilerTests(unittest.TestCase):
    def test_compiles_existing_archetypes_to_bounded_scene(self):
        deck = valid_deck_spec()
        deck["slides"][0]["source_ids"] = ["src_1"]
        scene = compile_deck(deck)
        validate_scene_spec(scene)
        self.assertEqual([slide["archetype"] for slide in scene["slides"]], ["cover", "comparison", "chart_with_takeaway"])
        self.assertEqual(scene["slides"][1]["variant"], "balanced")
        comparison_roles = {node["role"] for node in scene["slides"][1]["nodes"]}
        self.assertIn("comparison.left.card", comparison_roles)
        self.assertIn("comparison.right.items", comparison_roles)
        chart_nodes = [node for node in scene["slides"][2]["nodes"] if node["kind"] == "chart"]
        self.assertEqual(len(chart_nodes), 1)
        self.assertEqual(chart_nodes[0]["chart"], deck["slides"][2]["chart"])
        for slide in scene["slides"]:
            self.assertEqual(len([node for node in slide["nodes"] if node["role"] == "source.footer"]), 1)

    def test_compiler_returns_defensive_scene(self):
        scene = compile_deck(valid_deck_spec())
        scene["slides"][0]["nodes"][0]["fill"] = "000000"
        self.assertNotEqual(compile_deck(valid_deck_spec())["slides"][0]["nodes"][0]["fill"], "000000")

    def test_rejects_unknown_variant(self):
        with self.assertRaisesRegex(CompileError, "unsupported variant"):
            compile_deck(valid_deck_spec(), variants={"comparison": "freeform"})

    def test_can_compile_one_slide_for_local_iteration(self):
        scene = compile_deck(valid_deck_spec(), slide_ids=["s_compare"])
        self.assertEqual([slide["slide_id"] for slide in scene["slides"]], ["s_compare"])

    def test_rejects_unknown_slide_selection(self):
        with self.assertRaisesRegex(CompileError, "unknown slide_id"):
            compile_deck(valid_deck_spec(), slide_ids=["missing"])

    def test_rejects_string_instead_of_slide_id_array(self):
        with self.assertRaisesRegex(CompileError, "slide_ids must be an array"):
            compile_deck(valid_deck_spec(), slide_ids="s_compare")

    def test_rejects_unknown_variant_target_instead_of_ignoring_typo(self):
        with self.assertRaisesRegex(CompileError, "unknown variant target"):
            compile_deck(valid_deck_spec(), variants={"s_comapre": "compact"})

    def test_rejects_non_object_variants(self):
        with self.assertRaisesRegex(CompileError, "variants must be an object"):
            compile_deck(valid_deck_spec(), variants=[("s_compare", "compact")])

    def test_each_declared_alternative_variant_changes_scene(self):
        from pptx_ai_composer.library import get_catalog
        catalog = get_catalog()
        for deck in (valid_deck_spec(), expanded_deck_spec()):
            for slide in deck["slides"]:
                variants = catalog["archetypes"][slide["archetype"]]["variants"]
                balanced = compile_deck(deck, slide_ids=[slide["slide_id"]])
                for variant in variants:
                    if variant == "balanced":
                        continue
                    alternative = compile_deck(deck, slide_ids=[slide["slide_id"]], variants={slide["slide_id"]: variant})
                    self.assertNotEqual(balanced["slides"][0]["nodes"], alternative["slides"][0]["nodes"], f"variant {slide['archetype']}:{variant} is a no-op")

    def test_cover_png_asset_compiles_for_every_declared_variant(self):
        from pptx_ai_composer.library import get_catalog
        with tempfile.TemporaryDirectory() as directory:
            image = Path(directory) / "hero.png"
            Image.new("RGB", (40, 40), "red").save(image)
            deck = valid_deck_spec()
            deck["assets"] = [{
                "asset_id": "hero", "kind": "png", "path": str(image),
                "sha256": hashlib.sha256(image.read_bytes()).hexdigest(), "alt_text": "Hero",
            }]
            deck["slides"][0]["asset_id"] = "hero"
            for variant in get_catalog()["archetypes"]["cover"]["variants"]:
                with self.subTest(variant=variant):
                    scene = compile_deck(deck, slide_ids=["s_cover"], variants={"s_cover": variant})
                    nodes = {node["node_id"]: node for node in scene["slides"][0]["nodes"]}
                    self.assertEqual(nodes["cover-image"]["asset_id"], "hero")
                    self.assertIn("title", nodes)
                    self.assertIn("subtitle", nodes)
                    self.assertNotIn("cover-motif", nodes)
                    if variant == "dark":
                        self.assertEqual(nodes["title"]["style"]["color"], deck["brand"]["colors"]["surface"])

    def test_new_archetypes_compile_as_recipes_over_existing_primitives(self):
        deck = valid_deck_spec()
        deck["slides"] = [
            {"slide_id": "process", "archetype": "process", "title": "Процесс", "steps": [{"label": "Сбор", "description": "Собрать материалы"}, {"label": "Проверка", "description": "Проверить основания"}, {"label": "Выпуск", "description": "Утвердить результат"}], "source_ids": ["src_1"]},
            {"slide_id": "timeline", "archetype": "timeline", "title": "План", "milestones": [{"period": "Неделя 1", "label": "Подготовка"}, {"period": "Неделя 2", "label": "Первый цикл"}, {"period": "Неделя 3", "label": "Второй цикл"}], "source_ids": ["src_1"]},
            {"slide_id": "decision", "archetype": "decision_matrix", "title": "Решение", "criteria": ["Эффект", "Контроль"], "options": [{"label": "Продолжить", "ratings": ["positive", "positive"]}, {"label": "Изменить", "ratings": ["neutral", "positive"]}], "source_ids": ["src_1"]},
            {"slide_id": "kpis", "archetype": "kpi_grid", "title": "Метрики", "metrics": [{"label": "Время", "value": "Измерить", "note": "До и после"}, {"label": "Доработки", "value": "Считать", "note": "Каждый цикл"}, {"label": "Контроль", "value": "Проверять", "note": "Перед выпуском"}], "source_ids": ["src_1"]},
        ]
        scene = compile_deck(deck)
        validate_scene_spec(scene)
        self.assertEqual([slide["archetype"] for slide in scene["slides"]], ["process", "timeline", "decision_matrix", "kpi_grid"])
        for slide in scene["slides"]:
            self.assertTrue({node["kind"] for node in slide["nodes"]} <= {"text", "shape", "image", "chart"})
        self.assertIn("process.connector", {node["role"] for node in scene["slides"][0]["nodes"]})
        self.assertIn("timeline.card", {node["role"] for node in scene["slides"][1]["nodes"]})
        timeline_cards = [node for node in scene["slides"][1]["nodes"] if node["role"] == "timeline.card"]
        self.assertTrue(all(node["box"]["x"] >= 40 and node["box"]["x"] + node["box"]["w"] <= 1240 for node in timeline_cards))
        self.assertIn("decision.rating", {node["role"] for node in scene["slides"][2]["nodes"]})


if __name__ == "__main__":
    unittest.main()
