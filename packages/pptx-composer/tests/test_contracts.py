import math
import unittest

from pptx_ai_composer.contracts import ContractError, validate_deck_spec
from tests.fixtures import valid_deck_spec


class DeckSpecContractTests(unittest.TestCase):
    def valid_spec(self):
        return valid_deck_spec()

    def test_accepts_closed_minimal_vertical_slice(self):
        normalized = validate_deck_spec(self.valid_spec())
        self.assertEqual(normalized["deck_id"], "deck_demo")
        self.assertEqual([s["archetype"] for s in normalized["slides"]], ["cover", "comparison", "chart_with_takeaway"])

    def test_rejects_unknown_field_in_nested_slide(self):
        spec = self.valid_spec()
        spec["slides"][1]["raw_coordinates"] = {"x": 1, "y": 2}
        with self.assertRaisesRegex(ContractError, "unknown field"):
            validate_deck_spec(spec)

    def test_rejects_duplicate_stable_ids(self):
        spec = self.valid_spec()
        spec["slides"][1]["slide_id"] = "s_cover"
        with self.assertRaisesRegex(ContractError, "duplicate slide_id"):
            validate_deck_spec(spec)

    def test_rejects_chart_shape_mismatch(self):
        spec = self.valid_spec()
        spec["slides"][2]["chart"]["series"][0]["values"] = [8.0]
        with self.assertRaisesRegex(ContractError, "categories and values"):
            validate_deck_spec(spec)

    def test_rejects_unknown_source_reference(self):
        deck = valid_deck_spec()
        deck["slides"][1]["source_ids"] = ["missing"]
        with self.assertRaisesRegex(ContractError, "unknown source_id"):
            validate_deck_spec(deck)

    def test_accepts_bounded_process_timeline_decision_and_kpi_archetypes(self):
        deck = valid_deck_spec()
        deck["slides"] = [
            {
                "slide_id": "process",
                "archetype": "process",
                "title": "Контролируемый процесс",
                "steps": [
                    {"label": "Сбор", "description": "Получить утверждённые материалы."},
                    {"label": "Проверка", "description": "Показать основания и открытые вопросы."},
                    {"label": "Выпуск", "description": "Передать человеку на утверждение."},
                ],
                "source_ids": ["src_1"],
            },
            {
                "slide_id": "timeline",
                "archetype": "timeline",
                "title": "Четыре недели пилота",
                "milestones": [
                    {"period": "Неделя 1", "label": "Исходные значения"},
                    {"period": "Неделя 2", "label": "Первый цикл"},
                    {"period": "Неделя 3", "label": "Повторение"},
                    {"period": "Неделя 4", "label": "Решение"},
                ],
                "source_ids": ["src_1"],
            },
            {
                "slide_id": "decision",
                "archetype": "decision_matrix",
                "title": "Решение после пилота",
                "criteria": ["Измеримый эффект", "Контроль доступа", "Приемлемая нагрузка"],
                "options": [
                    {"label": "Продолжить", "ratings": ["positive", "positive", "positive"]},
                    {"label": "Скорректировать", "ratings": ["neutral", "positive", "neutral"]},
                    {"label": "Остановить", "ratings": ["negative", "negative", "negative"]},
                ],
                "source_ids": ["src_1"],
            },
            {
                "slide_id": "kpis",
                "archetype": "kpi_grid",
                "title": "Как измеряем",
                "metrics": [
                    {"label": "Время", "value": "Измерить", "note": "До и во время пилота"},
                    {"label": "Доработки", "value": "Считать", "note": "На каждом цикле"},
                    {"label": "Подтверждения", "value": "Проверять", "note": "Перед выпуском"},
                ],
                "source_ids": ["src_1"],
            },
        ]
        self.assertEqual(len(validate_deck_spec(deck)["slides"]), 4)

    def test_rejects_ragged_decision_matrix(self):
        deck = valid_deck_spec()
        deck["slides"] = [{
            "slide_id": "decision", "archetype": "decision_matrix", "title": "Decision",
            "criteria": ["Effect", "Control"],
            "options": [
                {"label": "Go", "ratings": ["positive"]},
                {"label": "Stop", "ratings": ["negative", "negative"]},
            ],
            "source_ids": ["src_1"],
        }]
        with self.assertRaisesRegex(ContractError, "ratings and criteria"):
            validate_deck_spec(deck)

    def test_rejects_non_finite_chart_values(self):
        for value in (math.nan, math.inf, -math.inf):
            deck = valid_deck_spec()
            deck["slides"][2]["chart"]["series"][0]["values"][0] = value
            with self.subTest(value=value), self.assertRaisesRegex(ContractError, "finite"):
                validate_deck_spec(deck)

    def test_rejects_source_footer_over_budget_before_compilation(self):
        deck = valid_deck_spec()
        deck["sources"] = [
            {"source_id": f"src_{index}", "label": "Источник с очень длинным пользовательским названием " + str(index)}
            for index in range(8)
        ]
        deck["slides"][1]["source_ids"] = [source["source_id"] for source in deck["sources"]]
        with self.assertRaisesRegex(ContractError, "source footer exceeds"):
            validate_deck_spec(deck)

    def test_accepts_equal_verified_chart_values(self):
        deck = valid_deck_spec()
        deck["slides"][2]["chart"]["series"][0]["values"] = [1, 1, 1]
        self.assertEqual(validate_deck_spec(deck)["slides"][2]["chart"]["data_source_id"], "src_1")

    def test_rejects_chart_without_dataset_source(self):
        deck = valid_deck_spec()
        del deck["slides"][2]["chart"]["data_source_id"]
        with self.assertRaisesRegex(ContractError, "data_source_id"):
            validate_deck_spec(deck)

    def test_rejects_chart_with_unknown_dataset_source(self):
        deck = valid_deck_spec()
        deck["slides"][2]["chart"]["data_source_id"] = "missing"
        with self.assertRaisesRegex(ContractError, "unknown data_source_id"):
            validate_deck_spec(deck)


if __name__ == "__main__":
    unittest.main()
