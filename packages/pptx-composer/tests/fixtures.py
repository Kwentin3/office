def valid_deck_spec():
    return {
        "contract_version": "1.0",
        "deck_id": "deck_demo",
        "title": "Hermes pilot",
        "brand": {
            "name": "Executive Sand",
            "colors": {
                "background": "F5F1E8",
                "surface": "FFFFFF",
                "text": "17212B",
                "muted": "62707D",
                "accent": "D95D39",
                "accent_secondary": "2A7F72",
            },
            "fonts": {"heading": "Aptos Display", "body": "Aptos"},
        },
        "sources": [{"source_id": "src_1", "label": "Pilot workbook"}],
        "assets": [],
        "slides": [
            {
                "slide_id": "s_cover",
                "archetype": "cover",
                "title": "Hermes pilot",
                "subtitle": "A bounded four-week experiment",
                "source_ids": [],
            },
            {
                "slide_id": "s_compare",
                "archetype": "comparison",
                "title": "From fragmented work to a controlled workflow",
                "left": {"label": "Today", "items": ["Five tools", "Manual transfer"]},
                "right": {"label": "With Hermes", "items": ["One workflow", "Auditable actions"]},
                "source_ids": ["src_1"],
            },
            {
                "slide_id": "s_chart",
                "archetype": "chart_with_takeaway",
                "title": "Manual work is concentrated in hand-offs",
                "chart": {
                    "type": "bar",
                    "data_source_id": "src_1",
                    "categories": ["CRM", "Documents", "Approval"],
                    "series": [{"name": "Hours", "values": [8.0, 7.4, 6.0]}],
                },
                "takeaway": "21.4 hours per week can be addressed in one bounded pilot.",
                "source_ids": ["src_1"],
            },
        ],
    }


def expanded_deck_spec():
    deck = valid_deck_spec()
    deck["slides"] = [
        {
            "slide_id": "process", "archetype": "process", "title": "Контролируемый процесс",
            "steps": [
                {"label": "Сбор", "description": "Получить утверждённые материалы."},
                {"label": "Проверка", "description": "Показать основания и открытые вопросы."},
                {"label": "Выпуск", "description": "Передать человеку на утверждение."},
            ], "source_ids": ["src_1"],
        },
        {
            "slide_id": "timeline", "archetype": "timeline", "title": "Четыре недели пилота",
            "milestones": [
                {"period": "Неделя 1", "label": "Исходные значения"},
                {"period": "Неделя 2", "label": "Первый цикл"},
                {"period": "Неделя 3", "label": "Повторение"},
                {"period": "Неделя 4", "label": "Решение"},
            ], "source_ids": ["src_1"],
        },
        {
            "slide_id": "decision", "archetype": "decision_matrix", "title": "Решение после пилота",
            "criteria": ["Измеримый эффект", "Контроль доступа", "Приемлемая нагрузка"],
            "options": [
                {"label": "Продолжить", "ratings": ["positive", "positive", "positive"]},
                {"label": "Скорректировать", "ratings": ["neutral", "positive", "neutral"]},
                {"label": "Остановить", "ratings": ["negative", "negative", "negative"]},
            ], "source_ids": ["src_1"],
        },
        {
            "slide_id": "kpis", "archetype": "kpi_grid", "title": "Как измеряем",
            "metrics": [
                {"label": "Время", "value": "Измерить", "note": "До и во время пилота"},
                {"label": "Доработки", "value": "Считать", "note": "На каждом цикле"},
                {"label": "Подтверждения", "value": "Проверять", "note": "Перед выпуском"},
            ], "source_ids": ["src_1"],
        },
    ]
    return deck
