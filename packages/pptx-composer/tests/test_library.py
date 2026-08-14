import unittest

from pptx_ai_composer.library import LibraryError, get_catalog, validate_catalog


class ManagedLibraryTests(unittest.TestCase):
    def test_catalog_exposes_archetype_component_composition(self):
        catalog = get_catalog()
        self.assertEqual(catalog["catalog_version"], "1.0")
        self.assertEqual(
            set(catalog["archetypes"]),
            {"cover", "comparison", "chart_with_takeaway", "process", "timeline", "decision_matrix", "kpi_grid"},
        )
        self.assertIn("text", catalog["components"])
        self.assertIn("shape", catalog["components"])
        self.assertEqual(catalog["archetypes"]["comparison"]["status"], "stable")
        self.assertEqual(catalog["archetypes"]["comparison"]["variants"], ["balanced", "compact"])
        self.assertEqual(catalog["archetypes"]["timeline"]["variants"], ["balanced"])
        self.assertEqual(
            catalog["archetypes"]["comparison"]["components"],
            ["background", "title", "two_column_cards", "source_footer"],
        )
        self.assertIn("source_footer", catalog["archetypes"]["cover"]["components"])

    def test_catalog_is_a_defensive_copy(self):
        first = get_catalog()
        first["archetypes"]["comparison"]["status"] = "broken"
        self.assertEqual(get_catalog()["archetypes"]["comparison"]["status"], "stable")

    def test_rejects_archetype_referencing_unknown_component(self):
        catalog = get_catalog()
        catalog["archetypes"]["comparison"]["components"].append("arbitrary_python")
        with self.assertRaisesRegex(LibraryError, "unknown component"):
            validate_catalog(catalog)

    def test_rejects_unbounded_component(self):
        catalog = get_catalog()
        catalog["components"]["raw_coordinates"] = {
            "kind": "execution",
            "status": "stable",
            "bounded": False,
        }
        with self.assertRaisesRegex(LibraryError, "must be bounded"):
            validate_catalog(catalog)


if __name__ == "__main__":
    unittest.main()
