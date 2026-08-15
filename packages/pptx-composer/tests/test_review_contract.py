import unittest

from pptx_ai_composer.review_contract import ReviewContractError, validate_review_packet


class ReviewContractTests(unittest.TestCase):
    def valid_packet(self):
        digest = "a" * 64
        return {
            "contract_version": "1.0",
            "kind": "pptx_chat_review",
            "interaction": "chat_only",
            "deck_id": "deck_demo",
            "revision": digest,
            "fidelity": "structural_preview_not_powerpoint_render",
            "limitations": ["Approximate font metrics"],
            "diagnostics": {"text_overflow": []},
            "slides": [{
                "slide_id": "s_cover",
                "number": 1,
                "png_file": "slide-01.png",
                "png_sha256": digest,
                "svg_file": "slide-01.svg",
                "svg_sha256": digest,
            }],
        }

    def test_accepts_closed_packet_and_returns_defensive_copy(self):
        packet = self.valid_packet()
        validated = validate_review_packet(packet)

        self.assertEqual(validated, packet)
        self.assertIsNot(validated, packet)
        validated["slides"][0]["slide_id"] = "changed"
        self.assertEqual(packet["slides"][0]["slide_id"], "s_cover")

    def test_refuses_unknown_top_level_or_nested_fields(self):
        top_level = self.valid_packet()
        top_level["prompt"] = "change the deck"
        with self.assertRaisesRegex(ReviewContractError, "unknown field: prompt"):
            validate_review_packet(top_level)

        nested = self.valid_packet()
        nested["slides"][0]["x"] = 10
        with self.assertRaisesRegex(ReviewContractError, "unknown field: x"):
            validate_review_packet(nested)

    def test_refuses_non_chat_interaction_and_non_local_artifacts(self):
        editable = self.valid_packet()
        editable["interaction"] = "direct_edit"
        with self.assertRaisesRegex(ReviewContractError, "chat_only"):
            validate_review_packet(editable)

        escaped = self.valid_packet()
        escaped["slides"][0]["png_file"] = "../slide.png"
        with self.assertRaisesRegex(ReviewContractError, "local .png basename"):
            validate_review_packet(escaped)

    def test_refuses_non_integer_slide_numbers(self):
        for invalid in (True, 1.0):
            with self.subTest(invalid=invalid):
                packet = self.valid_packet()
                packet["slides"][0]["number"] = invalid
                with self.assertRaisesRegex(ReviewContractError, "positive integer"):
                    validate_review_packet(packet)


if __name__ == "__main__":
    unittest.main()
