from __future__ import annotations

import unittest

from xlsx_artifact_tool.review_contract import ReviewContractError, validate_review_packet


class ReviewContractTests(unittest.TestCase):
    def valid_packet(self):
        digest = "a" * 64
        return {
            "contract_version": "1.0",
            "kind": "xlsx_chat_review",
            "interaction": "chat_only",
            "workbook_id": "book_demo",
            "revision": digest,
            "fidelity": "structural_preview_not_excel_render",
            "limitations": ["Formula text is displayed but never evaluated"],
            "diagnostics": {"truncated_sheets": [], "omitted_sheets": []},
            "sheets": [
                {
                    "sheet": "Summary",
                    "number": 1,
                    "html_file": "sheet-01.html",
                    "html_sha256": digest,
                }
            ],
        }

    def test_accepts_closed_packet_and_returns_defensive_copy(self):
        packet = self.valid_packet()

        validated = validate_review_packet(packet)

        self.assertEqual(validated, packet)
        self.assertIsNot(validated, packet)
        validated["sheets"][0]["sheet"] = "Changed"
        self.assertEqual(packet["sheets"][0]["sheet"], "Summary")

    def test_refuses_unknown_fields_at_every_packet_level(self):
        examples = []
        top = self.valid_packet()
        top["prompt"] = "edit this"
        examples.append(top)
        diagnostic = self.valid_packet()
        diagnostic["diagnostics"]["callback"] = "run"
        examples.append(diagnostic)
        sheet = self.valid_packet()
        sheet["sheets"][0]["office_object"] = "worksheet"
        examples.append(sheet)

        for packet in examples:
            with self.subTest(packet=packet), self.assertRaisesRegex(ReviewContractError, "unknown field"):
                validate_review_packet(packet)

    def test_refuses_non_chat_interaction_nonlocal_artifact_and_bad_digest(self):
        editable = self.valid_packet()
        editable["interaction"] = "direct_edit"
        with self.assertRaisesRegex(ReviewContractError, "chat_only"):
            validate_review_packet(editable)

        escaped = self.valid_packet()
        escaped["sheets"][0]["html_file"] = "../sheet.html"
        with self.assertRaisesRegex(ReviewContractError, "local .html basename"):
            validate_review_packet(escaped)

        uppercase = self.valid_packet()
        uppercase["sheets"][0]["html_sha256"] = "A" * 64
        with self.assertRaisesRegex(ReviewContractError, "lowercase SHA-256"):
            validate_review_packet(uppercase)


if __name__ == "__main__":
    unittest.main()
