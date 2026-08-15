from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from xlsx_artifact_tool import XlsxArtifactTool
from xlsx_artifact_tool import preview as preview_module
from xlsx_artifact_tool.api import validate_create_model
from xlsx_artifact_tool.preview import render_xlsx_preview


class CreateModelValidationTests(unittest.TestCase):
    def test_optional_workbook_id_is_validated_once_and_create_accepts_it(self):
        model = {
            "workbook_id": "book_quarterly",
            "sheets": [{"name": "Summary", "cells": {"A1": {"value": "Total"}}}],
        }

        validated = validate_create_model(model)

        self.assertEqual(validated, model)
        self.assertIsNot(validated, model)
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            result = XlsxArtifactTool(root / "work").create(model, root / "report.xlsx")
            self.assertEqual(result["status"], "ok")

    def test_create_schema_declares_canonical_structural_references(self):
        package_root = Path(__file__).parents[1]
        schemas = [
            json.loads((package_root / "schemas/create.schema.json").read_text(encoding="utf-8")),
            json.loads(
                (package_root / "xlsx_artifact_tool/resources/create.schema.json").read_text(encoding="utf-8")
            ),
        ]
        self.assertEqual(schemas[0], schemas[1])
        properties = schemas[0]["properties"]["model"]["properties"]["sheets"]["items"]["properties"]
        self.assertEqual(properties["cells"]["propertyNames"]["pattern"], "^[A-Z]{1,3}[1-9][0-9]{0,6}$")
        self.assertEqual(properties["freeze_panes"]["pattern"], "^(?!A1$)[A-Z]{1,3}[1-9][0-9]{0,6}$")
        self.assertEqual(
            properties["auto_filter"]["pattern"],
            "^[A-Z]{1,3}[1-9][0-9]{0,6}(?::[A-Z]{1,3}[1-9][0-9]{0,6})?$",
        )
        self.assertEqual(properties["column_widths"]["propertyNames"]["pattern"], "^[A-Z]{1,3}$")
        self.assertEqual(properties["row_heights"]["propertyNames"]["pattern"], "^[1-9][0-9]{0,6}$")
        self.assertEqual(
            properties["merged_ranges"]["items"]["pattern"],
            "^[A-Z]{1,3}[1-9][0-9]{0,6}:[A-Z]{1,3}[1-9][0-9]{0,6}$",
        )

    def test_complete_structural_model_is_accepted_by_preview_and_create(self):
        model = {
            "workbook_id": "book_layout",
            "sheets": [
                {
                    "name": "Data",
                    "cells": {
                        "A1": {"value": "Item", "style": "header"},
                        "B1": {"value": "Amount", "style": "header"},
                        "A2": {"value": "A"},
                        "B2": {"formula": "=1+1", "style": "integer"},
                    },
                    "freeze_panes": "A2",
                    "auto_filter": "A1:B2",
                    "column_widths": {"A": 24, "B": 14},
                    "row_heights": {"1": 28},
                    "merged_ranges": [],
                }
            ],
        }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            preview = render_xlsx_preview(model, root / "preview")
            created = XlsxArtifactTool(root / "work").create(model, root / "report.xlsx")

            self.assertEqual(preview["status"], "previewed")
            self.assertEqual(created["status"], "ok")
            self.assertTrue((root / "report.xlsx").exists())

    def test_create_and_preview_share_closed_model_rules_without_publishing_invalid_input(self):
        invalid_models = [
            {"sheets": [{"name": "Data", "cells": {}}], "prompt": "change cells"},
            {"sheets": [{"name": "Data", "cells": {}, "callback": "run"}]},
            {"sheets": [{"name": "Data", "cells": {"A1": {"value": "x", "office_object": "cell"}}}]},
            {"sheets": [{"name": "Data", "cells": {"A:A": {"value": "not a cell"}}}]},
            {"sheets": [{"name": "Data", "cells": {}, "freeze_panes": "not-a-cell"}]},
            {"sheets": [{"name": "Data", "cells": {}, "freeze_panes": "A1"}]},
            {"sheets": [{"name": "Data", "cells": {}, "auto_filter": "not-a-range"}]},
            {"sheets": [{"name": "Data", "cells": {}, "auto_filter": "a1:b2"}]},
            {"sheets": [{"name": "Data", "cells": {}, "column_widths": {"A1": 12}}]},
            {"sheets": [{"name": "Data", "cells": {}, "column_widths": {"1": 12}}]},
            {"sheets": [{"name": "Data", "cells": {}, "column_widths": {"a": 12}}]},
            {"sheets": [{"name": "Data", "cells": {}, "merged_ranges": ["A:A"]}]},
        ]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            tool = XlsxArtifactTool(root / "work")
            for index, model in enumerate(invalid_models):
                with self.subTest(index=index):
                    workbook = root / f"invalid-{index}.xlsx"
                    preview = root / f"preview-{index}"
                    result = tool.create(model, workbook)
                    self.assertEqual((result["status"], result["reason"]), ("refused", "validation_failure"))
                    self.assertFalse(workbook.exists())
                    preview_result = render_xlsx_preview(model, preview)
                    self.assertEqual(
                        (preview_result["status"], preview_result["reason"]),
                        ("refused", "validation_failure"),
                    )
                    self.assertFalse(preview.exists())


class XlsxPreviewTests(unittest.TestCase):
    def setUp(self):
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        self.output = self.root / "preview"

    def tearDown(self):
        self.tempdir.cleanup()

    def test_publishes_digest_bound_escaped_html_with_stable_cell_addresses(self):
        model = {
            "workbook_id": "book_demo",
            "sheets": [
                {
                    "name": 'Summary <&"',
                    "cells": {
                        "A1": {"value": '<script>alert("x")</script>&'},
                        "B2": {"formula": "=SUM(A1:A2)"},
                    },
                }
            ],
        }

        result = render_xlsx_preview(model, self.output)

        encoded = json.dumps(
            validate_create_model(model),
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
        expected_revision = hashlib.sha256(encoded).hexdigest()
        review_path = self.output / "review.json"
        html_path = self.output / "sheet-01.html"
        self.assertEqual(result, {
            "status": "previewed",
            "workbook_id": "book_demo",
            "interaction": "chat_only",
            "revision": expected_revision,
            "review_contract": str(review_path.resolve()),
            "display_artifacts": [str(html_path.resolve())],
            "output": str(self.output.resolve()),
        })
        review = json.loads(review_path.read_text(encoding="utf-8"))
        self.assertEqual(review["revision"], expected_revision)
        self.assertEqual(review["fidelity"], "structural_preview_not_excel_render")
        self.assertEqual(
            review["sheets"][0]["html_sha256"],
            hashlib.sha256(html_path.read_bytes()).hexdigest(),
        )
        rendered = html_path.read_text(encoding="utf-8")
        self.assertNotIn("<script", rendered.lower())
        self.assertNotIn("href=", rendered.lower())
        self.assertNotIn("src=", rendered.lower())
        self.assertIn("&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;&amp;", rendered)
        self.assertIn('data-sheet="Summary &lt;&amp;&quot;"', rendered)
        self.assertIn('data-cell="A1"', rendered)
        self.assertIn('data-cell="B2"', rendered)
        self.assertIn("=SUM(A1:A2)", rendered)

    def test_omits_hidden_sheets_and_reports_bounded_cell_truncation(self):
        model = {
            "sheets": [
                {
                    "name": "Data",
                    "cells": {
                        "A1": {"value": "shown"},
                        "A101": {"value": "row-secret"},
                        "AO1": {"value": "column-secret"},
                    },
                },
                {"name": "Hidden", "state": "hidden", "cells": {"A1": {"value": "hidden-secret"}}},
                {"name": "VeryHidden", "state": "veryHidden", "cells": {"A1": {"value": "very-secret"}}},
            ]
        }

        result = render_xlsx_preview(model, self.output)

        review = json.loads(Path(result["review_contract"]).read_text(encoding="utf-8"))
        self.assertEqual([sheet["sheet"] for sheet in review["sheets"]], ["Data"])
        self.assertEqual(review["diagnostics"], {
            "truncated_sheets": [
                {"sheet": "Data", "rows_shown": 100, "columns_shown": 40, "omitted_cells": 2}
            ],
            "omitted_sheets": [
                {"sheet": "Hidden", "reason": "hidden"},
                {"sheet": "VeryHidden", "reason": "very_hidden"},
            ],
        })
        rendered = Path(result["display_artifacts"][0]).read_text(encoding="utf-8")
        self.assertIn('data-cell="AN100"', rendered)
        self.assertNotIn('data-cell="A101"', rendered)
        self.assertNotIn('data-cell="AO1"', rendered)
        self.assertNotIn("secret", rendered)

    def test_limits_visible_sheet_artifacts_to_twenty_and_reports_omissions(self):
        model = {
            "sheets": [
                {"name": f"Sheet{number}", "cells": {"A1": {"value": number}}}
                for number in range(1, 23)
            ]
        }

        result = render_xlsx_preview(model, self.output)

        review = json.loads(Path(result["review_contract"]).read_text(encoding="utf-8"))
        self.assertEqual(len(result["display_artifacts"]), 20)
        self.assertEqual(len(review["sheets"]), 20)
        self.assertEqual(review["diagnostics"]["omitted_sheets"], [
            {"sheet": "Sheet21", "reason": "artifact_limit"},
            {"sheet": "Sheet22", "reason": "artifact_limit"},
        ])
        self.assertFalse((self.output / "sheet-21.html").exists())

    def test_atomically_replaces_existing_review_and_changes_exact_model_revision(self):
        first_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "first"}}}]}
        first = render_xlsx_preview(first_model, self.output)
        (self.output / "stale.txt").write_text("stale", encoding="utf-8")
        second_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "second"}}}]}

        second = render_xlsx_preview(second_model, self.output)

        self.assertNotEqual(first["revision"], second["revision"])
        expected = hashlib.sha256(
            json.dumps(
                validate_create_model(second_model),
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            ).encode("utf-8")
        ).hexdigest()
        self.assertEqual(second["revision"], expected)
        self.assertFalse((self.output / "stale.txt").exists())
        self.assertIn("second", (self.output / "sheet-01.html").read_text(encoding="utf-8"))
        review = json.loads((self.output / "review.json").read_text(encoding="utf-8"))
        self.assertEqual(review["revision"], second["revision"])

    def test_replacing_review_never_makes_published_directory_disappear(self):
        first_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "first"}}}]}
        render_xlsx_preview(first_model, self.output)
        original_replace = preview_module.os.replace
        visibility = []

        def observe_replace(source, target):
            visibility.append(
                self.output.exists()
                and (self.output / "sheet-01.html").exists()
                and (self.output / "review.json").exists()
            )
            original_replace(source, target)
            visibility.append(
                self.output.exists()
                and (self.output / "sheet-01.html").exists()
                and (self.output / "review.json").exists()
            )

        second_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "second"}}}]}
        with mock.patch.object(preview_module.os, "replace", side_effect=observe_replace):
            result = render_xlsx_preview(second_model, self.output)

        self.assertEqual(result["status"], "previewed")
        self.assertTrue(visibility)
        self.assertTrue(all(visibility))

    def test_failed_atomic_swap_restores_previous_complete_review(self):
        first_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "first"}}}]}
        render_xlsx_preview(first_model, self.output)
        original_html = (self.output / "sheet-01.html").read_bytes()
        original_review = (self.output / "review.json").read_bytes()

        def fail_publish(source, target):
            raise OSError("simulated publish failure")

        second_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "second"}}}]}
        with mock.patch.object(preview_module, "_exchange_directories", side_effect=fail_publish):
            with self.assertRaisesRegex(OSError, "simulated publish failure"):
                render_xlsx_preview(second_model, self.output)

        self.assertEqual((self.output / "sheet-01.html").read_bytes(), original_html)
        self.assertEqual((self.output / "review.json").read_bytes(), original_review)

    def test_cleanup_failure_after_commit_returns_successful_new_review(self):
        first_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "first"}}}]}
        first = render_xlsx_preview(first_model, self.output)
        original_rmtree = preview_module.shutil.rmtree
        backup = self.output.with_name(f".{self.output.name}.old")

        def fail_backup_cleanup(path, *args, **kwargs):
            if Path(path) == backup:
                raise OSError("simulated backup cleanup failure")
            original_rmtree(path, *args, **kwargs)

        second_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "second"}}}]}
        with mock.patch.object(preview_module.shutil, "rmtree", side_effect=fail_backup_cleanup):
            second = render_xlsx_preview(second_model, self.output)

        self.assertEqual(second["status"], "previewed")
        self.assertNotEqual(second["revision"], first["revision"])
        self.assertIn("second", (self.output / "sheet-01.html").read_text(encoding="utf-8"))
        review = json.loads((self.output / "review.json").read_text(encoding="utf-8"))
        self.assertEqual(review["revision"], second["revision"])
        self.assertTrue(backup.exists())

    def test_non_oserror_cleanup_failure_after_commit_returns_successful_new_review(self):
        first_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "first"}}}]}
        first = render_xlsx_preview(first_model, self.output)
        original_rmtree = preview_module.shutil.rmtree
        backup = self.output.with_name(f".{self.output.name}.old")

        def fail_backup_cleanup(path, *args, **kwargs):
            if Path(path) == backup:
                raise RecursionError("simulated deeply nested backup")
            original_rmtree(path, *args, **kwargs)

        second_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "second"}}}]}
        with mock.patch.object(preview_module.shutil, "rmtree", side_effect=fail_backup_cleanup):
            second = render_xlsx_preview(second_model, self.output)

        self.assertEqual(second["status"], "previewed")
        self.assertNotEqual(second["revision"], first["revision"])
        self.assertIn("second", (self.output / "sheet-01.html").read_text(encoding="utf-8"))
        review = json.loads((self.output / "review.json").read_text(encoding="utf-8"))
        self.assertEqual(review["revision"], second["revision"])
        self.assertTrue(backup.exists())

    def test_failed_old_generation_rename_is_cleaned_without_unbounded_temporary_directory(self):
        first_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "first"}}}]}
        render_xlsx_preview(first_model, self.output)
        original_replace = preview_module.os.replace
        backup = self.output.with_name(f".{self.output.name}.old")

        def fail_old_generation_rename(source, target):
            if Path(target) == backup:
                raise RecursionError("simulated old-generation rename failure")
            original_replace(source, target)

        second_model = {"sheets": [{"name": "Data", "cells": {"A1": {"value": "second"}}}]}
        with mock.patch.object(preview_module.os, "replace", side_effect=fail_old_generation_rename):
            result = render_xlsx_preview(second_model, self.output)

        self.assertEqual(result["status"], "previewed")
        self.assertIn("second", (self.output / "sheet-01.html").read_text(encoding="utf-8"))
        self.assertFalse(backup.exists())
        self.assertEqual([path for path in self.root.iterdir() if path != self.output], [])


if __name__ == "__main__":
    unittest.main()
