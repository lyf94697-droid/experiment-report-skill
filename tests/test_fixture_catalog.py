from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from tests.fixture_catalog import build_fixture_catalog
from universal_report.template_contract import analyze_template


class FixtureCatalogTests(unittest.TestCase):
    def test_catalog_builds_all_required_template_scenarios(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            catalog = build_fixture_catalog(Path(temp_dir))

            expected = {
                "repository-default-template",
                "four-column-metadata",
                "five-column-metadata",
                "no-outer-border",
                "existing-page-border",
                "cover-body-sections",
                "image-placeholder",
                "blank-docx",
                "legacy-doc-conversion-failure",
                "long-course-name-cell-pressure",
            }
            self.assertEqual(set(catalog), expected)
            for name, item in catalog.items():
                self.assertTrue(item["path"].exists(), name)
                if item["path"].suffix.lower() == ".docx":
                    contract = analyze_template(item["path"])
                    self.assertEqual(contract["schemaVersion"], "1.0")


if __name__ == "__main__":
    unittest.main()
