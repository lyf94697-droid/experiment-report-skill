from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from universal_report.template_catalog import (
    audit_builtin_templates,
    load_template_catalog,
    recommend_builtin_template,
    resolve_template_selection,
)


REPO_ROOT = Path(__file__).resolve().parents[1]


class TemplateCatalogTests(unittest.TestCase):
    def test_catalog_contains_exactly_five_neutral_templates(self) -> None:
        catalog = load_template_catalog(REPO_ROOT)

        self.assertEqual(len(catalog["templates"]), 5)
        self.assertEqual(
            {item["id"] for item in catalog["templates"]},
            {
                "neutral-classic-lab",
                "neutral-bordered-lab",
                "neutral-engineering-lab",
                "neutral-course-design",
                "neutral-modern-minimal",
            },
        )
        self.assertTrue(all(Path(item["path"]).is_file() for item in catalog["templates"]))

    def test_user_template_always_has_priority(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            user_template = Path(temp_dir) / "teacher-template.docx"
            user_template.write_bytes(b"placeholder")

            selection = resolve_template_selection(
                REPO_ROOT,
                report_type="实验报告",
                user_template=user_template,
                preference="现代简洁实验报告",
            )

        self.assertEqual(selection["source"], "user")
        self.assertIsNone(selection["templateId"])
        self.assertTrue(selection["preserveFormatting"])

    def test_auto_selection_uses_course_and_report_type(self) -> None:
        catalog = load_template_catalog(REPO_ROOT)

        engineering = recommend_builtin_template(
            catalog,
            report_type="实验报告",
            course_name="计算机网络",
        )
        course_design = recommend_builtin_template(
            catalog,
            report_type="课程设计报告",
            course_name="软件工程",
        )
        modern = recommend_builtin_template(
            catalog,
            report_type="实验报告",
            request_text="希望使用现代极简排版",
        )

        self.assertEqual(engineering["template"]["id"], "neutral-engineering-lab")
        self.assertEqual(course_design["template"]["id"], "neutral-course-design")
        self.assertEqual(modern["template"]["id"], "neutral-modern-minimal")

    def test_all_builtins_pass_identity_and_provenance_gate(self) -> None:
        audit = audit_builtin_templates(REPO_ROOT)

        self.assertTrue(audit["passed"])
        self.assertEqual(audit["templateCount"], 5)
        self.assertTrue(all(item["identityAudit"]["passed"] for item in audit["templates"]))
        self.assertTrue(all(item["provenancePassed"] for item in audit["templates"]))


if __name__ == "__main__":
    unittest.main()
