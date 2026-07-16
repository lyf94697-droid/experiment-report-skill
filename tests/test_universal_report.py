from __future__ import annotations

import json
import os
import tempfile
import unittest
from pathlib import Path

from tests.fixture_factory import make_template
from universal_report.config import load_config
from universal_report.content import build_report_plan
from universal_report.format_validation import validate_format
from universal_report.images import build_image_manifest
from universal_report.pipeline import PipelineRun
from universal_report.template_contract import (
    analyze_template,
    analyze_template_cached,
    recommend_quality_mode,
)


class TemplateContractTests(unittest.TestCase):
    def test_contract_resolves_inherited_styles_and_table_geometry(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            template = make_template(
                Path(temp_dir) / "four-column.docx",
                columns=[1800, 3000, 1800, 4200],
                title_font="黑体",
                title_size=52,
                body_font="宋体",
                body_size=24,
                body_line=300,
                first_line=480,
            )

            contract = analyze_template(template)

            self.assertEqual(contract["schemaVersion"], "1.0")
            self.assertEqual(contract["styles"]["roles"]["reportTitle"]["font"]["eastAsia"], "黑体")
            self.assertEqual(contract["styles"]["roles"]["reportTitle"]["sizeHalfPoints"], 52)
            self.assertEqual(contract["styles"]["roles"]["body"]["font"]["eastAsia"], "宋体")
            self.assertEqual(contract["styles"]["roles"]["body"]["lineSpacing"]["line"], 300)
            self.assertEqual(contract["styles"]["roles"]["body"]["indent"]["firstLine"], 480)
            self.assertEqual(contract["tables"][0]["gridColumnsTwips"], [1800, 3000, 1800, 4200])
            self.assertEqual(contract["structure"]["metadataTable"]["columnCount"], 4)
            self.assertFalse(contract["structure"]["metadataTable"]["verticalTextRisk"])

    def test_contract_detects_sections_page_border_placeholders_and_five_columns(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            template = make_template(
                Path(temp_dir) / "five-column.docx",
                columns=[1400, 2100, 1400, 2600, 2600],
                page_border=True,
                sections=2,
                image_placeholder=True,
            )

            contract = analyze_template(template)

            self.assertEqual(len(contract["page"]["sections"]), 2)
            self.assertTrue(contract["page"]["sections"][0]["differentFirstPage"])
            self.assertEqual(
                set(contract["page"]["sections"][0]["pageBorder"]["sides"]),
                {"top", "left", "bottom", "right"},
            )
            self.assertEqual(contract["structure"]["metadataTable"]["columnCount"], 5)
            self.assertEqual(len(contract["structure"]["imagePlaceholders"]), 1)

    def test_cache_reuses_hash_and_invalidates_after_template_change(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = make_template(root / "template.docx")
            cache_dir = root / "cache"

            first = analyze_template_cached(template, cache_dir)
            second = analyze_template_cached(template, cache_dir)
            make_template(template, title_size=60)
            third = analyze_template_cached(template, cache_dir)

            self.assertFalse(first["cache"]["hit"])
            self.assertTrue(second["cache"]["hit"])
            self.assertNotEqual(first["source"]["sha256"], third["source"]["sha256"])
            self.assertFalse(third["cache"]["hit"])

    def test_risk_recommends_strict_for_blank_or_unframed_template(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            template = make_template(
                Path(temp_dir) / "blank.docx",
                blank=True,
                outer_table_border=False,
            )
            contract = analyze_template(template)

            recommendation = recommend_quality_mode(contract, verified=False)

            self.assertEqual(recommendation["recommendedMode"], "strict")
            self.assertGreaterEqual(recommendation["riskScore"], 50)
            self.assertIn("blank-template", recommendation["reasons"])
            self.assertIn("metadata-table-without-outer-border", recommendation["reasons"])


class FormatValidationTests(unittest.TestCase):
    def test_format_validation_reports_values_tolerances_and_locations(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = make_template(root / "template.docx", title_size=52)
            final = make_template(
                root / "final.docx",
                title_size=44,
                margins=(1440, 1000, 1440, 1440),
                columns=[1800, 2700, 1800, 4500],
            )

            result = validate_format(template, final)

            self.assertFalse(result["passed"])
            checks = {item["code"]: item for item in result["checks"]}
            self.assertEqual(checks["report-title-size"]["templateValue"], 52)
            self.assertEqual(checks["report-title-size"]["documentValue"], 44)
            self.assertIn("tolerance", checks["report-title-size"])
            self.assertEqual(checks["report-title-size"]["location"], "role:reportTitle")
            self.assertFalse(checks["page-margins"]["passed"])
            self.assertFalse(checks["metadata-table-grid"]["passed"])

    def test_existing_template_vertical_text_risk_is_reported_without_false_drift(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            template = make_template(
                Path(temp_dir) / "long-course.docx",
                columns=[1300, 1400, 1300, 1400],
                long_course_name=True,
            )

            result = validate_format(template, template)

            self.assertTrue(result["passed"])
            check = next(
                item
                for item in result["checks"]
                if item["code"] == "metadata-vertical-text-risk"
            )
            self.assertTrue(check["templateValue"])
            self.assertTrue(check["documentValue"])
            self.assertEqual(check["tolerance"], "must-match-template")

    def test_template_without_metadata_table_matches_itself(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            template = make_template(
                Path(temp_dir) / "blank.docx",
                blank=True,
                include_metadata_table=False,
            )

            result = validate_format(template, template)

            self.assertTrue(result["passed"])
            grid_check = next(
                item
                for item in result["checks"]
                if item["code"] == "metadata-table-grid"
            )
            self.assertIsNone(grid_check["templateValue"])
            self.assertIsNone(grid_check["documentValue"])


class PipelineAndContentTests(unittest.TestCase):
    def test_image_manifest_selects_exact_count_and_records_reasons(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            images = []
            for name, data in (
                ("01-setup.png", b"setup"),
                ("02-result.png", b"result"),
                ("03-result-copy.png", b"result"),
            ):
                path = root / name
                path.write_bytes(data)
                images.append(path)

            manifest = build_image_manifest(images, requested_count=2)

            self.assertEqual(len(manifest["images"]), 2)
            self.assertEqual(manifest["requestedCount"], 2)
            self.assertTrue(all(item["selectionReason"] for item in manifest["images"]))
            self.assertEqual(manifest["duplicatesFiltered"], 1)
            self.assertTrue(all(item["layout"]["mode"] == "single" for item in manifest["images"]))

    def test_content_plan_is_course_specific_and_student_variants_differ(self) -> None:
        network = build_report_plan(
            course_name="计算机网络",
            experiment_name="VLAN 间通信",
            detail_level="standard",
            variant_seed="student-a",
        )
        operating_system = build_report_plan(
            course_name="操作系统",
            experiment_name="进程调度",
            detail_level="long",
            variant_seed="student-b",
        )

        self.assertIn("网络拓扑与地址规划", [item["title"] for item in network["sections"]])
        self.assertIn("进程状态与调度依据", [item["title"] for item in operating_system["sections"]])
        self.assertNotEqual(network["writingVariant"], operating_system["writingVariant"])
        self.assertGreater(operating_system["targetCharacters"], network["targetCharacters"])

    def test_pipeline_records_structured_stage_results_and_needs_fix_status(self) -> None:
        run = PipelineRun()
        run.complete("materials", {"count": 3})
        run.complete("template-analysis", {"risk": "high"})
        run.fail("visual-validation", "render-unavailable", "未找到 LibreOffice", "安装 LibreOffice 后重试")

        payload = run.to_dict()

        self.assertEqual(payload["status"], "needs-fix")
        self.assertEqual(payload["currentStage"], "visual-validation")
        self.assertEqual(payload["stages"][0]["output"]["count"], 3)
        self.assertEqual(payload["errors"][0]["suggestion"], "安装 LibreOffice 后重试")

    def test_config_uses_environment_paths_without_hard_coding_local_machine(self) -> None:
        old_value = os.environ.get("EXPERIMENT_REPORT_OUTPUT_ROOT")
        try:
            os.environ["EXPERIMENT_REPORT_OUTPUT_ROOT"] = str(Path("D:/portable/reports"))
            config = load_config(repo_root=Path("D:/repo"))
        finally:
            if old_value is None:
                os.environ.pop("EXPERIMENT_REPORT_OUTPUT_ROOT", None)
            else:
                os.environ["EXPERIMENT_REPORT_OUTPUT_ROOT"] = old_value

        self.assertEqual(config["outputRoot"], str(Path("D:/portable/reports")))
        self.assertEqual(
            config["defaultTemplate"],
            str(Path("D:/repo/examples/report-templates/experiment-report-template.docx")),
        )


if __name__ == "__main__":
    unittest.main()
