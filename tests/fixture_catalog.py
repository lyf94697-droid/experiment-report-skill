from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

from tests.fixture_factory import make_template


def build_fixture_catalog(output_dir: Path) -> dict[str, dict[str, Any]]:
    output_dir.mkdir(parents=True, exist_ok=True)
    repo_root = Path(__file__).resolve().parents[1]
    catalog: dict[str, dict[str, Any]] = {}

    default_source = (
        repo_root / "examples" / "report-templates" / "experiment-report-template.docx"
    )
    default_target = output_dir / "repository-default-template.docx"
    shutil.copy2(default_source, default_target)
    catalog["repository-default-template"] = {
        "path": default_target,
        "description": "仓库内置默认实验报告模板",
    }

    generated = {
        "four-column-metadata": make_template(
            output_dir / "four-column-metadata.docx",
            columns=[1800, 3000, 1800, 4200],
        ),
        "five-column-metadata": make_template(
            output_dir / "five-column-metadata.docx",
            columns=[1400, 2200, 1400, 2600, 2600],
        ),
        "no-outer-border": make_template(
            output_dir / "no-outer-border.docx",
            outer_table_border=False,
        ),
        "existing-page-border": make_template(
            output_dir / "existing-page-border.docx",
            page_border=True,
        ),
        "cover-body-sections": make_template(
            output_dir / "cover-body-sections.docx",
            sections=2,
        ),
        "image-placeholder": make_template(
            output_dir / "image-placeholder.docx",
            image_placeholder=True,
        ),
        "blank-docx": make_template(
            output_dir / "blank-docx.docx",
            blank=True,
        ),
        "long-course-name-cell-pressure": make_template(
            output_dir / "long-course-name-cell-pressure.docx",
            columns=[1300, 1400, 1300, 1400],
            long_course_name=True,
        ),
    }
    descriptions = {
        "four-column-metadata": "四列学生信息表模板",
        "five-column-metadata": "五列学生信息表模板",
        "no-outer-border": "信息表缺少外框的风险模板",
        "existing-page-border": "已经包含页面边框的模板",
        "cover-body-sections": "封面与正文分节的模板",
        "image-placeholder": "带图片占位符的模板",
        "blank-docx": "空白普通 DOCX 模板",
        "long-course-name-cell-pressure": "中文长课程名造成单元格拥挤的模板",
    }
    for name, path in generated.items():
        catalog[name] = {"path": path, "description": descriptions[name]}

    legacy_doc = output_dir / "legacy-doc-conversion-failure.doc"
    legacy_doc.write_bytes(b"not-a-real-binary-doc")
    catalog["legacy-doc-conversion-failure"] = {
        "path": legacy_doc,
        "description": "旧版 DOC 无可用转换器时的失败场景",
    }
    return catalog
