from __future__ import annotations

from pathlib import Path
from typing import Any

from .template_contract import analyze_template


def _check(
    code: str,
    template_value: Any,
    document_value: Any,
    tolerance: Any,
    location: str,
    *,
    comparator=None,
) -> dict[str, Any]:
    if comparator is None:
        passed = template_value == document_value
    else:
        passed = comparator(template_value, document_value, tolerance)
    return {
        "code": code,
        "templateValue": template_value,
        "documentValue": document_value,
        "tolerance": tolerance,
        "passed": bool(passed),
        "location": location,
    }


def _numeric_close(template_value: Any, document_value: Any, tolerance: int) -> bool:
    if template_value is None and document_value is None:
        return True
    if template_value is None or document_value is None:
        return False
    return abs(int(template_value) - int(document_value)) <= tolerance


def _margins_close(template_value: Any, document_value: Any, tolerance: int) -> bool:
    if not isinstance(template_value, dict) or not isinstance(document_value, dict):
        return template_value == document_value
    keys = ("top", "right", "bottom", "left", "header", "footer")
    return all(
        _numeric_close(template_value.get(key), document_value.get(key), tolerance)
        for key in keys
    )


def validate_format(
    template_path: Path | str, document_path: Path | str
) -> dict[str, Any]:
    template = analyze_template(template_path)
    document = analyze_template(document_path)
    template_section = (template["page"]["sections"] or [{}])[0]
    document_section = (document["page"]["sections"] or [{}])[0]
    template_roles = template["styles"]["roles"]
    document_roles = document["styles"]["roles"]
    template_metadata = template["structure"].get("metadataTable") or {}
    document_metadata = document["structure"].get("metadataTable") or {}

    checks = [
        _check(
            "page-size",
            template_section.get("pageSize"),
            document_section.get("pageSize"),
            10,
            "section:1",
            comparator=lambda a, b, tolerance: all(
                _numeric_close(a.get(key), b.get(key), tolerance)
                for key in ("widthTwips", "heightTwips")
            )
            if isinstance(a, dict) and isinstance(b, dict)
            else a == b,
        ),
        _check(
            "page-margins",
            template_section.get("margins"),
            document_section.get("margins"),
            20,
            "section:1",
            comparator=_margins_close,
        ),
        _check(
            "page-border",
            template_section.get("pageBorder"),
            document_section.get("pageBorder"),
            "exact",
            "section:1",
        ),
        _check(
            "report-title-font",
            (template_roles["reportTitle"].get("font") or {}).get("eastAsia"),
            (document_roles["reportTitle"].get("font") or {}).get("eastAsia"),
            "exact",
            "role:reportTitle",
        ),
        _check(
            "report-title-size",
            template_roles["reportTitle"].get("sizeHalfPoints"),
            document_roles["reportTitle"].get("sizeHalfPoints"),
            0,
            "role:reportTitle",
            comparator=_numeric_close,
        ),
        _check(
            "report-title-alignment",
            template_roles["reportTitle"].get("alignment"),
            document_roles["reportTitle"].get("alignment"),
            "exact",
            "role:reportTitle",
        ),
        _check(
            "body-font",
            (template_roles["body"].get("font") or {}).get("eastAsia"),
            (document_roles["body"].get("font") or {}).get("eastAsia"),
            "exact",
            "role:body",
        ),
        _check(
            "body-size",
            template_roles["body"].get("sizeHalfPoints"),
            document_roles["body"].get("sizeHalfPoints"),
            0,
            "role:body",
            comparator=_numeric_close,
        ),
        _check(
            "body-line-spacing",
            template_roles["body"].get("lineSpacing"),
            document_roles["body"].get("lineSpacing"),
            "exact",
            "role:body",
        ),
        _check(
            "body-first-line-indent",
            (template_roles["body"].get("indent") or {}).get("firstLine"),
            (document_roles["body"].get("indent") or {}).get("firstLine"),
            0,
            "role:body",
            comparator=_numeric_close,
        ),
        _check(
            "metadata-table-grid",
            template_metadata.get("gridColumnsTwips"),
            document_metadata.get("gridColumnsTwips"),
            20,
            f"table:{template_metadata.get('tableIndex', 1)}",
            comparator=lambda a, b, tolerance: (
                (a is None and b is None)
                or (
                    isinstance(a, list)
                    and isinstance(b, list)
                    and len(a) == len(b)
                    and all(_numeric_close(x, y, tolerance) for x, y in zip(a, b))
                )
            ),
        ),
        _check(
            "metadata-vertical-text-risk",
            bool(template_metadata.get("verticalTextRisk")),
            bool(document_metadata.get("verticalTextRisk")),
            "must-match-template",
            f"table:{document_metadata.get('tableIndex', 1)}",
        ),
    ]
    failed = [item for item in checks if not item["passed"]]
    return {
        "schemaVersion": "1.0",
        "templatePath": template["source"]["path"],
        "documentPath": document["source"]["path"],
        "passed": not failed,
        "summary": {
            "checkCount": len(checks),
            "passedCount": len(checks) - len(failed),
            "failedCount": len(failed),
            "failedCodes": [item["code"] for item in failed],
        },
        "checks": checks,
    }
