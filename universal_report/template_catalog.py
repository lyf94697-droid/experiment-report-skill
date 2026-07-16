from __future__ import annotations

import json
import re
import zipfile
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


CATALOG_RELATIVE_PATH = Path("examples") / "report-templates" / "catalog.json"
SUPPORTED_TEMPLATE_EXTENSIONS = {".docx", ".doc"}
GENERIC_ORGANIZATION_LABELS = {
    "学校",
    "学院",
    "学校名称",
    "学院名称",
    "学校或学院名称",
    "学校/学院名称",
    "单位名称",
}
DEFAULT_FORBIDDEN_IDENTITY_TERMS = {
    "云南师范大学",
    "四川大学",
    "宜宾学院",
    "华中科技大学",
    "合肥工业大学",
    "浙江大学",
    "广州大学",
    "长春工业大学",
    "中国矿业大学",
    "信息学院",
    "计算机学院",
}
EXAMPLE_IDENTITY_PATTERNS = (
    re.compile(r"(?<!\d)\d{8,14}(?!\d)"),
    re.compile(r"学号\s*[:：]?\s*\d{5,}"),
)
ORGANIZATION_NAME_PATTERN = re.compile(
    r"[\u4e00-\u9fffA-Za-z·]{2,24}(?:大学|学院|学校)"
)


def _catalog_path(repo_root: Path | str, catalog_path: Path | str | None = None) -> Path:
    root = Path(repo_root).resolve()
    return Path(catalog_path).resolve() if catalog_path else root / CATALOG_RELATIVE_PATH


def load_template_catalog(
    repo_root: Path | str,
    catalog_path: Path | str | None = None,
) -> dict[str, Any]:
    path = _catalog_path(repo_root, catalog_path)
    payload = json.loads(path.read_text(encoding="utf-8-sig"))
    templates = payload.get("templates")
    if not isinstance(templates, list) or not templates:
        raise ValueError(f"Template catalog has no templates: {path}")

    seen: set[str] = set()
    resolved: list[dict[str, Any]] = []
    for raw in templates:
        item = dict(raw)
        template_id = str(item.get("id") or "").strip()
        relative_file = str(item.get("file") or "").strip()
        if not template_id or template_id in seen:
            raise ValueError(f"Template catalog contains an invalid or duplicate id: {template_id!r}")
        if not relative_file:
            raise ValueError(f"Template catalog entry has no file: {template_id}")
        seen.add(template_id)
        item["path"] = str((path.parent / relative_file).resolve())
        resolved.append(item)

    result = dict(payload)
    result["catalogPath"] = str(path)
    result["templates"] = resolved
    return result


def _normalize_report_type(value: str) -> str:
    lowered = str(value or "").strip().lower()
    if "课程设计" in lowered or lowered in {"course-design", "course-design-report"}:
        return "course-design-report"
    return "experiment-report"


def _find_explicit_template(
    templates: Iterable[dict[str, Any]],
    preference: str,
) -> dict[str, Any] | None:
    normalized = str(preference or "").strip().lower()
    if not normalized or normalized in {"auto", "自动", "自动选择", "智能选择"}:
        return None
    for item in templates:
        candidates = {
            str(item.get("id") or "").strip().lower(),
            str(item.get("displayName") or "").strip().lower(),
            *(str(alias).strip().lower() for alias in item.get("aliases", [])),
        }
        if normalized in candidates:
            return item
    raise ValueError(f"Unknown built-in template preference: {preference}")


def recommend_builtin_template(
    catalog: dict[str, Any],
    report_type: str,
    course_name: str = "",
    preference: str = "auto",
    request_text: str = "",
) -> dict[str, Any]:
    templates = list(catalog["templates"])
    explicit = _find_explicit_template(templates, preference)
    if explicit is not None:
        return {
            "template": explicit,
            "reason": "用户明确选择了该内置模板。",
            "automatic": False,
        }

    normalized_type = _normalize_report_type(report_type)
    context = " ".join((str(course_name or ""), str(request_text or ""))).lower()
    target_id = str(catalog.get("defaults", {}).get(normalized_type) or "")
    reason = "未提供用户模板，使用该报告类型的默认中性模板。"

    if normalized_type == "course-design-report":
        target_id = "neutral-course-design"
        reason = "课程设计报告优先使用带封面和正文分节的中性模板。"
    elif any(keyword in context for keyword in ("闭合外框", "外框", "表格式", "纸质填写", "传统学校")):
        target_id = "neutral-bordered-lab"
        reason = "需求强调闭合外框或传统表格式版面。"
    elif any(keyword in context for keyword in ("现代", "简洁", "极简", "minimal", "modern")):
        target_id = "neutral-modern-minimal"
        reason = "需求强调现代、简洁或极简版式。"
    elif any(
        keyword in context
        for keyword in (
            "计算机网络",
            "操作系统",
            "数据库",
            "java",
            "web",
            "android",
            "软件工程",
            "编程",
            "工程",
            "代码",
        )
    ):
        target_id = "neutral-engineering-lab"
        reason = "计算机与工程类实验优先使用代码和图表层级更清楚的技术模板。"

    for item in templates:
        if item["id"] == target_id:
            return {"template": item, "reason": reason, "automatic": True}
    raise ValueError(f"Built-in template id is not present in catalog: {target_id}")


def resolve_template_selection(
    repo_root: Path | str,
    report_type: str,
    user_template: Path | str | None = None,
    course_name: str = "",
    preference: str = "auto",
    request_text: str = "",
) -> dict[str, Any]:
    if user_template:
        path = Path(user_template).expanduser().resolve()
        if not path.exists() or not path.is_file():
            raise FileNotFoundError(f"User template does not exist: {path}")
        if path.suffix.lower() not in SUPPORTED_TEMPLATE_EXTENSIONS:
            raise ValueError(f"Unsupported user template extension: {path.suffix}")
        return {
            "source": "user",
            "path": str(path),
            "templateId": None,
            "displayName": path.name,
            "preserveFormatting": True,
            "automatic": False,
            "reason": "用户上传模板具有最高优先级，按原模板保真处理。",
        }

    catalog = load_template_catalog(repo_root)
    recommendation = recommend_builtin_template(
        catalog,
        report_type=report_type,
        course_name=course_name,
        preference=preference,
        request_text=request_text,
    )
    item = recommendation["template"]
    path = Path(item["path"])
    if not path.exists():
        raise FileNotFoundError(f"Built-in template file does not exist: {path}")
    return {
        "source": "builtin",
        "path": str(path),
        "templateId": item["id"],
        "displayName": item["displayName"],
        "preserveFormatting": True,
        "automatic": recommendation["automatic"],
        "reason": recommendation["reason"],
    }


def _extract_docx_text(path: Path) -> tuple[str, list[str]]:
    text_parts: list[str] = []
    media_parts: list[str] = []
    with zipfile.ZipFile(path, "r") as archive:
        for name in archive.namelist():
            lowered = name.lower()
            if lowered.startswith("word/media/") and not lowered.endswith("/"):
                media_parts.append(name)
            if not lowered.endswith(".xml"):
                continue
            try:
                root = ET.fromstring(archive.read(name))
            except ET.ParseError:
                continue
            for element in root.iter():
                if element.text and element.text.strip():
                    text_parts.append(element.text.strip())
    return "\n".join(text_parts), media_parts


def scan_template_identity(
    template_path: Path | str,
    forbidden_terms: Iterable[str] | None = None,
) -> dict[str, Any]:
    path = Path(template_path).resolve()
    violations: list[dict[str, str]] = []
    if path.suffix.lower() != ".docx":
        return {
            "schemaVersion": "1.0",
            "path": str(path),
            "passed": False,
            "violations": [
                {
                    "code": "unsupported-format",
                    "value": path.suffix,
                    "message": "内置模板必须是可审计的 DOCX。",
                }
            ],
        }

    try:
        full_text, media_parts = _extract_docx_text(path)
    except (OSError, zipfile.BadZipFile) as exc:
        return {
            "schemaVersion": "1.0",
            "path": str(path),
            "passed": False,
            "violations": [
                {
                    "code": "invalid-docx",
                    "value": str(exc),
                    "message": "模板不是有效的 DOCX 包。",
                }
            ],
        }

    terms = set(DEFAULT_FORBIDDEN_IDENTITY_TERMS)
    terms.update(str(term).strip() for term in (forbidden_terms or []) if str(term).strip())
    for term in sorted(terms):
        if term and term in full_text:
            violations.append(
                {
                    "code": "forbidden-identity-term",
                    "value": term,
                    "message": "模板包含真实学校、学院或其他机构标识。",
                }
            )

    for match in ORGANIZATION_NAME_PATTERN.findall(full_text):
        if match not in GENERIC_ORGANIZATION_LABELS:
            violations.append(
                {
                    "code": "organization-name",
                    "value": match,
                    "message": "模板包含疑似真实机构名称。",
                }
            )

    for pattern in EXAMPLE_IDENTITY_PATTERNS:
        for match in pattern.findall(full_text):
            value = match if isinstance(match, str) else "".join(match)
            violations.append(
                {
                    "code": "example-student-identity",
                    "value": value,
                    "message": "模板包含疑似示例学号或其他长数字身份信息。",
                }
            )

    if media_parts:
        violations.append(
            {
                "code": "embedded-media",
                "value": ", ".join(media_parts),
                "message": "中性内置模板不得携带校徽、照片或其他不可追溯图片资产。",
            }
        )

    unique: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for item in violations:
        key = (item["code"], item["value"])
        if key not in seen:
            unique.append(item)
            seen.add(key)
    return {
        "schemaVersion": "1.0",
        "path": str(path),
        "passed": not unique,
        "embeddedMediaCount": len(media_parts),
        "violations": unique,
    }


def audit_builtin_templates(repo_root: Path | str) -> dict[str, Any]:
    catalog = load_template_catalog(repo_root)
    results: list[dict[str, Any]] = []
    for item in catalog["templates"]:
        identity = scan_template_identity(
            item["path"],
            forbidden_terms=catalog.get("identityPolicy", {}).get("forbiddenTerms", []),
        )
        provenance_ok = (
            item.get("origin") == "original-neutral-reconstruction"
            and item.get("license") == "MIT"
            and item.get("thirdPartyAssets") is False
        )
        results.append(
            {
                "id": item["id"],
                "path": item["path"],
                "fileExists": Path(item["path"]).is_file(),
                "identityAudit": identity,
                "provenancePassed": provenance_ok,
                "passed": Path(item["path"]).is_file() and identity["passed"] and provenance_ok,
            }
        )
    return {
        "schemaVersion": "1.0",
        "catalogPath": catalog["catalogPath"],
        "templateCount": len(results),
        "passed": len(results) == 5 and all(item["passed"] for item in results),
        "templates": results,
    }
