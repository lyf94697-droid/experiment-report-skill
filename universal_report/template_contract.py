from __future__ import annotations

import hashlib
import json
import re
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W_NS, "r": R_NS}
ANALYZER_VERSION = "1.3"


def _q(name: str) -> str:
    return f"{{{W_NS}}}{name}"


def _attr(node: ET.Element | None, name: str) -> str | None:
    return None if node is None else node.get(_q(name))


def _attrs(node: ET.Element | None) -> dict[str, str]:
    if node is None:
        return {}
    return {key.rsplit("}", 1)[-1]: value for key, value in node.attrib.items()}


def _child(node: ET.Element | None, name: str) -> ET.Element | None:
    return None if node is None else node.find(f"w:{name}", NS)


def _integer(value: str | None) -> int | None:
    if value is None:
        return None
    try:
        return int(value)
    except ValueError:
        return None


def _on_off(node: ET.Element | None) -> bool | None:
    if node is None:
        return None
    return _attr(node, "val") not in {"0", "false", "off"}


def _merge(*values: dict[str, Any]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for value in values:
        for key, item in value.items():
            if isinstance(item, dict) and isinstance(result.get(key), dict):
                result[key] = {**result[key], **item}
            else:
                result[key] = item
    return result


def _run_properties(node: ET.Element | None) -> dict[str, Any]:
    if node is None:
        return {}
    fonts = _child(node, "rFonts")
    return {
        key: value
        for key, value in {
            "font": _attrs(fonts) if fonts is not None else None,
            "sizeHalfPoints": _integer(_attr(_child(node, "sz"), "val")),
            "sizeCsHalfPoints": _integer(_attr(_child(node, "szCs"), "val")),
            "bold": _on_off(_child(node, "b")),
            "italic": _on_off(_child(node, "i")),
            "underline": _attr(_child(node, "u"), "val"),
            "color": _attr(_child(node, "color"), "val"),
            "characterSpacingTwips": _integer(_attr(_child(node, "spacing"), "val")),
            "scalePercent": _integer(_attr(_child(node, "w"), "val")),
            "verticalAlign": _attr(_child(node, "vertAlign"), "val"),
        }.items()
        if value is not None
    }


def _paragraph_properties(node: ET.Element | None) -> dict[str, Any]:
    if node is None:
        return {}
    spacing = _child(node, "spacing")
    indent = _child(node, "ind")
    tabs = _child(node, "tabs")
    tab_values = []
    if tabs is not None:
        tab_values = [_attrs(tab) for tab in tabs.findall("w:tab", NS)]
    numbering = _child(node, "numPr")
    return {
        key: value
        for key, value in {
            "alignment": _attr(_child(node, "jc"), "val"),
            "indent": {
                key: _integer(value) if value.lstrip("-").isdigit() else value
                for key, value in _attrs(indent).items()
            }
            if indent is not None
            else None,
            "lineSpacing": {
                key: _integer(value) if value.lstrip("-").isdigit() else value
                for key, value in _attrs(spacing).items()
            }
            if spacing is not None
            else None,
            "outlineLevel": _integer(_attr(_child(node, "outlineLvl"), "val")),
            "keepWithNext": _on_off(_child(node, "keepNext")),
            "keepLinesTogether": _on_off(_child(node, "keepLines")),
            "pageBreakBefore": _on_off(_child(node, "pageBreakBefore")),
            "widowControl": _on_off(_child(node, "widowControl")),
            "tabs": tab_values or None,
            "numbering": _attrs(numbering) if numbering is not None else None,
        }.items()
        if value is not None
    }


def _text(node: ET.Element) -> str:
    return "".join(text.text or "" for text in node.findall(".//w:t", NS)).strip()


class _StyleResolver:
    def __init__(self, styles_xml: ET.Element | None) -> None:
        self.styles: dict[str, dict[str, Any]] = {}
        self.default_paragraph: dict[str, Any] = {}
        self.default_run: dict[str, Any] = {}
        self.normal_style_id: str | None = None
        if styles_xml is None:
            return

        defaults = styles_xml.find("w:docDefaults", NS)
        if defaults is not None:
            self.default_paragraph = _paragraph_properties(
                defaults.find("w:pPrDefault/w:pPr", NS)
            )
            self.default_run = _run_properties(defaults.find("w:rPrDefault/w:rPr", NS))

        for style in styles_xml.findall("w:style", NS):
            style_id = _attr(style, "styleId")
            if not style_id:
                continue
            name_node = _child(style, "name")
            based_on = _child(style, "basedOn")
            item = {
                "id": style_id,
                "type": _attr(style, "type"),
                "name": _attr(name_node, "val"),
                "basedOn": _attr(based_on, "val"),
                "paragraph": _paragraph_properties(_child(style, "pPr")),
                "run": _run_properties(_child(style, "rPr")),
            }
            self.styles[style_id] = item
            if (
                item["type"] == "paragraph"
                and (item["name"] == "Normal" or _attr(style, "default") == "1")
                and self.normal_style_id is None
            ):
                self.normal_style_id = style_id

    def _style_chain(self, style_id: str | None, property_name: str) -> dict[str, Any]:
        chain: list[dict[str, Any]] = []
        visited: set[str] = set()
        while style_id and style_id not in visited and style_id in self.styles:
            visited.add(style_id)
            style = self.styles[style_id]
            chain.append(style[property_name])
            style_id = style["basedOn"]
        chain.reverse()
        return _merge(*chain)

    def paragraph(self, paragraph: ET.Element) -> dict[str, Any]:
        paragraph_properties = _child(paragraph, "pPr")
        style_node = _child(paragraph_properties, "pStyle")
        style_id = _attr(style_node, "val") or self.normal_style_id
        run = next(
            (item for item in paragraph.findall("w:r", NS) if _text(item)),
            paragraph.find("w:r", NS),
        )
        run_properties = _child(run, "rPr") if run is not None else None
        run_style_node = _child(run_properties, "rStyle")
        run_style_id = _attr(run_style_node, "val")
        effective_paragraph = _merge(
            self.default_paragraph,
            self._style_chain(self.normal_style_id, "paragraph"),
            self._style_chain(style_id, "paragraph"),
            _paragraph_properties(paragraph_properties),
        )
        effective_run = _merge(
            self.default_run,
            self._style_chain(self.normal_style_id, "run"),
            self._style_chain(style_id, "run"),
            self._style_chain(run_style_id, "run"),
            _run_properties(run_properties),
        )
        return {
            "text": _text(paragraph),
            "styleId": style_id,
            "styleName": self.styles.get(style_id or "", {}).get("name"),
            "paragraph": effective_paragraph,
            "run": effective_run,
        }


def _paragraph_role_samples(
    document_xml: ET.Element, resolver: _StyleResolver
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    table_paragraphs = set(document_xml.findall(".//w:tbl//w:p", NS))
    paragraphs = []
    for index, paragraph in enumerate(document_xml.findall(".//w:p", NS), start=1):
        sample = resolver.paragraph(paragraph)
        sample["index"] = index
        sample["inTable"] = paragraph in table_paragraphs
        paragraphs.append(sample)

    nonempty = [item for item in paragraphs if item["text"]]
    if not nonempty:
        empty = {
            "index": 0,
            "text": "",
            "styleId": resolver.normal_style_id,
            "styleName": "Normal",
            "paragraph": resolver.default_paragraph,
            "run": resolver.default_run,
        }
        return paragraphs, {
            "collegeTitle": empty,
            "reportTitle": empty,
            "heading1": empty,
            "heading2": empty,
            "heading3": empty,
            "body": empty,
            "caption": empty,
            "code": empty,
            "tableBody": empty,
        }

    title_candidates = [item for item in nonempty[:30] if not item["inTable"]] or nonempty[:30]
    report_title = max(
        title_candidates,
        key=lambda item: (
            int(item["run"].get("sizeHalfPoints") or 0),
            bool(re.search(r"报告|设计|实验", item["text"])),
            -item["index"],
        ),
    )
    college_title = next(
        (
            item
            for item in title_candidates
            if item["index"] < report_title["index"]
            and re.search(r"大学|学院|学校", item["text"])
        ),
        report_title,
    )
    named_heading_pattern = re.compile(
        r"^(?:第?[一二三四五六七八九十\d]+(?:章|节)?[.、．\s]*)?"
        r"(?:实验目的|实验环境|实验原理|任务要求|实验步骤|实验过程|实验结果|"
        r"问题分析|需求分析|系统设计|实现结果|设计总结|实验总结|总结|参考文献)"
        r"(?:$|[（(:：])"
    )
    heading_candidates = [
        item
        for item in nonempty
        if item["index"] != report_title["index"]
        and not item["inTable"]
        and len(item["text"]) <= 40
        and (
            item["run"].get("bold")
            or (
                item["paragraph"].get("outlineLevel") is not None
                and len(item["text"]) <= 24
            )
            or re.match(r"^[一二三四五六七八九十\d]+[.、．\s]", item["text"])
            or named_heading_pattern.match(item["text"])
        )
    ]
    heading_candidates.sort(
        key=lambda item: (
            -(item["run"].get("sizeHalfPoints") or 0),
            item["index"],
        )
    )
    heading1 = heading_candidates[0] if heading_candidates else report_title
    heading2 = heading_candidates[1] if len(heading_candidates) > 1 else heading1
    heading3 = heading_candidates[2] if len(heading_candidates) > 2 else heading2
    last_table_paragraph_index = max(
        (item["index"] for item in paragraphs if item["inTable"]),
        default=report_title["index"],
    )

    def body_candidate(*, after_index: int, in_table: bool, meaningful: bool):
        return next(
            (
                item
                for item in nonempty
                if item["index"] > after_index
                and bool(item["inTable"]) is in_table
                and item not in heading_candidates
                and (
                    not meaningful
                    or (
                        len(item["text"]) >= 18
                        and not re.fullmatch(r"[_＿.\s]{6,}", item["text"])
                        and not re.search(
                            r"姓名|学号|班级|课程名称|教师评语", item["text"][:12]
                        )
                    )
                )
            ),
            None,
        )

    body = (
        body_candidate(
            after_index=last_table_paragraph_index, in_table=False, meaningful=True
        )
        or body_candidate(
            after_index=report_title["index"], in_table=False, meaningful=True
        )
        or body_candidate(
            after_index=report_title["index"], in_table=True, meaningful=True
        )
        or body_candidate(
            after_index=last_table_paragraph_index, in_table=False, meaningful=False
        )
        or body_candidate(
            after_index=report_title["index"], in_table=False, meaningful=False
        )
        or nonempty[-1]
    )
    table_body = next(
        (
            item
            for item in nonempty
            if item["inTable"]
            and not re.fullmatch(r"[_＿.\s]{2,}", item["text"])
        ),
        body,
    )
    caption = next(
        (item for item in nonempty if re.match(r"^图\s*\d+", item["text"])),
        body,
    )
    code = next(
        (
            item
            for item in nonempty
            if re.search(r"\b(import|class|public|select|interface|ping|ipconfig)\b", item["text"], re.I)
        ),
        body,
    )

    return paragraphs, {
        "collegeTitle": college_title,
        "reportTitle": report_title,
        "heading1": heading1,
        "heading2": heading2,
        "heading3": heading3,
        "body": body,
        "caption": caption,
        "code": code,
        "tableBody": table_body,
    }


def _role_contract(sample: dict[str, Any]) -> dict[str, Any]:
    return {
        "source": {
            "paragraphIndex": sample["index"],
            "styleId": sample["styleId"],
            "styleName": sample["styleName"],
            "sampleText": sample["text"][:100],
        },
        **sample["run"],
        **sample["paragraph"],
    }


def _table_contract(table: ET.Element, index: int) -> dict[str, Any]:
    properties = _child(table, "tblPr")
    grid = [
        _integer(_attr(column, "w")) or 0
        for column in table.findall("w:tblGrid/w:gridCol", NS)
    ]
    rows = table.findall("w:tr", NS)
    borders_node = _child(properties, "tblBorders")
    borders = {}
    if borders_node is not None:
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            side_node = _child(borders_node, side)
            if side_node is not None:
                borders[side] = _attrs(side_node)
    cell_margins_node = _child(properties, "tblCellMar")
    cell_margins = {}
    if cell_margins_node is not None:
        for side in ("top", "left", "bottom", "right", "start", "end"):
            side_node = _child(cell_margins_node, side)
            if side_node is not None:
                cell_margins[side] = _attrs(side_node)
    merge_relationships = []
    for row_index, row in enumerate(rows, start=1):
        for cell_index, cell in enumerate(row.findall("w:tc", NS), start=1):
            cell_properties = _child(cell, "tcPr")
            grid_span = _integer(_attr(_child(cell_properties, "gridSpan"), "val"))
            vertical_merge = _attr(_child(cell_properties, "vMerge"), "val")
            if grid_span or vertical_merge is not None:
                merge_relationships.append(
                    {
                        "row": row_index,
                        "cell": cell_index,
                        "gridSpan": grid_span,
                        "verticalMerge": vertical_merge or "continue",
                    }
                )
    first_row_text = []
    if rows:
        first_row_text = [_text(cell) for cell in rows[0].findall("w:tc", NS)]
    return {
        "index": index,
        "width": _attrs(_child(properties, "tblW")),
        "alignment": _attr(_child(properties, "jc"), "val"),
        "indent": _attrs(_child(properties, "tblInd")),
        "layout": _attr(_child(properties, "tblLayout"), "type"),
        "gridColumnsTwips": grid,
        "columnCount": len(grid),
        "rowCount": len(rows),
        "firstRowText": first_row_text,
        "mergeRelationships": merge_relationships,
        "cellMargins": cell_margins,
        "borders": borders,
        "repeatHeader": any(
            row.find("w:trPr/w:tblHeader", NS) is not None for row in rows[:1]
        ),
        "allowRowBreak": not any(
            row.find("w:trPr/w:cantSplit", NS) is not None for row in rows
        ),
    }


def _metadata_table(tables: list[dict[str, Any]]) -> dict[str, Any] | None:
    if not tables:
        return None
    label_pattern = re.compile(r"姓名|学号|班级|课程|实验|教师|学院|专业")
    best = max(
        tables,
        key=lambda table: sum(
            1 for text in table["firstRowText"] if label_pattern.search(text or "")
        )
        + (3 if 2 <= table["columnCount"] <= 6 else 0),
    )
    narrow_columns = [
        index + 1 for index, width in enumerate(best["gridColumnsTwips"]) if width < 1000
    ]
    long_value_narrow = any(
        len(text) >= 6
        and index < len(best["gridColumnsTwips"])
        and best["gridColumnsTwips"][index] < 1800
        for index, text in enumerate(best["firstRowText"])
    )
    return {
        "tableIndex": best["index"],
        "columnCount": best["columnCount"],
        "gridColumnsTwips": best["gridColumnsTwips"],
        "narrowColumns": narrow_columns,
        "verticalTextRisk": bool(narrow_columns or long_value_narrow),
        "hasOuterBorder": all(
            side in best["borders"] and best["borders"][side].get("val") not in {None, "nil", "none"}
            for side in ("top", "left", "bottom", "right")
        ),
    }


def analyze_template(path: Path | str) -> dict[str, Any]:
    source_path = Path(path).resolve()
    if source_path.suffix.lower() != ".docx":
        raise ValueError(f"only .docx templates can be analyzed directly: {source_path}")
    data = source_path.read_bytes()
    source_hash = hashlib.sha256(data).hexdigest()
    with zipfile.ZipFile(source_path) as archive:
        names = set(archive.namelist())
        if "word/document.xml" not in names:
            raise ValueError(f"invalid DOCX: missing word/document.xml: {source_path}")
        document_xml = ET.fromstring(archive.read("word/document.xml"))
        styles_xml = (
            ET.fromstring(archive.read("word/styles.xml"))
            if "word/styles.xml" in names
            else None
        )
        settings_xml = (
            ET.fromstring(archive.read("word/settings.xml"))
            if "word/settings.xml" in names
            else None
        )

        resolver = _StyleResolver(styles_xml)
        paragraphs, roles = _paragraph_role_samples(document_xml, resolver)
        tables = [
            _table_contract(table, index)
            for index, table in enumerate(document_xml.findall(".//w:tbl", NS), start=1)
        ]
        sections = []
        for index, section in enumerate(document_xml.findall(".//w:sectPr", NS), start=1):
            size = _child(section, "pgSz")
            margins = _child(section, "pgMar")
            border_node = _child(section, "pgBorders")
            border_sides = {}
            if border_node is not None:
                for side in ("top", "left", "bottom", "right"):
                    side_node = _child(border_node, side)
                    if side_node is not None:
                        border_sides[side] = _attrs(side_node)
            sections.append(
                {
                    "index": index,
                    "pageSize": {
                        "widthTwips": _integer(_attr(size, "w")),
                        "heightTwips": _integer(_attr(size, "h")),
                        "orientation": _attr(size, "orient") or "portrait",
                    },
                    "margins": {
                        key: _integer(_attr(margins, key))
                        for key in ("top", "right", "bottom", "left", "header", "footer", "gutter")
                    },
                    "breakType": _attr(_child(section, "type"), "val") or "nextPage",
                    "differentFirstPage": _child(section, "titlePg") is not None,
                    "differentOddEven": (
                        settings_xml is not None
                        and settings_xml.find("w:evenAndOddHeaders", NS) is not None
                    ),
                    "pageBorder": None
                    if border_node is None
                    else {
                        "position": _attr(border_node, "offsetFrom"),
                        "display": _attr(border_node, "display"),
                        "sides": border_sides,
                    },
                    "headerReferences": [
                        _attrs(item) for item in section.findall("w:headerReference", NS)
                    ],
                    "footerReferences": [
                        _attrs(item) for item in section.findall("w:footerReference", NS)
                    ],
                }
            )

        placeholders = []
        headings = []
        for paragraph in paragraphs:
            if re.search(r"\[\[\s*IMAGE|图片占位|插入图片|截图位置", paragraph["text"], re.I):
                placeholders.append(
                    {
                        "paragraphIndex": paragraph["index"],
                        "text": paragraph["text"],
                    }
                )
            if paragraph in roles.values() or not paragraph["text"]:
                continue
            if len(paragraph["text"]) <= 50 and (
                paragraph["run"].get("bold")
                or paragraph["paragraph"].get("outlineLevel") is not None
            ):
                headings.append(
                    {
                        "paragraphIndex": paragraph["index"],
                        "text": paragraph["text"],
                        "styleId": paragraph["styleId"],
                    }
                )

        metadata = _metadata_table(tables)
        all_text = "\n".join(item["text"] for item in paragraphs if item["text"])
        top_level_paragraphs = document_xml.findall("./w:body/w:p", NS)
        opening_title_text = _text(top_level_paragraphs[0]) if top_level_paragraphs else ""
        blank_template = len(all_text.strip()) < 12 or not opening_title_text
        risk_reasons = []
        risk_score = 0
        if blank_template:
            risk_reasons.append("blank-template")
            risk_score += 45
        if metadata and not metadata["hasOuterBorder"]:
            risk_reasons.append("metadata-table-without-outer-border")
            risk_score += 25
        if metadata and metadata["verticalTextRisk"]:
            risk_reasons.append("metadata-column-vertical-text-risk")
            risk_score += 30
        if len(sections) > 1:
            risk_reasons.append("multi-section-template")
            risk_score += 15
        if placeholders:
            risk_reasons.append("image-placeholders-present")
            risk_score += 5
        if styles_xml is None:
            risk_reasons.append("styles-part-missing")
            risk_score += 35

        return {
            "schemaVersion": "1.0",
            "analyzerVersion": ANALYZER_VERSION,
            "source": {
                "path": str(source_path),
                "sha256": source_hash,
                "sizeBytes": len(data),
                "packagePartCount": len(names),
            },
            "page": {"sections": sections},
            "styles": {
                "defaults": {
                    "paragraph": resolver.default_paragraph,
                    "run": resolver.default_run,
                    "normalStyleId": resolver.normal_style_id,
                },
                "roles": {name: _role_contract(sample) for name, sample in roles.items()},
            },
            "structure": {
                "titleArea": roles["reportTitle"]["index"],
                "collegeName": (
                    roles["collegeTitle"]["text"]
                    if roles["collegeTitle"]["index"] != roles["reportTitle"]["index"]
                    else None
                ),
                "metadataTable": metadata,
                "bodyStartParagraph": roles["body"]["index"],
                "existingHeadings": headings,
                "imagePlaceholders": placeholders,
                "sectionCount": len(sections),
                "headerPartCount": len([name for name in names if re.match(r"word/header\d+\.xml", name)]),
                "footerPartCount": len([name for name in names if re.match(r"word/footer\d+\.xml", name)]),
            },
            "tables": tables,
            "images": {
                "embeddedCount": len([name for name in names if name.startswith("word/media/")])
            },
            "risk": {
                "score": min(risk_score, 100),
                "level": "high" if risk_score >= 50 else "medium" if risk_score >= 20 else "low",
                "reasons": risk_reasons,
            },
        }


def analyze_template_cached(
    path: Path | str, cache_dir: Path | str
) -> dict[str, Any]:
    source_path = Path(path).resolve()
    cache_root = Path(cache_dir).resolve()
    cache_root.mkdir(parents=True, exist_ok=True)
    source_digest = hashlib.sha256(source_path.read_bytes()).hexdigest()
    cache_key = f"{ANALYZER_VERSION}-{source_digest}"
    cache_path = cache_root / f"{cache_key}.template-style-contract.json"
    if cache_path.exists():
        payload = json.loads(cache_path.read_text(encoding="utf-8-sig"))
        payload["cache"] = {"hit": True, "path": str(cache_path), "key": cache_key}
        return payload

    payload = analyze_template(source_path)
    payload["cache"] = {"hit": False, "path": str(cache_path), "key": cache_key}
    cache_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8-sig",
    )
    return payload


def recommend_quality_mode(
    contract: dict[str, Any], *, verified: bool = False
) -> dict[str, Any]:
    risk = contract.get("risk", {})
    score = int(risk.get("score") or 0)
    reasons = list(risk.get("reasons") or [])
    if verified and score < 50:
        return {
            "recommendedMode": "fast",
            "riskScore": score,
            "riskLevel": risk.get("level", "low"),
            "reasons": reasons + ["verified-template-cache"],
        }
    return {
        "recommendedMode": "strict" if score >= 20 else "fast",
        "riskScore": score,
        "riskLevel": risk.get("level", "low"),
        "reasons": reasons,
    }
