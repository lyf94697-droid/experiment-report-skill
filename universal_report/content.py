from __future__ import annotations

import hashlib
from typing import Any


BASE_SECTIONS = [
    "实验目的",
    "实验环境",
    "实验原理",
    "实验内容",
    "实验步骤",
    "关键代码、命令或配置",
    "实验结果",
    "结果分析",
    "问题与解决方法",
    "实验总结与心得",
]

COURSE_SECTIONS = {
    "network": ["网络拓扑与地址规划", "设备或主机配置", "连通性与协议验证"],
    "operating-system": ["进程状态与调度依据", "系统调用或核心命令", "资源状态与运行现象"],
    "database": ["数据模型与表结构", "SQL 实现与事务过程", "查询结果与约束验证"],
    "programming": ["需求与模块划分", "关键实现与代码说明", "测试用例与运行结果"],
    "android": ["界面与组件设计", "生命周期与数据流", "设备或模拟器验证"],
    "software-engineering": ["需求分析", "系统设计", "实现与测试", "缺陷与改进"],
}


def _course_family(course_name: str) -> str:
    normalized = course_name.lower()
    if "网络" in course_name or "network" in normalized:
        return "network"
    if "操作系统" in course_name or "operating system" in normalized:
        return "operating-system"
    if "数据库" in course_name or "database" in normalized or "sql" in normalized:
        return "database"
    if "android" in normalized or "安卓" in course_name:
        return "android"
    if "软件工程" in course_name:
        return "software-engineering"
    return "programming"


def build_report_plan(
    *,
    course_name: str,
    experiment_name: str,
    detail_level: str = "standard",
    variant_seed: str = "",
) -> dict[str, Any]:
    family = _course_family(course_name)
    course_sections = COURSE_SECTIONS[family]
    ordered_titles = BASE_SECTIONS[:4] + course_sections + BASE_SECTIONS[4:]
    seen: set[str] = set()
    sections = []
    for index, title in enumerate(ordered_titles, start=1):
        if title in seen:
            continue
        seen.add(title)
        sections.append(
            {
                "id": f"section-{len(sections) + 1}",
                "title": title,
                "order": len(sections) + 1,
                "evidenceRequired": title in {"实验步骤", "实验结果", "结果分析"},
            }
        )

    variants = ["过程导向", "验证导向", "问题导向", "原理导向"]
    variant_source = variant_seed or experiment_name
    digest = hashlib.sha256(variant_source.encode("utf-8")).hexdigest()
    variant_index = (
        sum((index + 1) * ord(character) for index, character in enumerate(variant_source))
        + int(digest[-2:], 16)
    ) % len(variants)
    writing_variant = variants[variant_index]
    long_mode = detail_level.lower() in {"long", "full"}

    return {
        "schemaVersion": "1.0",
        "courseName": course_name,
        "experimentName": experiment_name,
        "courseFamily": family,
        "detailLevel": "long" if long_mode else "standard",
        "targetCharacters": 2600 if long_mode else 1200,
        "writingVariant": writing_variant,
        "sections": sections,
        "factualityRules": [
            "截图、代码、命令和运行结果优先作为事实证据",
            "材料中没有出现的成功结果不得编造",
            "参考教程仅用于解释步骤，不能替代用户实际结果",
        ],
    }
