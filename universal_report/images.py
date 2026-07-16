from __future__ import annotations

import hashlib
import re
from pathlib import Path
from typing import Any, Iterable


def _section_for_name(name: str) -> str:
    lowered = name.lower()
    if re.search(r"result|ping|output|success|结果|验证|运行", lowered):
        return "实验结果"
    if re.search(r"error|problem|analysis|异常|报错|问题", lowered):
        return "问题与解决方法"
    if re.search(r"environment|install|setup|topology|拓扑|环境|安装", lowered):
        return "实验环境"
    return "实验步骤"


def _caption(section: str, index: int, stem: str) -> str:
    cleaned = re.sub(r"^[\d._\-\s]+", "", stem).strip()
    cleaned = re.sub(r"[_\-]+", " ", cleaned)
    fallback = {
        "实验环境": "实验环境与准备情况",
        "实验步骤": "关键操作过程",
        "实验结果": "实验运行与验证结果",
        "问题与解决方法": "问题现象与排查过程",
    }[section]
    return f"图{index} {cleaned or fallback}"


def build_image_manifest(
    paths: Iterable[Path | str],
    *,
    requested_count: int | None = None,
    allow_grid: bool = False,
) -> dict[str, Any]:
    unique: list[tuple[Path, str, int]] = []
    seen_hashes: set[str] = set()
    duplicates = 0
    rejected: list[dict[str, str]] = []

    for raw_path in paths:
        path = Path(raw_path)
        if not path.exists() or not path.is_file():
            rejected.append({"path": str(path), "reason": "file-not-found"})
            continue
        data = path.read_bytes()
        digest = hashlib.sha256(data).hexdigest()
        if digest in seen_hashes:
            duplicates += 1
            rejected.append({"path": str(path), "reason": "duplicate-content"})
            continue
        seen_hashes.add(digest)
        score = 20
        if re.search(r"result|ping|output|结果|验证", path.name, re.IGNORECASE):
            score += 20
        if re.search(r"setup|step|config|步骤|配置|拓扑", path.name, re.IGNORECASE):
            score += 15
        score += min(path.stat().st_size // 1024, 30)
        unique.append((path.resolve(), digest, score))

    unique.sort(key=lambda item: (-item[2], item[0].name.lower()))
    if requested_count is not None:
        if requested_count < 0:
            raise ValueError("requested_count must be non-negative")
        if len(unique) < requested_count:
            raise ValueError(
                f"requested {requested_count} images, but only {len(unique)} usable images remain"
            )
        unique = unique[:requested_count]

    images: list[dict[str, Any]] = []
    for index, (path, digest, score) in enumerate(unique, start=1):
        section = _section_for_name(path.name)
        images.append(
            {
                "path": str(path),
                "sha256": digest,
                "section": section,
                "caption": _caption(section, index, path.stem),
                "order": index,
                "selectionScore": score,
                "selectionReason": (
                    f"按文件名语义映射到“{section}”，并在去重后按证据相关性排序"
                ),
                "layout": {
                    "mode": "grid" if allow_grid else "single",
                    "columns": 2 if allow_grid else 1,
                    "preserveAspectRatio": True,
                    "captionBelow": True,
                    "keepWithCaption": True,
                },
            }
        )

    return {
        "schemaVersion": "1.0",
        "requestedCount": requested_count,
        "selectedCount": len(images),
        "duplicatesFiltered": duplicates,
        "rejected": rejected,
        "images": images,
    }
