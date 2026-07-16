from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any


def _first_value(*values: str | None) -> str | None:
    for value in values:
        if value and value.strip():
            return value.strip()
    return None


def load_config(repo_root: Path, config_path: Path | None = None) -> dict[str, Any]:
    repo_root = Path(repo_root)
    file_config: dict[str, Any] = {}
    candidate = config_path or (
        Path(os.environ["EXPERIMENT_REPORT_CONFIG"])
        if os.environ.get("EXPERIMENT_REPORT_CONFIG")
        else None
    )
    if candidate and candidate.exists():
        file_config = json.loads(candidate.read_text(encoding="utf-8-sig"))

    default_template = _first_value(
        os.environ.get("EXPERIMENT_REPORT_TEMPLATE_PATH"),
        str(file_config.get("defaultTemplate") or ""),
        str(repo_root / "examples" / "report-templates" / "experiment-report-template.docx"),
    )
    output_root = _first_value(
        os.environ.get("EXPERIMENT_REPORT_OUTPUT_ROOT"),
        str(file_config.get("outputRoot") or ""),
        str(repo_root / "outputs"),
    )
    cache_root = _first_value(
        os.environ.get("EXPERIMENT_REPORT_CACHE_ROOT"),
        str(file_config.get("cacheRoot") or ""),
        str(Path.home() / ".cache" / "experiment-report" / "templates"),
    )

    return {
        "schemaVersion": "1.0",
        "defaultTemplate": default_template,
        "outputRoot": output_root,
        "cacheRoot": cache_root,
        "pdf": {
            "preferLibreOffice": True,
            "allowOfficeCom": os.environ.get("EXPERIMENT_REPORT_ALLOW_OFFICE_COM") == "1",
        },
    }
