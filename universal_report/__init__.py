"""Core building blocks for the universal experiment-report pipeline."""

from .template_contract import analyze_template, analyze_template_cached
from .template_catalog import (
    audit_builtin_templates,
    load_template_catalog,
    recommend_builtin_template,
    resolve_template_selection,
    scan_template_identity,
)

__all__ = [
    "analyze_template",
    "analyze_template_cached",
    "audit_builtin_templates",
    "load_template_catalog",
    "recommend_builtin_template",
    "resolve_template_selection",
    "scan_template_identity",
]
