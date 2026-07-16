"""Core building blocks for the universal experiment-report pipeline."""

from .template_contract import analyze_template, analyze_template_cached

__all__ = ["analyze_template", "analyze_template_cached"]
