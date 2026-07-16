from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from .content import build_report_plan
from .format_validation import validate_format
from .images import build_image_manifest
from .template_contract import (
    analyze_template,
    analyze_template_cached,
    recommend_quality_mode,
)
from .template_catalog import (
    audit_builtin_templates,
    load_template_catalog,
    resolve_template_selection,
)
from .visual_validation import inspect_pdf


def _write(payload: dict[str, Any], output: str | None) -> None:
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    if output:
        path = Path(output)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8-sig")
    else:
        print(text)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="universal-report")
    subparsers = parser.add_subparsers(dest="command", required=True)

    analyze = subparsers.add_parser("analyze-template")
    analyze.add_argument("template")
    analyze.add_argument("--cache-dir")
    analyze.add_argument("--verified", action="store_true")
    analyze.add_argument("--output")

    validate = subparsers.add_parser("validate-format")
    validate.add_argument("template")
    validate.add_argument("document")
    validate.add_argument("--output")

    images = subparsers.add_parser("image-manifest")
    images.add_argument("images", nargs="+")
    images.add_argument("--count", type=int)
    images.add_argument("--allow-grid", action="store_true")
    images.add_argument("--output")

    plan = subparsers.add_parser("plan-content")
    plan.add_argument("--course", required=True)
    plan.add_argument("--experiment", required=True)
    plan.add_argument("--detail", choices=["standard", "long"], default="standard")
    plan.add_argument("--variant-seed", default="")
    plan.add_argument("--output")

    visual = subparsers.add_parser("visual-validate")
    visual.add_argument("pdf")
    visual.add_argument("--preview-dir", required=True)
    visual.add_argument("--require-closed-frame", action="store_true")
    visual.add_argument("--dpi", type=int, default=144)
    visual.add_argument("--output")

    catalog = subparsers.add_parser("list-templates")
    catalog.add_argument("--repo-root", default=".")
    catalog.add_argument("--output")

    select = subparsers.add_parser("select-template")
    select.add_argument("--repo-root", default=".")
    select.add_argument("--report-type", default="experiment-report")
    select.add_argument("--course", default="")
    select.add_argument("--preference", default="auto")
    select.add_argument("--request", default="")
    select.add_argument("--user-template")
    select.add_argument("--output")

    audit = subparsers.add_parser("audit-template-catalog")
    audit.add_argument("--repo-root", default=".")
    audit.add_argument("--output")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.command == "analyze-template":
        payload = (
            analyze_template_cached(args.template, args.cache_dir)
            if args.cache_dir
            else analyze_template(args.template)
        )
        payload["qualityRecommendation"] = recommend_quality_mode(
            payload, verified=args.verified
        )
        _write(payload, args.output)
        return 0
    if args.command == "validate-format":
        _write(validate_format(args.template, args.document), args.output)
        return 0
    if args.command == "image-manifest":
        _write(
            build_image_manifest(
                args.images,
                requested_count=args.count,
                allow_grid=args.allow_grid,
            ),
            args.output,
        )
        return 0
    if args.command == "plan-content":
        _write(
            build_report_plan(
                course_name=args.course,
                experiment_name=args.experiment,
                detail_level=args.detail,
                variant_seed=args.variant_seed,
            ),
            args.output,
        )
        return 0
    if args.command == "visual-validate":
        _write(
            inspect_pdf(
                args.pdf,
                output_dir=args.preview_dir,
                require_closed_frame=args.require_closed_frame,
                dpi=args.dpi,
            ),
            args.output,
        )
        return 0
    if args.command == "list-templates":
        catalog = load_template_catalog(args.repo_root)
        _write(
            {
                "schemaVersion": catalog["schemaVersion"],
                "catalogPath": catalog["catalogPath"],
                "templateQuestion": catalog["templateQuestion"],
                "templates": catalog["templates"],
            },
            args.output,
        )
        return 0
    if args.command == "select-template":
        _write(
            resolve_template_selection(
                repo_root=args.repo_root,
                report_type=args.report_type,
                user_template=args.user_template,
                course_name=args.course,
                preference=args.preference,
                request_text=args.request,
            ),
            args.output,
        )
        return 0
    if args.command == "audit-template-catalog":
        payload = audit_builtin_templates(args.repo_root)
        _write(payload, args.output)
        return 0 if payload["passed"] else 1
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
