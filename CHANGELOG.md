# Changelog

All notable changes to this repository will be documented in this file.

The format is based on Keep a Changelog, and this project currently tracks changes under a rolling `Unreleased` section until the first tagged release.

## [Unreleased]

### Added

- Added ten original, school-neutral DOCX templates covering classic, bordered, engineering, course-design, modern-minimal, compact-header, review-panel, code-notebook, data-analysis, and project-dossier report styles, together with a machine-readable catalog and reproducible builder.
- Added template selection and catalog-audit commands. User-supplied templates always win; built-ins are checked for institution names, example identities, embedded media, provenance, and license metadata.
- Added a persistent ten-template integration test that builds every neutral template through the main report pipeline, plus a five-family uploaded-template fidelity corpus.
- Added the `universal_report` core package for cross-machine configuration, structured pipeline traces, content planning, template contracts, image manifests, format validation, and strict visual validation.
- Added `TemplateStyleContract` analysis with effective style inheritance, template-hash/version caching, page and section settings, role-based typography, table geometry, headers, footers, placeholders, and structural risks.
- Added ten generated compatibility fixtures, Python unit tests, core smoke tests, five universal end-to-end scenarios, and real-template regression coverage.
- Added `docs/architecture.md`, `docs/troubleshooting.md`, and `docs/compatibility.md`.
- Added PowerShell entry points for template analysis, format validation, content planning, strict visual validation, legacy template conversion, and fixture generation.
- Added `scripts/run-one-click-demo.ps1` so the repository now has a deterministic demo that can be run without preparing external templates or screenshots.
- Added documentation indexes under `docs/`, `examples/`, `scripts/`, `profiles/`, and `references/` to make the repository easier to navigate on first open.
- Added `docs/one-click-demo.md` and `docs/social-launch-kit.md` to cover the demo flow and public-facing launch materials.
- Added `examples/demo-one-click/` and `examples/report-templates/experiment-report-template.docx` as a self-contained onboarding bundle.
- Added `scripts/build-report.ps1` as a single local entry point for validation, field-map generation, template filling, image insertion, and final style formatting.
- Added `scripts/report-defaults.ps1` so generated runs can remember the last course name and experiment name.
- Added `CONTRIBUTING.md` with repository workflow, testing expectations, and contribution scope.
- Added `CODE_OF_CONDUCT.md`, `SECURITY.md`, `SUPPORT.md`, and `ROADMAP.md` for GitHub-facing repository completeness.
- Added issue templates, a PR template, and a matrix quality workflow under `.github/`.
- Added `demo/` assets and a demo guide suitable for GitHub repository previews.
- Added `.gitattributes` to keep Office files and demo images treated as binary content.

### Changed

- The CLI, PowerShell wrappers, Web UI, documentation, and skill instructions now ask about a preferred template once and otherwise select from exactly ten neutral built-in templates.
- The former `experiment-report-template.docx` and `course-design-report-template.docx` files are now compatibility aliases for the classic laboratory and course-design templates.
- User-supplied templates now default to fidelity-preserving mode; repository style normalization and course-design enhancements run only when explicitly requested or when using a repository default template.
- The main wrappers and Web UI now expose structured progress, generation status, template/format/image/visual artifacts, exact image-count requests, and actionable quality recommendations.
- Image handling now validates paths, deduplicates by content hash, records selection reasons, defaults to one image per line, and keeps image paragraphs with their captions where practical.
- Strict mode now requires DOCX-to-PDF conversion, per-page preview rendering, and visual checks; incomplete validation returns `needs-fix`.
- LibreOffice is preferred for conversion, while WPS/Microsoft Word COM is opt-in with timeout and scoped process cleanup.
- Reworked `README.md` into a Chinese-first project homepage with clear positioning, quick-start commands, directory overview, and documentation navigation.
- Refreshed the demo and example documentation so the repo reads like a complete open-source project instead of a loose collection of helper scripts.
- Updated the example image JSON files to reference repo-contained demo assets instead of non-existent placeholder paths.
- Expanded `README.md` with a quick-start build flow, demo links, and contributor-oriented repository structure notes.
- Expanded `README.md` with repository health notes and a future profile-driven document roadmap.
- Expanded smoke tests to cover the new local build entry point and required repository files.
- Improved the final docx style formatter so body table rows can flow more naturally in common report templates instead of preserving awkward row-splitting constraints.
- Improved direct-chat image handling so OpenClaw-staged relative attachment paths such as `media/inbound/example.png` can resolve into the final docx image pipeline.

### Fixed

- Fixed block field filling so section body text no longer inherits a locked heading paragraph style.
- Fixed template analyzer role sampling for table-heavy and cover/body templates, and versioned cache invalidation when analyzer rules change.
- Fixed format validation so risks already present in the source template are not reported as newly introduced drift.
- Fixed `install-skill.ps1` so editor metadata and Python cache artifacts are not copied into installed skill directories.
- Fixed repository hygiene by ignoring Python cache artifacts such as `__pycache__/` and `*.pyc`.
- Fixed the report-style formatter so it no longer leaks XML attribute return values into the PowerShell pipeline.
- Fixed title detection for common report titles such as `计算机网络实验报告`.
