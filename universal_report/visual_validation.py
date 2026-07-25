from __future__ import annotations

import shutil
import subprocess
from pathlib import Path
from typing import Any


def _long_dark_line(image, *, horizontal: bool, start: int, end: int) -> bool:
    grayscale = image.convert("L")
    width, height = grayscale.size
    pixels = grayscale.load()
    if horizontal:
        x_start = int(width * 0.04)
        x_end = max(x_start + 1, int(width * 0.96))
        for y in range(max(0, start), min(height, end)):
            dark = sum(1 for x in range(x_start, x_end) if pixels[x, y] < 100)
            if dark / max(1, x_end - x_start) >= 0.62:
                return True
        return False

    y_start = int(height * 0.04)
    y_end = max(y_start + 1, int(height * 0.96))
    for x in range(max(0, start), min(width, end)):
        dark = sum(1 for y in range(y_start, y_end) if pixels[x, y] < 100)
        if dark / max(1, y_end - y_start) >= 0.62:
            return True
    return False


def _page_metrics(image, *, require_closed_frame: bool) -> dict[str, Any]:
    grayscale = image.convert("L")
    width, height = grayscale.size
    histogram = grayscale.histogram()
    nonwhite = sum(histogram[:245])
    nonwhite_ratio = nonwhite / max(1, width * height)
    top = _long_dark_line(
        image,
        horizontal=True,
        start=int(height * 0.02),
        end=int(height * 0.16),
    )
    bottom = _long_dark_line(
        image,
        horizontal=True,
        start=int(height * 0.84),
        end=int(height * 0.98),
    )
    left = _long_dark_line(
        image,
        horizontal=False,
        start=int(width * 0.02),
        end=int(width * 0.16),
    )
    right = _long_dark_line(
        image,
        horizontal=False,
        start=int(width * 0.84),
        end=int(width * 0.98),
    )
    closed_frame = top and bottom and left and right
    not_blank = nonwhite_ratio >= 0.0025
    passed = not_blank and (closed_frame if require_closed_frame else True)
    return {
        "passed": passed,
        "nonWhiteRatio": round(nonwhite_ratio, 6),
        "checks": {
            "notBlank": not_blank,
            "closedFrameRequired": require_closed_frame,
            "closedFrame": closed_frame,
            "frameSides": {
                "top": top,
                "bottom": bottom,
                "left": left,
                "right": right,
            },
        },
    }


def inspect_pdf(
    pdf_path: Path | str,
    *,
    output_dir: Path | str,
    require_closed_frame: bool = False,
    dpi: int = 144,
) -> dict[str, Any]:
    try:
        from PIL import Image
    except ImportError:
        return {
            "schemaVersion": "1.0",
            "status": "needs-fix",
            "passed": False,
            "pageCount": 0,
            "pages": [],
            "errors": [
                {
                    "code": "visual-render-dependency-missing",
                    "message": "缺少 Pillow，无法执行严格视觉验收。",
                    "suggestion": "使用工作区自带的 Python 运行，或安装 requirements-web.txt 后重试。",
                }
            ],
        }

    source = Path(pdf_path).resolve()
    target_dir = Path(output_dir).resolve()
    target_dir.mkdir(parents=True, exist_ok=True)
    pages = []
    try:
        import fitz
    except ImportError:
        fitz = None

    if fitz is not None:
        document = fitz.open(str(source))
        try:
            scale = max(1.0, dpi / 72.0)
            for page_index in range(document.page_count):
                page = document.load_page(page_index)
                pixmap = page.get_pixmap(
                    matrix=fitz.Matrix(scale, scale),
                    alpha=False,
                )
                image = Image.frombytes(
                    "RGB",
                    (pixmap.width, pixmap.height),
                    pixmap.samples,
                )
                preview_path = target_dir / f"page-{page_index + 1}.png"
                image.save(preview_path)
                metrics = _page_metrics(
                    image,
                    require_closed_frame=require_closed_frame,
                )
                pages.append(
                    {
                        "page": page_index + 1,
                        "previewPath": str(preview_path),
                        "width": pixmap.width,
                        "height": pixmap.height,
                        **metrics,
                    }
                )
        finally:
            document.close()
    else:
        pdftoppm = shutil.which("pdftoppm")
        if pdftoppm is None:
            return {
                "schemaVersion": "1.0",
                "status": "needs-fix",
                "passed": False,
                "pageCount": 0,
                "pages": [],
                "errors": [
                    {
                        "code": "visual-render-dependency-missing",
                        "message": "缺少 PyMuPDF 和 pdftoppm，无法把 PDF 渲染为逐页预览图。",
                        "suggestion": "使用工作区依赖中的 pdftoppm，或安装 requirements-web.txt 后重试。",
                    }
                ],
            }
        pdftoppm_path = Path(pdftoppm).resolve()
        if pdftoppm_path.suffix.lower() in {".cmd", ".bat"}:
            dependency_root = pdftoppm_path.parents[2]
            native_executable = (
                dependency_root
                / "native"
                / "poppler"
                / "Library"
                / "bin"
                / "pdftoppm.exe"
            )
            if native_executable.is_file():
                pdftoppm = str(native_executable)
        prefix = target_dir / "page"
        completed = subprocess.run(
            [
                pdftoppm,
                "-png",
                "-r",
                str(max(72, dpi)),
                str(source),
                str(prefix),
            ],
            check=False,
            capture_output=True,
            text=True,
        )
        if completed.returncode != 0:
            return {
                "schemaVersion": "1.0",
                "status": "needs-fix",
                "passed": False,
                "pageCount": 0,
                "pages": [],
                "errors": [
                    {
                        "code": "pdf-preview-render-failed",
                        "message": completed.stderr.strip()
                        or "pdftoppm 未能生成逐页预览图。",
                        "suggestion": "检查 PDF 文件是否完整，以及 pdftoppm 是否可以运行。",
                    }
                ],
            }
        preview_paths = sorted(
            target_dir.glob("page-*.png"),
            key=lambda path: int(path.stem.rsplit("-", 1)[-1]),
        )
        for page_index, preview_path in enumerate(preview_paths, start=1):
            with Image.open(preview_path) as source_image:
                image = source_image.convert("RGB")
                metrics = _page_metrics(
                    image,
                    require_closed_frame=require_closed_frame,
                )
                pages.append(
                    {
                        "page": page_index,
                        "previewPath": str(preview_path),
                        "width": image.width,
                        "height": image.height,
                        **metrics,
                    }
                )

    failed_pages = [item["page"] for item in pages if not item["passed"]]
    passed = bool(pages) and not failed_pages
    errors = []
    if not pages:
        errors.append(
            {
                "code": "pdf-has-no-pages",
                "message": "PDF 没有可渲染页面。",
                "suggestion": "检查 DOCX 到 PDF 的转换结果。",
            }
        )
    if failed_pages:
        errors.append(
            {
                "code": "visual-page-check-failed",
                "message": f"第 {', '.join(map(str, failed_pages))} 页未通过视觉检查。",
                "suggestion": "检查空白页、页面外框和内容是否被裁切，再重新生成。",
            }
        )

    return {
        "schemaVersion": "1.0",
        "pdfPath": str(source),
        "previewDir": str(target_dir),
        "status": "completed" if passed else "needs-fix",
        "passed": passed,
        "pageCount": len(pages),
        "failedPages": failed_pages,
        "pages": pages,
        "errors": errors,
    }
