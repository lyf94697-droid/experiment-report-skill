from __future__ import annotations

import importlib.util
import tempfile
import unittest
from types import SimpleNamespace
from pathlib import Path
from unittest.mock import patch

from universal_report.visual_validation import inspect_pdf


@unittest.skipUnless(
    importlib.util.find_spec("fitz") and importlib.util.find_spec("PIL"),
    "PyMuPDF and Pillow are optional visual-validation dependencies",
)
class VisualValidationTests(unittest.TestCase):
    def test_pdf_is_rendered_per_page_and_closed_frame_is_detected(self) -> None:
        import fitz

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            pdf_path = root / "framed.pdf"
            document = fitz.open()
            for page_number in range(2):
                page = document.new_page(width=595, height=842)
                page.draw_rect(fitz.Rect(30, 30, 565, 812), color=(0, 0, 0), width=1.5)
                page.insert_text((72, 90), f"Experiment report page {page_number + 1}")
            document.save(pdf_path)
            document.close()

            result = inspect_pdf(
                pdf_path,
                output_dir=root / "preview",
                require_closed_frame=True,
            )

            self.assertTrue(result["passed"])
            self.assertEqual(result["pageCount"], 2)
            self.assertEqual(len(result["pages"]), 2)
            self.assertTrue(all(Path(item["previewPath"]).exists() for item in result["pages"]))
            self.assertTrue(all(item["checks"]["closedFrame"] for item in result["pages"]))


@unittest.skipUnless(
    importlib.util.find_spec("PIL"),
    "Pillow is an optional visual-validation dependency",
)
class VisualValidationFallbackTests(unittest.TestCase):
    def test_pdftoppm_fallback_renders_pages_when_pymupdf_is_missing(self) -> None:
        from PIL import Image, ImageDraw

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            pdf_path = root / "fallback.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")
            original_import = __import__

            def import_without_fitz(name, *args, **kwargs):
                if name == "fitz":
                    raise ImportError("fitz deliberately unavailable")
                return original_import(name, *args, **kwargs)

            def fake_pdftoppm(command, **kwargs):
                prefix = Path(command[-1])
                image = Image.new("RGB", (600, 800), "white")
                drawer = ImageDraw.Draw(image)
                drawer.rectangle((80, 80, 520, 720), outline="black", width=4)
                drawer.text((120, 140), "Fallback preview", fill="black")
                image.save(prefix.parent / f"{prefix.name}-1.png")
                return SimpleNamespace(returncode=0, stderr="")

            with (
                patch("builtins.__import__", side_effect=import_without_fitz),
                patch(
                    "universal_report.visual_validation.shutil.which",
                    return_value=str(root / "pdftoppm"),
                ),
                patch(
                    "universal_report.visual_validation.subprocess.run",
                    side_effect=fake_pdftoppm,
                ),
            ):
                result = inspect_pdf(
                    pdf_path,
                    output_dir=root / "preview",
                )

            self.assertTrue(result["passed"])
            self.assertEqual(result["pageCount"], 1)
            self.assertTrue(Path(result["pages"][0]["previewPath"]).exists())


if __name__ == "__main__":
    unittest.main()
