from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path

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


if __name__ == "__main__":
    unittest.main()
