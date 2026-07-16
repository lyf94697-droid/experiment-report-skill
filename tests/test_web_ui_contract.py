from __future__ import annotations

import re
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]


class WebUiContractTests(unittest.TestCase):
    def test_web_ui_uses_cross_machine_config_and_preserves_uploaded_templates(self) -> None:
        source = (REPO_ROOT / "web_ui.py").read_text(encoding="utf-8")

        self.assertIn("load_config(REPO_ROOT)", source)
        self.assertIn("load_template_catalog(REPO_ROOT)", source)
        self.assertIn("built_in_template_choice", source)
        self.assertIn("用户上传模板始终优先", source)
        self.assertIn('"TemplateStyleMode": "preserve"', source)
        self.assertIn('"-TemplateStyleMode"', source)
        self.assertNotIn(r"E:\实验报告", source)
        self.assertNotIn(r"E:\新建文件夹", source)
        self.assertNotIn('"-StyleFinalDocx"', source)

    def test_wrappers_only_normalize_when_explicitly_requested(self) -> None:
        for script_name in ("build-report-from-url.ps1", "build-report-from-feishu.ps1"):
            source = (REPO_ROOT / "scripts" / script_name).read_text(encoding="utf-8-sig")
            self.assertIn('[ValidateSet("preserve", "normalize")]', source)
            self.assertIn("$templatePathDefaulted -and", source)
            self.assertNotIn("    StyleFinalDocx = $true", source)
            self.assertRegex(
                source,
                re.compile(
                    r'Equals\(\$TemplateStyleMode,\s*"normalize".+?StyleFinalDocx\s*=\s*\$true',
                    re.DOTALL,
                ),
            )

    def test_office_com_is_explicit_and_timeout_guarded(self) -> None:
        source = (REPO_ROOT / "scripts" / "export-docx-pdf.ps1").read_text(
            encoding="utf-8-sig"
        )
        self.assertIn("[switch]$AllowOfficeCom", source)
        self.assertIn("$OfficeComTimeoutSeconds", source)
        self.assertIn("Wait-Job", source)
        self.assertIn("Stop-NewOfficeProcesses", source)


if __name__ == "__main__":
    unittest.main()
