# 脚本目录说明

本目录是一套“材料到可交付 DOCX”的本地流水线。用户模板默认保真；统一样式、Office COM 和多图并排都需要显式启用。

## 主入口

- `build-report.ps1`
  已有正文时的主入口。执行材料分析、内容计划、模板契约、字段填充、图片清单、DOCX 渲染和三类验证。
- `build-report-from-url.ps1`
  教程链接或参考文本入口，正文生成后进入同一主流水线。
- `build-report-from-feishu.ps1`
  直聊、附件路径和本地归档入口。
- `run-one-click-demo.ps1`
  使用仓库自带材料运行确定性演示。

用户模板推荐参数：

```powershell
-TemplateStyleMode preserve
```

没有用户模板时，可自动推荐，也可显式选择十套中性模板之一：

```powershell
python -m universal_report list-templates --repo-root .
python -m universal_report select-template --repo-root . --report-type experiment --course "计算机网络"

powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -BuiltInTemplateId neutral-engineering-lab `
  -ReportPath ".\examples\sample-report.txt" `
  -OutputDir ".\tests-output\engineering-demo"
```

可选 ID 为 `neutral-classic-lab`、`neutral-bordered-lab`、`neutral-engineering-lab`、`neutral-course-design`、`neutral-modern-minimal`、`neutral-compact-header-lab`、`neutral-review-panel-lab`、`neutral-code-notebook-lab`、`neutral-data-analysis-lab` 和 `neutral-project-dossier`。只要提供 `-TemplatePath`，用户模板就拥有最高优先级。

只有明确需要统一样式时使用：

```powershell
-TemplateStyleMode normalize -StyleFinalDocx
```

## 通用核心

- `universal-report-core.ps1`
  PowerShell 与 `universal_report` Python 包之间的稳定调用层。
- `plan-report-content.ps1`
  生成材料分析和结构化内容计划。
- `analyze-docx-template.ps1`
  生成 `TemplateStyleContract`，支持模板哈希与分析器版本缓存。
- `validate-docx-format.ps1`
  将生成文档与模板契约逐项比较。
- `run-visual-validation.ps1`
  严格模式 PDF、逐页预览和视觉检查。
- `convert-report-template.ps1`
  旧版 DOC 转 DOCX，LibreOffice 优先，Office COM 仅显式兜底。
- `new-report-test-fixtures.ps1`
  生成模板兼容性测试夹具。
- `build-neutral-templates.py`
  从零重建十套不含学校标识的内置 DOCX，并更新两个旧文件名兼容别名。

## 模板与字段

- `extract-docx-template.ps1`
- `generate-docx-field-map.ps1`
- `apply-docx-field-map.ps1`
- `convert-docx-template-frame.ps1`

保真模式不会默认调用统一样式或外框转换。模板结构无法安全填充时，应保留诊断并返回 `needs-fix`。

## 图片

- `generate-docx-image-map.ps1`
- `insert-docx-images.ps1`
- `render-vertical-lab-flowchart.py`

默认一图一行并保持比例。图片清单负责内容哈希去重、章节匹配、精确数量和选择原因；行布局或 2x2 网格仅在显式请求时使用。

## 转换与质量检查

- `export-docx-pdf.ps1`
- `validate-report-draft.ps1`
- `check-docx-layout.ps1`
- `check-report-profile-template-fit.ps1`
- `check-fast-report-session.ps1`
- `run-smoke-tests.ps1`
- `check-project-readiness.ps1`
- `self-check.ps1`

严格模式要求 PDF 导出和逐页视觉检查成功。LibreOffice 是首选；WPS / Word COM 默认关闭并受超时保护。

## 测试

```powershell
python -m unittest discover -s tests -v
powershell -ExecutionPolicy Bypass -File .\tests\run-core-smoke.ps1
powershell -ExecutionPolicy Bypass -File .\tests\run-neutral-template-catalog.ps1
powershell -ExecutionPolicy Bypass -File .\tests\run-universal-e2e.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\run-smoke-tests.ps1
```

修改内置模板后还必须运行：

```powershell
python -m universal_report audit-template-catalog --repo-root .
```

## 环境与安装

- `install-skill.ps1`
- `reset-openclaw-session.ps1`
- `report-defaults.ps1`
- `report-profiles.ps1`

路径和默认模板通过环境变量或配置文件提供，不绑定开发机盘符。详见 `docs/compatibility.md`。
