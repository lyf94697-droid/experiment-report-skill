# 排障指南

## 严格模式提示找不到 LibreOffice

现象：

- `visual-validation.json` 中出现 `pdf-export-failed`
- `generationStatus` 为 `needs-fix`

处理：

1. 安装 LibreOffice。
2. 确认 `soffice` 在 `PATH` 中，或位于标准安装目录。
3. 重新运行严格模式。

不要为了绕过检查而把严格模式结果当成成功。只有明确允许时才使用：

```powershell
$env:EXPERIMENT_REPORT_ALLOW_OFFICE_COM = "1"
```

WPS/Word 兜底有超时保护，并只清理本次自动化新启动的 Office 进程。

## 旧版 `.doc` 模板无法转换

转换顺序是 LibreOffice 优先、WPS/Word COM 显式兜底。失败时原模板不会被修改。

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\convert-report-template.ps1 `
  -SourcePath ".\template.doc" `
  -OutputDir ".\converted"
```

如明确允许 Office 自动化，再加 `-AllowOfficeCom`。

## `format-validation.json` 未通过

先查看：

- `failedCodes`
- 每项的 `templateValue`
- `documentValue`
- `tolerance`
- `location`

常见代码：

- `page-size` / `page-margins`
  页面尺寸或页边距发生变化。
- `report-title-*`
  报告标题字体、字号或对齐变化。
- `body-*`
  正文段落误用了标题或其他样式。
- `metadata-table-grid`
  信息表列宽或列数变化。
- `metadata-vertical-text-risk`
  最终信息表出现了模板原本没有的竖排风险。

快速模式会保留 DOCX 并标记 `needs-fix`；严格模式或 `-FailOnFormatValidation` 会直接失败。

## 图片数量不对

查看 `image-manifest.json`：

- `requestedCount`
- `selectedCount`
- `duplicatesFiltered`
- `rejected`
- 每张图的 `selectionReason`

如需精确数量：

```powershell
-ImagePaths ".\a.png",".\b.png",".\c.png" -RequestedImageCount 3
```

可用图片不足时流程会失败，不会用重复图片补数。

## 图片和图注被分页拆开

图片段落带 `keepNext`，图注带 `keepLines`。如果仍发生拆分：

1. 减小图片宽度。
2. 检查模板段前段后距。
3. 严格模式导出逐页预览。
4. 避免在窄表格单元格中放大图。

## 信息表出现竖排或字符拥挤

查看 `template-style-contract.json` 中：

```text
structure.metadataTable.verticalTextRisk
```

处理优先级：

1. 保留模板列结构并缩短字段值。
2. 使用模板预留的跨列单元格。
3. 仅在用户允许修改模板时调整列宽。

不要在保留模式下擅自统一重建信息表。

## Web UI 不接受本机路径

浏览器上传默认只允许临时上传文件。可信本机环境如需直接填写路径：

```powershell
$env:OPENCLAW_WEB_UI_ALLOW_LOCAL_PATHS = "1"
python web_ui.py
```

## 中文路径或 JSON 乱码

- PowerShell 中使用 `-Encoding UTF8`。
- 新增 `.ps1` 文件使用 UTF-8 BOM，以兼容 Windows PowerShell 5.1。
- `SKILL.md` 不要加 BOM，避免 frontmatter 解析失败。

## 缓存看起来没有更新

模板契约缓存会按分析器版本和模板哈希自动失效。也可以设置独立缓存目录：

```powershell
$env:EXPERIMENT_REPORT_CACHE_ROOT = "D:\cache\experiment-report"
```

排障时可临时指定新的 `-TemplateCacheDir`。
