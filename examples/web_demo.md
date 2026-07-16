# Web UI 演示

本示例说明如何从浏览器上传模板、截图、代码和文字材料，生成 DOCX 及结构化诊断产物。DOCX 是主交付物；PDF 和逐页预览用于严格模式验收。

## 启动

安装可选依赖：

```powershell
python -m pip install -r requirements-web.txt
```

启动界面：

```powershell
python web_ui.py
```

浏览器打开：

```text
http://127.0.0.1:7860
```

## 输入

界面支持：

- 实验报告或课程设计报告
- 快速本地草稿或智能长文模式
- 快速生成或严格检查
- 课程、题目、学生信息和实验要求
- 对话式需求、教程链接和补充说明
- 可选 DOCX / DOC 模板
- 无上传模板时的五套中性模板选择
- 多张截图和多个代码文件
- 可信本机环境下的本地目录或文件路径

手工填写字段的优先级高于对话式文本自动提取。

页面会先提示是否有老师、学校或自己认可的优秀模板。上传模板始终优先；没有模板时可以自动推荐，也可以手工选择经典、闭合外框、工程技术、课程设计或现代简洁模板。

## 模板保真

上传用户模板时默认使用 `TemplateStyleMode=preserve`。系统会：

1. 分析页面、字体、段落、标题、表格、边框、页眉页脚和分节。
2. 生成 `template-style-contract.json`。
3. 按模板结构填入正文、字段和图片。
4. 生成 `format-validation.json`，逐项检查关键格式是否漂移。

界面不会为上传模板默认开启统一样式或自动重建外框。只有显式选择规范化时，才使用仓库样式 profile。

## 进度与结果

点击“生成报告”后，界面会显示：

- 当前阶段与整体进度
- 最终状态、警告和可执行的修复建议
- 模板契约摘要
- 图片选择、去重和数量结果
- 布局、格式和视觉验证状态
- DOCX、PDF、预览图和诊断文件路径

常见阶段包括材料解析、内容规划、模板分析、图片匹配、DOCX 渲染、布局验证、格式验证和视觉验证。

## 图片

默认是一图一行、图注在下。图片先按内容哈希去重，再按文件名和章节语义匹配。若指定精确图片数量，可在脚本入口使用 `-RequestedImageCount`；可用图片不足时会明确失败，不会重复图片凑数。

## 快速模式与严格模式

- 快速模式：完成结构、布局和格式检查，适合日常生成。
- 严格模式：额外执行 `DOCX → PDF → 每页 PNG → 视觉检查`。任何一环无法完成，状态为 `needs-fix`。

严格模式推荐安装 LibreOffice，并确保 `soffice` 可用。逐页预览需要 PyMuPDF / Pillow。

WPS / Microsoft Word COM 默认关闭。如果用户明确允许本机 Office 自动化，可在启动前设置：

```powershell
$env:EXPERIMENT_REPORT_ALLOW_OFFICE_COM = "1"
python web_ui.py
```

COM 兜底有超时保护，只清理本次自动化新启动的 Office 进程。

## 跨机器配置

默认输出位于仓库的 `outputs/web-ui/`。可以通过环境变量设置模板、输出和缓存目录：

```powershell
$env:EXPERIMENT_REPORT_TEMPLATE_PATH = "D:\templates\experiment.docx"
$env:EXPERIMENT_REPORT_COURSE_DESIGN_TEMPLATE_PATH = "D:\templates\course-design.docx"
$env:EXPERIMENT_REPORT_BUILTIN_TEMPLATE_ID = "neutral-engineering-lab"
$env:EXPERIMENT_REPORT_OUTPUT_ROOT = "D:\report-output"
$env:EXPERIMENT_REPORT_CACHE_ROOT = "D:\report-cache"
```

也可指定 JSON 配置：

```powershell
$env:EXPERIMENT_REPORT_CONFIG = "D:\config\experiment-report.json"
```

环境变量优先于配置文件，配置文件优先于仓库默认值。

## 本地路径安全

浏览器上传是默认方式。只有可信本机环境需要直接读取填写的路径时，才启用：

```powershell
$env:OPENCLAW_WEB_UI_ALLOW_LOCAL_PATHS = "1"
python web_ui.py
```

生成失败时，界面会保留 DOCX、日志和 JSON 诊断，并给出失败阶段与下一步建议。
