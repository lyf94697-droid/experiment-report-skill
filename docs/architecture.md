# 架构说明

## 目标

本项目是一条“材料到可交付报告”的通用实验报告流水线。核心约束是：

1. 用户提供 DOCX 模板时，模板版式是最高优先级。
2. 用户没有模板时，从五套原创中性模板中选择，不直接复制学校模板。
3. 正文、图片和元数据按模板结构填入，不默认统一套版。
4. 每个阶段都有结构化产物和可定位错误。
5. 快速模式适合日常生成；严格模式必须完成 PDF 逐页视觉验收，否则状态为 `needs-fix`。

## 分层

- `universal_report/`
  Python 标准库优先的核心层，负责模板目录与选型、身份扫描、模板契约、内容计划、图片筛选、格式校验、视觉校验和跨机器配置。

- `scripts/build-report.ps1`
  主编排入口，串联已有字段映射、DOCX 填充、图片插入、布局检查和新的核心层。

- `scripts/build-report-from-url.ps1`
  教程链接、参考文本或对话材料入口。生成正文后进入同一主流水线。

- `scripts/build-report-from-feishu.ps1`
  直聊、附件路径和本地归档入口。不会为上传模板默认开启统一样式。

- `web_ui.py`
  Gradio 界面，负责材料上传、历史信息、进度显示、质量模式建议和产物下载。

## 阶段与产物

主流程按以下阶段记录到 `pipeline-trace.json`：

1. `materials-parsing`
2. `content-planning`
3. `body-generation`
4. `template-analysis`
5. `image-selection-matching`
6. `course-design-enhancements`（仅默认模板或显式规范化）
7. `docx-rendering`
8. `layout-validation`
9. `format-validation`
10. `visual-validation`

关键产物：

- `materials-analysis.json`
- `content-plan.json`
- `template-style-contract.json`
- `generated-field-map.json`
- `generated-image-map.json`
- `image-manifest.json`
- `layout-check.json`
- `format-validation.json`
- `visual-validation.json`
- `pipeline-trace.json`
- `summary.json`
- 最终 DOCX

## 模板样式契约

`TemplateStyleContract` 从 DOCX Open XML 包读取：

- 页面大小、方向、页边距、分节、页面边框
- 默认字体和 `Normal` 样式
- 标题、正文、图注、代码、表格正文等角色的有效样式
- `basedOn` 继承、段落直接格式、Run 直接格式和 East Asia 字体
- 表格列宽、边框、合并单元格、垂直对齐
- 页眉、页脚、标题区域、正文起点、图片占位符
- 空白模板、窄列、无外框等风险

缓存键由“分析器版本 + 模板 SHA-256”组成。模板内容或分析器规则改变时会自动失效。

## 保留与规范化模式

- `-TemplateStyleMode preserve`
  默认。用户模板只填内容，保留页面、字体、段落、表格、边框、页眉页脚和分节。

- `-TemplateStyleMode normalize -StyleFinalDocx`
  显式规范化。适合仓库默认模板、演示模板或用户明确要求统一风格时使用。

课程设计自动流程图和自动结构表格也只在默认模板或显式规范化模式启用，避免破坏用户模板。

## 五套中性模板

模板目录位于 `examples/report-templates/catalog.json`。选型顺序为：

1. 用户上传模板：直接使用并保真。
2. 课程设计：`neutral-course-design`。
3. 明确要求闭合外框：`neutral-bordered-lab`。
4. 明确要求现代简洁：`neutral-modern-minimal`。
5. 计算机或工程课程：`neutral-engineering-lab`。
6. 其他普通实验：`neutral-classic-lab`。

所有内置模板由 `scripts/build-neutral-templates.py` 从零生成。`audit-template-catalog` 会检查数量、文件、来源字段、真实学校/学院名称、示例长数字身份和嵌入媒体；审计不通过的模板不能作为回退模板。

## 图片策略

`ImagePaths` 会先经过：

- 文件存在性检查
- SHA-256 内容去重
- 文件名语义分类
- 章节匹配
- 图注生成
- 相关性排序

默认布局是一图一行、图注在下、保持比例并尽量与图注同页。只有显式请求网格布局时才使用多图并排。

## 严格模式

严格模式链路：

```text
DOCX -> LibreOffice PDF -> 每页 PNG -> 非空白、闭合外框等视觉检查
```

LibreOffice 是默认和推荐的无头转换器。WPS/Word COM 只有显式传入 `-AllowOfficeCom` 或设置 `EXPERIMENT_REPORT_ALLOW_OFFICE_COM=1` 时才会使用，并带超时与新启动进程清理。

如果 PDF 转换或视觉检查失败，DOCX 和 JSON 诊断仍会保留，但最终状态不能是成功。
