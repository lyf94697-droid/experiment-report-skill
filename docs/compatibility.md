# 兼容性说明

## 支持范围

| 项目 | 支持情况 |
| --- | --- |
| Windows PowerShell 5.1 | 支持，CI 与烟测覆盖 |
| PowerShell 7 | 支持，CI 矩阵覆盖 |
| Python 3 | 必需；核心分析主要使用标准库 |
| DOCX | 原生支持 |
| 旧版 DOC | 需 LibreOffice，或显式允许 WPS/Word COM |
| LibreOffice | 严格模式和稳定 PDF 导出的推荐依赖 |
| WPS / Microsoft Word | 仅显式 COM 兜底 |
| Gradio Web UI | 可选，见 `requirements-web.txt` |
| PyMuPDF / Pillow | PDF 逐页预览和视觉检查需要 |

## 跨机器配置

默认值不绑定本机盘符。可使用：

```powershell
$env:EXPERIMENT_REPORT_TEMPLATE_PATH = "D:\templates\experiment.docx"
$env:EXPERIMENT_REPORT_COURSE_DESIGN_TEMPLATE_PATH = "D:\templates\course-design.docx"
$env:EXPERIMENT_REPORT_BUILTIN_TEMPLATE_ID = "neutral-classic-lab"
$env:EXPERIMENT_REPORT_COURSE_DESIGN_TEMPLATE_ID = "neutral-course-design"
$env:EXPERIMENT_REPORT_OUTPUT_ROOT = "D:\report-output"
$env:EXPERIMENT_REPORT_CACHE_ROOT = "D:\report-cache"
$env:EXPERIMENT_REPORT_PYTHON = "C:\Python311\python.exe"
```

也可以设置：

```powershell
$env:EXPERIMENT_REPORT_CONFIG = "D:\config\experiment-report.json"
```

配置文件示例：

```json
{
  "defaultTemplate": "D:\\templates\\experiment.docx",
  "defaultBuiltInTemplateId": "neutral-classic-lab",
  "outputRoot": "D:\\report-output",
  "cacheRoot": "D:\\report-cache"
}
```

环境变量优先于配置文件，配置文件优先于仓库默认值。

没有配置外部默认模板时，仓库使用五套中性模板目录。普通实验默认 `neutral-classic-lab`，课程设计默认 `neutral-course-design`；也可通过 `-BuiltInTemplateId` 或 Web UI 下拉框选择其他模板。

## 模板兼容性

已覆盖的自动化夹具包括：

- 五套仓库中性模板及身份/来源审计
- 仓库默认模板
- 4 列学生信息表
- 5 列学生信息表
- 无外框表格
- 已有页面边框
- 封面与正文分节
- 图片占位符
- 空白 DOCX
- 旧版 DOC 转换失败
- 中文长课程名单元格拥挤

真实模板回归覆盖：

- 表格式操作系统实验报告模板
- 带封面、页眉页脚的课程设计模板
- 多表格、分节和复杂信息栏的综合实验报告模板

## 保证边界

保留模式保证的是“关键格式契约一致”，包括页面、核心角色样式和元数据表格几何。以下情况仍可能需要人工复核：

- 文本框、SmartArt、复杂浮动对象
- 宏、域代码或受保护表单
- 极端嵌套表格
- WPS 与 Word 的分页差异
- 模板本身已经存在窄列、竖排风险或错误样式

严格模式用于发现这些风险，不会在无法完成视觉验收时假装成功。
