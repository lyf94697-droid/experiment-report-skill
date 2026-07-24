# Experiment Report Skill

把实验要求、正文、Word 模板、截图和代码整理成一份可继续编辑的中文实验报告。普通实验和课程设计都可以走这套流程。

这套工具主要处理两件麻烦事：

- 老师发了模板，填完内容后字体、表格、边框和分页不能乱。
- 截图很多，插入位置、大小、图注和编号需要统一。

[![Quality Checks](https://github.com/lyf94697-droid/experiment-report-skill/actions/workflows/quality.yml/badge.svg)](https://github.com/lyf94697-droid/experiment-report-skill/actions/workflows/quality.yml)
[![Smoke Tests](https://github.com/lyf94697-droid/experiment-report-skill/actions/workflows/smoke-tests.yml/badge.svg)](https://github.com/lyf94697-droid/experiment-report-skill/actions/workflows/smoke-tests.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## 先跑一份看看

环境建议：

- Windows 10 / 11
- Windows PowerShell 5.1 或 PowerShell 7
- Python 3.11
- 严格检查需要 LibreOffice；普通生成不需要

下载仓库：

```powershell
git clone https://github.com/lyf94697-droid/experiment-report-skill.git
cd experiment-report-skill
```

不使用 Git，也可以在仓库页面点 **Code → Download ZIP**。

运行仓库自带的示例：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1
```

生成结果放在：

```text
tests-output/one-click-demo-时间戳/
```

先打开其中的最终 DOCX，再看 `summary.json` 和 `layout-check.json`。

## 模板怎么处理

要生成 DOCX，而对话里还没有提到模板时，Skill 会先问：

> 你有老师、学校或自己认可的优秀 DOCX/DOC 模板吗？有的话请上传，我会优先保持原格式；没有的话，我会从十套不含学校标识的内置模板中选择。

选模板时按下面的顺序：

1. 老师或学校发的模板优先。
2. 用户上传了模板，就沿用这份模板，不和内置模板混搭。
3. 只填写内容，尽量保留纸张、页边距、字体、字号、行距、段距、表格、边框、页眉页脚、分节和图注样式。
4. 模板只有局部问题时，只修对应位置，不整份重排。
5. 没有模板时，再按课程和报告类型选择一套内置模板。
6. 模板无法解析时给出原因，不悄悄换成另一种版式。

DOCX 在 Word、WPS 和 LibreOffice 中可能出现少量分页差异，所以“保留模板”指关键格式和结构保持一致，不承诺不同软件逐像素相同。

模板分析结果保存在 `template-style-contract.json`，脚本和文档中称为 `TemplateStyleContract`。它记录页面、字体、段落、表格、页眉页脚和图注等后续需要对照的格式。

## 十套内置模板

内置模板不含学校名、学院名、校徽、水印、真实学生信息和第三方图片。

| 模板 ID | 样式 | 适合的内容 |
| --- | --- | --- |
| `neutral-classic-lab` | 经典实验报告 | 大多数普通课程实验 |
| `neutral-bordered-lab` | 闭合外框 | 需要传统纸质表格和页面外框 |
| `neutral-engineering-lab` | 工程技术 | 计算机网络、操作系统、数据库、Web、Android、软件工程 |
| `neutral-course-design` | 课程设计 | 综合实验、课程设计和中等篇幅项目报告 |
| `neutral-modern-minimal` | 现代简洁 | 不要求传统学校表格的报告 |
| `neutral-compact-header-lab` | 紧凑信息条 | 周实验、短实验和快速记录 |
| `neutral-review-panel-lab` | 评阅记录 | 需要成绩、教师评语和签名栏 |
| `neutral-code-notebook-lab` | 程序设计 | C/C++、Python、Java、算法、测试和调试 |
| `neutral-data-analysis-lab` | 数据分析 | 原始数据、统计、趋势和误差分析 |
| `neutral-project-dossier` | 项目技术 | 长篇系统设计、课程项目和参考文献 |

普通实验默认使用 `neutral-classic-lab`，课程设计默认使用 `neutral-course-design`。也可以在命令行或 Web UI 里手工选择。

这些模板由仓库脚本生成。参考过的公开版式、采用了哪些通用排版原则，以及哪些内容明确不复制，记录在 [模板调研说明](docs/template-research.md) 和 [模板目录](examples/report-templates/catalog.json) 中。

## 准备材料

材料不必一次凑齐，常见输入如下：

| 材料 | 用途 | 是否必需 |
| --- | --- | --- |
| 实验要求或任务书 | 确定章节、评分点和检查项 | 建议提供 |
| 报告正文 | 已经写好的 `.txt` 正文 | 本地脚本必需 |
| DOCX / DOC 模板 | 决定最终版式 | 可选 |
| 姓名、学号、班级、课程名、实验名 | 填写封面和信息表 | 可后补 |
| 截图 | 作为步骤、结果和排障证据 | 可选 |
| 代码、配置、命令记录 | 补充实现过程和关键内容 | 可选 |
| 教程链接 | 参考操作顺序和背景 | 可选 |

教程只作参考。报告中的结果、数据和结论应以自己的截图、代码和运行记录为准。

<!-- project-readiness:usage-tutorial -->
## 三种用法

### 1. 在 Codex 中使用

安装到本机 Skill 目录：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1
```

默认安装位置：

```text
%USERPROFILE%\.codex\skills\experiment-report
```

安装后新开一个任务，直接说明课程、实验名称、模板和材料路径。例如：

```text
请用 experiment-report skill 整理实验报告。

课程：计算机网络
实验：局域网配置与连通性测试
模板：.\materials\实验报告模板.docx
正文：.\materials\report.txt
截图目录：.\materials\images
输出目录：.\outputs\network-report

沿用模板原有字体、表格和边框。截图一张一行，图注放在图片下方。
```

给 OpenClaw 或 agents 目录安装：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1 -Platform openclaw -Force
```

材料来自聊天附件或本地附件路径时，可以使用：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1
```

完整参数见 [聊天附件示例](examples/feishu-uploaded-images-docx-prompt.md)。

### 2. 直接运行 PowerShell

使用内置工程技术模板：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -BuiltInTemplateId "neutral-engineering-lab" `
  -ReportPath ".\examples\cases\network-dos\report.txt" `
  -MetadataPath ".\examples\cases\network-dos\metadata.json" `
  -RequirementsPath ".\examples\cases\network-dos\requirements.json" `
  -ImageSpecsPath ".\examples\cases\network-dos\image-specs.json" `
  -OutputDir ".\outputs\network-dos"
```

使用老师发的模板时，把 `-BuiltInTemplateId` 换成：

```powershell
-TemplatePath ".\materials\teacher-template.docx" `
-TemplateStyleMode preserve
```

课程设计再加：

```powershell
-ReportProfileName course-design-report
```

### 3. 使用本地网页

安装网页依赖：

```powershell
python -m pip install -r requirements-web.txt
```

启动：

```powershell
python web_ui.py
```

浏览器打开：

```text
http://127.0.0.1:7860
```

网页里可以填写学生信息、实验名称和要求，上传模板、正文、截图、代码及附件，也可以在没有模板时选择内置模板。每次运行的材料和过程记录保存在 `outputs/web-ui/`，最终文件路径会显示在结果区。

如需改到其他目录：

```powershell
$env:EXPERIMENT_REPORT_OUTPUT_ROOT = "D:\report-output"
$env:EXPERIMENT_REPORT_CACHE_ROOT = "D:\report-cache"
python web_ui.py
```

## 图片排版

默认规则：

- 一张图占一行，不默认拼成两列。
- 保持原图比例，宽度不超过正文区域。
- 图注紧跟在图片下方，居中，连续编号。
- 按实验步骤、结果或问题分析插到对应章节。
- 相同内容的图片按文件哈希去重。
- 指定图片数量后，数量不符就报错，不复制图片凑数。
- 只有明确要求多图并排时，才使用双列或网格。

图片的路径、章节、图注、顺序、布局、选择原因和被排除文件会写入 `image-manifest.json`。

## 快速检查与严格检查

日常生成默认使用快速设置：

```powershell
-PipelineMode fast -QualityMode fast
```

模板格式对照每次都会运行。有截图时会检查图片和图注；没有截图的快速任务会跳过布局检查。

新模板、旧版 DOC、复杂表格或正式交付前，建议使用：

```powershell
-PipelineMode full -QualityMode strict
```

严格检查会导出 PDF、渲染每一页预览图，再检查分页、空白页、文字拥挤、表格溢出、边框、图片和图注。推荐安装 LibreOffice。无法完成 PDF 或页面检查时，DOCX 和诊断文件仍会保留，但结果会标为 `needs-fix`。

<!-- project-readiness:input-output -->
## 输出目录

根据材料和检查方式，目录中会出现以下文件：

```text
outputs/network-dos/
├─ <最终报告>.docx
├─ summary.json
├─ pipeline-trace.json
├─ materials-analysis.json
├─ content-plan.json
├─ template-style-contract.json
├─ generated-field-map.json
├─ generated-image-map.json
├─ image-manifest.json
├─ image-placement-plan.md
├─ layout-check.json
├─ format-validation.json
├─ visual-validation.json
├─ strict-preview.pdf
└─ strict-preview/
```

没有截图时不会生成图片映射；没有开启严格检查时不会生成 PDF、页面预览和 `visual-validation.json`。

常用文件：

- `summary.json`：最终 DOCX 路径、使用的模板、检查结果和当前状态。
- `pipeline-trace.json`：每个处理阶段的状态和错误原因。
- `template-style-contract.json`：从模板读取到的页面、字体、段落、表格和图注规则。
- `format-validation.json`：模板与最终 DOCX 的关键格式对照。
- `layout-check.json`：图片、图注、分页和文档结构检查。

<!-- project-readiness:scenarios -->
## 示例

仓库里放了三组完整材料，不写死页数，实际页数以本机 Word、WPS 或 LibreOffice 渲染结果为准。

- [局域网与 DOS 命令](examples/cases/network-dos/README.md)：拓扑、主机配置、连通性和故障分析，共 5 张示例图。
- [进程调度](examples/cases/os-process-scheduling/README.md)：流程、甘特图、指标、控制台和测试结果。
- [学生成绩管理系统课程设计](examples/cases/course-design-student-management/README.md)：需求、架构、数据模型、界面和测试材料。

演示图片见 [demo/README.md](demo/README.md)，模板文件见 [examples/report-templates](examples/report-templates/README.md)。

<!-- project-readiness:limitations -->
## 使用边界

- 复杂学校模板仍可能需要调整字段映射，尤其是嵌套表格、文本框、公式、宏和特殊域。
- 旧版 `.doc` 需要先转换成 `.docx`。优先使用 LibreOffice；WPS 或 Word COM 只在明确开启时使用。
- Word、WPS 和 LibreOffice 的分页算法不同，最终提交前仍应在实际使用的软件里打开检查。
- 模糊截图无法补出看不清的数据，也不会根据教程编造运行结果。
- 从公开教程整理正文时，不直接复制原文；实验结果仍以用户材料为准。
- 严格检查能发现常见排版问题，但不能代替课程要求和人工复核。

## 文档导航

- [使用流程](docs/usage-flow.md)
- [模板填充规则](docs/template-filling.md)
- [模板调研与来源边界](docs/template-research.md)
- [截图与图注](docs/screenshot-evidence.md)
- [一键示例说明](docs/one-click-demo.md)
- [架构说明](docs/architecture.md)
- [兼容性](docs/compatibility.md)
- [常见问题](docs/troubleshooting.md)
- [后续安排](ROADMAP.md)
- [GitHub 与社交平台发布材料](docs/social-launch-kit.md)
- [全部文档](docs/README.md)

## 检查仓库

首页和示例材料：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-project-readiness.ps1
```

模板目录：

```powershell
python -m universal_report audit-template-catalog --repo-root .
```

完整烟测：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-smoke-tests.ps1
```

十套模板与上传模板保真：

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\run-neutral-template-catalog.ps1
powershell -ExecutionPolicy Bypass -File .\tests\run-template-fidelity-corpus.ps1
```

通用端到端场景：

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\run-universal-e2e.ps1
```

## 仓库目录

```text
experiment-report-skill/
├─ docs/                  使用说明、模板规则和排障
├─ examples/              示例正文、模板、截图和 JSON
├─ profiles/              普通实验与课程设计规则
├─ references/            Skill 运行时参考
├─ scripts/               生成、填充、插图和检查脚本
├─ tests/                 单元测试与端到端测试
├─ universal_report/      模板、图片、配置和验证模块
├─ web_ui.py              本地网页入口
└─ SKILL.md               Skill 规则
```

## License

[MIT](LICENSE)
