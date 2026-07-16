# Experiment Report Skill

> 面向中文大学实验报告 / 课程设计报告的通用 Skill + PowerShell 本地流水线。
> 支持 Codex、OpenClaw 和本地 Web UI，把“实验题目、教程链接、正文、学校模板、截图证据、版式检查”收成一条可复查、可复用、可交付的 `docx` 生成链路。

[![Quality Checks](https://github.com/lyf94697-droid/experiment-report-skill/actions/workflows/quality.yml/badge.svg)](https://github.com/lyf94697-droid/experiment-report-skill/actions/workflows/quality.yml)
[![Smoke Tests](https://github.com/lyf94697-droid/experiment-report-skill/actions/workflows/smoke-tests.yml/badge.svg)](https://github.com/lyf94697-droid/experiment-report-skill/actions/workflows/smoke-tests.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

## 5 分钟上手

### 如何简单下载

有 Git 的用户可以直接克隆：

```powershell
git clone https://github.com/lyf94697-droid/experiment-report-skill.git
cd experiment-report-skill
```

不想装 Git 的用户，可以在 GitHub 仓库页点击 **Code** -> **Download ZIP**，解压后进入 `experiment-report-skill` 文件夹。

### 如何简单使用

最快体验方式是跑仓库自带的一键 demo：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1
```

它会使用内置模板、正文和截图，自动生成可打开检查的 `docx` 文件。输出目录默认在 `tests-output/one-click-demo-时间戳/`。

如果想用网页界面：

```powershell
python -m pip install -r requirements-web.txt
python web_ui.py
```

然后打开：

```text
http://127.0.0.1:7860
```

如果想作为 Codex / OpenClaw 的本地 Skill 使用：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1
```

安装后新开一个 Codex 会话，直接描述你的实验报告需求、模板路径和截图路径即可。

## 通用流水线与模板保真

主流程现在按“材料解析 → 内容规划 → 正文生成 → 模板分析 → 图片筛选匹配 → DOCX 渲染 → 格式/布局/视觉验证”分阶段执行，每个阶段都写入 `pipeline-trace.json`，失败时可定位到具体阶段和错误代码。

用户提供学校 DOCX 模板时，默认使用 `-TemplateStyleMode preserve`：

- 先生成 `TemplateStyleContract`（产物为 `template-style-contract.json`），记录页面、字体、段落、标题层级、表格几何、边框、页眉页脚、分节和图注等有效样式；
- 正文、元数据和图片按模板结构填入，不默认套用仓库的统一字体、边框或课程设计增强；
- 生成后用 `format-validation.json` 对照模板契约逐项验证；
- 只有显式传入 `-TemplateStyleMode normalize -StyleFinalDocx` 时，才应用仓库样式 profile。

如果用户要 DOCX 但尚未说明模板，Skill 会先问一次是否有老师、学校或自己认可的优秀模板。用户上传的模板始终优先；明确没有模板时，再从五套不含学校标识的内置模板中选择：

| ID | 风格 | 适用场景 |
| --- | --- | --- |
| `neutral-classic-lab` | 经典实验报告 | 普通课程实验默认选择 |
| `neutral-bordered-lab` | 闭合外框 | 强调传统纸质版式或页面外框 |
| `neutral-engineering-lab` | 工程技术 | 计算机网络、操作系统、数据库、Java、Web、Android 等 |
| `neutral-course-design` | 课程设计 | 课程设计、综合实验和项目报告 |
| `neutral-modern-minimal` | 现代简洁 | 不要求传统表格式的轻量报告 |

五份 DOCX 都由仓库脚本从零生成，不包含真实学校名、校徽、示例学生身份或第三方图片。公开模板只作为设计原则参考，来源和使用边界记录在 `examples/report-templates/catalog.json`。

图片默认一图一行、图注在下。系统会按内容哈希去重，记录选择和拒绝原因；传入 `-RequestedImageCount` 时会严格满足数量，不会复制图片凑数。

快速模式执行结构、布局和格式检查。严格模式额外执行 `DOCX → PDF → 每页预览 → 视觉检查`；缺少 LibreOffice 或视觉检查失败时，最终状态是 `needs-fix`，不会把未验收结果标成成功。

### 直接复制的提示词

#### 1. 从教程链接和截图生成实验报告

```text
请使用 experiment-report skill 帮我生成并填充一份中文实验报告。

工作目录：当前仓库目录

基础信息：
- 课程名称：计算机网络
- 实验名称：局域网搭建与常用 DOS 命令使用
- 学校模板：.\materials\学校实验报告模板.docx
- 姓名：示例学生
- 学号：20260001
- 班级：计科2401
- 指导教师：示例教师

材料：
- 教程链接：https://blog.csdn.net/你的文章链接
- 截图路径：
  - .\materials\images\step-1.png
  - .\materials\images\result-1.png

要求：
- 不要照抄教程，把教程改写成实验报告正文。
- 正文必须能对应我提供的真实截图和实验结果。
- 最终生成 docx，并插入截图、图注和必要的版式检查结果。
```

#### 2. 已有正文，直接填学校模板

```text
我已经有实验报告正文，请帮我填进学校 docx 模板并整理格式。

工作目录：当前仓库目录

请使用：
- 模板：.\materials\学校实验报告模板.docx
- 正文：.\materials\report.txt
- 输出目录：.\outputs\final-report

要求：
- 保留学校模板原有表格、外框和标题结构。
- 使用模板保真模式，不要默认统一字体或重建页面结构。
- 自动补全姓名、学号、课程名、实验名等字段。
- 如果我提供截图，也要插入到合适章节并生成具体图注。
- 完成后给出最终 docx 路径和 layout check 结果。
```

#### 3. 课程设计报告

```text
请按课程设计报告模式生成并填充报告。

工作目录：当前仓库目录

项目信息：
- 课程名称：软件工程课程设计
- 题目：学生成绩管理系统
- 模板路径：.\materials\课程设计模板.docx
- 输出目录：.\outputs\course-design

材料：
- 需求说明、功能模块、数据库设计、运行截图和测试结果我会提供在当前对话或本地文件中。

要求：
- 使用 course-design-report profile。
- 包含需求分析、总体设计、详细设计、运行结果、测试分析和总结。
- 流程图或总体设计图要单独放大展示，不要和普通截图挤在一行。
```

更多可直接套用的提示词见：

- [examples/one-shot-uploaded-images-docx-prompt.md](examples/one-shot-uploaded-images-docx-prompt.md)
- [examples/local-uploaded-images-docx-prompt.md](examples/local-uploaded-images-docx-prompt.md)
- [examples/feishu-uploaded-images-docx-prompt.md](examples/feishu-uploaded-images-docx-prompt.md)

## 项目定位

这个仓库不是“万能文档生成器”，而是一个把中文大学实验报告场景做深、做稳的通用 Skill 项目：

- 先生成或接收结构化中文报告正文
- 再把正文和基础信息填进 `docx` / Word / WPS 模板
- 再插入截图、生成图注、处理多图布局
- 最后做样式收尾和 layout check，得到可检查的成品

适合解决这些问题：

- 实验要求、教程、截图、代码和结果散落在不同地方
- 学校模板是空白 `docx`，字段和正文都需要手工填
- 截图不知道该插到哪一节、图注怎么写、几张图怎么排
- 交作业前还要手动修标题、空行、分页、外框和图注编号

## 项目亮点

- 默认主线聚焦 `experiment-report`，优先把常见中文实验报告做稳
- 显式支持 `course-design-report`，适合课程设计报告和学校固定模板
- 支持“已有正文直接填模板”和“从教程链接生成正文再填模板”两条路径
- 用户模板默认保真，支持有效样式继承分析、契约缓存和生成后格式对照
- 支持截图语义匹配、内容哈希去重、精确数量要求和图片选择清单
- 默认一图一行；显式需要时仍支持每行 2 图 / 2x2 图片块
- 支持快速/严格质量模式、逐阶段错误、PDF 逐页预览和视觉检查
- 附带一键 demo、项目就绪检查、烟测、GitHub workflow、Issue / PR 模板和开源协作文档
- 仓库内置演示素材和 3 个典型案例，适合做 GitHub 展示、录屏或简历项目

## 演示预览

| 步骤截图 | 结果截图 |
| --- | --- |
| ![Network config preview](demo/assets/step-network-config.png) | ![Ping result preview](demo/assets/result-ping.png) |

完整演示素材、2x2 拼图预览和展示建议见 [demo/README.md](demo/README.md)。

<!-- project-readiness:usage-tutorial -->
## 中文使用教程

### 1. 安装 skill

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1
```

默认安装到 Codex 的本地 skill 目录：`$HOME\.codex\skills\experiment-report`。安装后新开一个 Codex 会话，提到“实验报告 / 课程设计报告 / 模板填充”时即可触发。

如果还要给 OpenClaw 使用，可以安装到兼容的 agents/OpenClaw 目录：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-skill.ps1 -Platform openclaw -Force
```

### 2. 跑一键演示

这个演示不依赖在线生成正文，直接用仓库自带正文、模板和截图走完“填模板 + 插图 + 排版 + layout check”。

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1
```

默认会在 `tests-output/one-click-demo-时间戳/` 下生成最终 `docx`、字段映射、图片映射、图片放置计划、layout check 和摘要文件。

### 3. 检查项目展示材料

如果你要把仓库发到 GitHub、写进简历或拿给别人试用，先跑：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-project-readiness.ps1
```

它会检查 README、关键 docs、典型案例和案例 JSON 是否齐全。

### 4. 已有正文 + 学校模板

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -TemplatePath ".\examples\report-templates\experiment-report-template.docx" `
  -ReportPath ".\examples\sample-report.txt" `
  -MetadataPath ".\examples\cases\network-dos\metadata.json" `
  -ImageSpecsPath ".\examples\cases\network-dos\image-specs.json" `
  -RequirementsPath ".\examples\cases\network-dos\requirements.json" `
  -OutputDir ".\tests-output\network-dos-case" `
  -TemplateStyleMode preserve
```

### 5. 教程链接 + 截图 + 学校模板

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-url.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
  -ExperimentName "局域网搭建与常用 DOS 命令使用" `
  -TemplatePath ".\materials\template.docx" `
  -StudentName "示例学生" `
  -StudentId "20260001" `
  -ClassName "计科 2201" `
  -ImagePaths ".\materials\step-1.png",".\materials\result-1.png" `
  -RequestedImageCount 2 `
  -TemplateStyleMode preserve `
  -OutputDir ".\outputs\from-url"
```

### 6. 飞书 / 直聊场景

如果材料来自飞书、聊天窗口或本地附件路径，可以使用本地 wrapper，把正文、截图和生成产物归档到一个输出目录：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report-from-feishu.ps1 `
  -ReferenceUrls "https://blog.csdn.net/..." `
  -CourseName "计算机网络" `
  -ExperimentName "局域网搭建与常用 DOS 命令使用" `
  -TemplatePath ".\materials\template.docx" `
  -StudentName "示例学生" `
  -StudentId "20260001" `
  -ClassName "计科 2201" `
  -ImagePaths ".\materials\step-1.png",".\materials\result-1.png" `
  -RequestedImageCount 2 `
  -TemplateStyleMode preserve `
  -OutputDir ".\outputs\from-chat"
```

课程设计报告请显式传：

```powershell
-ReportProfileName course-design-report
```

更完整的流程见 [docs/usage-flow.md](docs/usage-flow.md)。

<!-- project-readiness:input-output -->
## 输入与输出示例

### 典型输入

| 输入 | 说明 | 示例 |
| --- | --- | --- |
| `TemplatePath` | 用户上传的学校或教师模板，优先级最高 | `materials/teacher-template.docx` |
| `BuiltInTemplateId` | 无用户模板时选择五套中性模板之一 | `neutral-engineering-lab` |
| `ReportPath` | 已有报告正文 | `examples/sample-report.txt` |
| `MetadataPath` | 姓名、学号、课程名、实验名等短字段 | `examples/cases/network-dos/metadata.json` |
| `RequirementsPath` | 章节、关键词和禁用词检查 | `examples/cases/network-dos/requirements.json` |
| `ImageSpecsPath` | 截图路径、图注、章节和布局 | `examples/cases/network-dos/image-specs.json` |
| `ReferenceUrls` | 教程或 CSDN 参考链接 | `https://blog.csdn.net/...` |

### 典型输出

```text
tests-output/network-dos-case/
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
├─ pipeline-trace.json
├─ summary.json
├─ report.cleaned.txt
└─ <最终报告>.docx
```

其中 `summary.json` 汇总最终状态，`pipeline-trace.json` 记录阶段与错误，模板契约和三类验证文件用于定位格式、布局与视觉风险，最终 `docx` 用来人工打开复核和交付。

<!-- project-readiness:scenarios -->
## 适用场景

- 中文大学实验报告：计算机网络、操作系统、数据库、软件测试、程序设计等课程
- 课程设计报告：需要需求分析、方案设计、运行结果和设计总结的小型项目
- 学校固定模板：老师给了 `.docx` 模板，需要保留字段、表格、外框和标题结构
- 截图证据型作业：报告必须嵌入命令输出、程序界面、流程图或测试结果截图
- 项目展示：需要一个能演示“提示词 + 本地脚本 + 文档自动化 + 质量检查”的简历项目

## 典型案例

- [network-dos](examples/cases/network-dos/README.md)：7 页局域网实验成品，含拓扑、命令证据和排障流程
- [os-process-scheduling](examples/cases/os-process-scheduling/README.md)：8 页进程调度实验成品，含甘特图、性能指标和回归测试
- [course-design-student-management](examples/cases/course-design-student-management/README.md)：16 页课程设计成品，含 6 张图和与正文一致的数据库表结构

<!-- project-readiness:limitations -->
## 限制说明

- 不能保证任意复杂学校模板都零人工适配，复杂模板仍需要看字段映射 diagnostics
- 从 CSDN 或公开教程生成正文时，参考内容只提供背景和流程，结果必须以用户真实截图和数据为准
- 这个项目能降低照抄风险，但不能自动保证学术合规；提交前仍需使用者复核
- 图像插入依赖本地图片路径可访问，截图模糊或信息不足时不能凭空补细节
- `build-report-from-url.ps1` 属于可选智能长文通道，需要 OpenClaw CLI 和浏览器 profile 可用；稳定主线不依赖它
- 严格模式依赖 LibreOffice 完成稳定 PDF 导出，并依赖 PyMuPDF / Pillow 做逐页视觉检查
- Word / WPS GUI 自动操作不是主路径；COM 兜底默认关闭，只能显式启用并受超时保护

后续方向见 [ROADMAP.md](ROADMAP.md)。

## 仓库目录

```text
experiment-report-skill/
├─ demo/                  GitHub / 小红书 / 抖音友好的演示素材
├─ docs/                  使用流程、模板机制、CSDN 改写、截图证据等文档
├─ examples/              典型案例、样例正文、JSON、Prompt、模板
├─ profiles/              报告 profile 定义
├─ references/            skill 运行时参考规则
├─ scripts/               主流程脚本、辅助脚本和检查脚本
├─ tests/                 模板夹具、单元测试和端到端测试
├─ universal_report/      模板契约、内容、图片、验证和配置核心
├─ agents/                Codex / OpenClaw UI 元数据
├─ SKILL.md               skill 主说明
└─ README.md              项目首页
```

第一次阅读建议顺序：

1. [docs/usage-flow.md](docs/usage-flow.md)
2. [examples/cases/README.md](examples/cases/README.md)
3. [docs/template-filling.md](docs/template-filling.md)
4. [docs/architecture.md](docs/architecture.md)
5. [docs/compatibility.md](docs/compatibility.md)
6. [docs/troubleshooting.md](docs/troubleshooting.md)

## 文档导航

- [docs/README.md](docs/README.md)：文档总导航
- [docs/architecture.md](docs/architecture.md)：分阶段流水线、模板契约和验证架构
- [docs/compatibility.md](docs/compatibility.md)：平台、依赖、跨机器配置和模板边界
- [docs/troubleshooting.md](docs/troubleshooting.md)：严格模式、格式漂移、图片数量和 Office 转换排障
- [docs/usage-flow.md](docs/usage-flow.md)：完整使用流程
- [docs/template-filling.md](docs/template-filling.md)：模板填充机制
- [docs/csdn-reference-policy.md](docs/csdn-reference-policy.md)：CSDN 参考内容如何避免照抄
- [docs/screenshot-evidence.md](docs/screenshot-evidence.md)：截图证据如何嵌入报告
- [docs/one-click-demo.md](docs/one-click-demo.md)：一键演示流程
- [docs/course-design-fastline.md](docs/course-design-fastline.md)：课程设计报告快线
- [demo/README.md](demo/README.md)：演示素材和 2x2 布局预览
- [examples/README.md](examples/README.md)：示例文件总览

## 验证方式

项目展示材料检查：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\check-project-readiness.ps1
```

一键 demo：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-one-click-demo.ps1
```

完整烟测：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run-smoke-tests.ps1
```

独立运行新核心与通用端到端场景：

```powershell
python -m unittest discover -s tests -v
powershell -ExecutionPolicy Bypass -File .\tests\run-core-smoke.ps1
powershell -ExecutionPolicy Bypass -File .\tests\run-universal-e2e.ps1
```

可选 OpenClaw 智能通道检查：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\self-check.ps1
```

## 开源协作

仓库已包含 [CONTRIBUTING.md](CONTRIBUTING.md)、[CHANGELOG.md](CHANGELOG.md)、[CODE_OF_CONDUCT.md](CODE_OF_CONDUCT.md)、[SECURITY.md](SECURITY.md)、[SUPPORT.md](SUPPORT.md)，以及 `.github/` 下的 issue / PR 模板和 workflow。

如果你想公开发布，先看 [docs/GITHUB_LAUNCH.md](docs/GITHUB_LAUNCH.md)；如果要同步做小红书 / 抖音内容，看 [docs/social-launch-kit.md](docs/social-launch-kit.md)。

## 本地 Web UI

仓库包含 Gradio 界面，支持学生信息、课程与实验名称、正文长度、参考链接、对话材料、DOCX/DOC 模板、截图和代码文件。上传模板默认走保真模式，界面会展示阶段进度、模板契约、图片清单、格式检查、视觉检查和质量建议。

界面提供“无上传模板时使用”选项，可自动推荐或手工选择五套中性模板。只要上传了用户模板，该选择就会被忽略。

默认工作目录是仓库下的 `outputs/web-ui/`，不会绑定某个盘符。可通过环境变量或 JSON 配置覆盖：

```powershell
$env:EXPERIMENT_REPORT_TEMPLATE_PATH = "D:\templates\experiment.docx"
$env:EXPERIMENT_REPORT_BUILTIN_TEMPLATE_ID = "neutral-engineering-lab"
$env:EXPERIMENT_REPORT_OUTPUT_ROOT = "D:\report-output"
$env:EXPERIMENT_REPORT_CONFIG = "D:\config\experiment-report.json"
```

成功生成后的历史选项保存在 `outputs/web-ui/web-ui-history.json`。

Install the optional UI dependencies:

```powershell
python -m pip install -r requirements-web.txt
```

Start the UI:

```powershell
python web_ui.py
```

Then open:

```text
http://127.0.0.1:7860
```

快速模式适合日常生成；严格模式会尝试导出 PDF、渲染全部页面并运行视觉检查。LibreOffice / `soffice` 是推荐路径。WPS 或 Microsoft Word COM 默认关闭，如明确接受本机 Office 自动化，可设置：

```powershell
$env:EXPERIMENT_REPORT_ALLOW_OFFICE_COM = "1"
python web_ui.py
```

严格链路无法完成时，DOCX 和诊断文件仍会保留，但状态显示为 `needs-fix`。详见 [examples/web_demo.md](examples/web_demo.md)。

## License

MIT
