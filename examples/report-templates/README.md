# 模板示例

这个目录放的是仓库内置、可直接引用的示例模板，主要用于：

- 一键演示
- 新用户上手
- 自定义模板前的字段结构参考

## 当前十套内置模板

| ID | 展示名称 | 模板文件 | 用途 |
| --- | --- | --- | --- |
| `neutral-classic-lab` | 经典实验报告 | `neutral-classic-lab.docx` | 普通课程实验默认模板 |
| `neutral-bordered-lab` | 闭合外框实验报告 | `neutral-bordered-lab.docx` | 强调页面闭合外框和传统版式 |
| `neutral-engineering-lab` | 工程技术实验报告 | `neutral-engineering-lab.docx` | 网络、操作系统、数据库、Web 和工程类实验 |
| `neutral-course-design` | 课程设计报告 | `neutral-course-design.docx` | 课程设计、综合实验和项目报告 |
| `neutral-modern-minimal` | 现代简洁实验报告 | `neutral-modern-minimal.docx` | 不要求传统表格式的简洁报告 |
| `neutral-compact-header-lab` | 紧凑信息条实验报告 | `neutral-compact-header-lab.docx` | 周实验、短实验和快速记录 |
| `neutral-review-panel-lab` | 评阅记录实验报告 | `neutral-review-panel-lab.docx` | 带成绩、教师评语和签名区的纸质归档 |
| `neutral-code-notebook-lab` | 程序设计实验报告 | `neutral-code-notebook-lab.docx` | 代码、测试、运行结果和调试记录 |
| `neutral-data-analysis-lab` | 实验数据分析报告 | `neutral-data-analysis-lab.docx` | 原始数据、处理值、趋势和误差分析 |
| `neutral-project-dossier` | 项目技术报告 | `neutral-project-dossier.docx` | 长篇系统设计、课程项目与参考文献 |

`experiment-report-template.docx` 和 `course-design-report-template.docx` 是兼容旧命令的别名，分别指向经典实验报告和课程设计报告的当前版本。

每套模板填入完整正文后的效果见 [十套模板成品](../template-examples/README.md)。这些成品分别使用网络、操作系统、数据库、Web、Linux、JUnit、Python、数据分析和 Android 等不同主题，不是同一篇正文的重复套版。

## 使用建议

- 有用户模板时始终优先使用用户模板，并保持其格式。
- 无模板时可传 `-BuiltInTemplateId neutral-engineering-lab` 等 ID。
- 不指定 ID 时，普通实验默认经典模板，课程设计默认课程设计模板。
- 真实学校模板适配前，先跑 `scripts/check-report-profile-template-fit.ps1`
- 内置模板修改后运行 `python -m universal_report audit-template-catalog --repo-root .`

## 来源与许可

十份 DOCX 都由 `scripts/build-neutral-templates.py` 从零生成，按仓库 MIT 许可证分发，不包含第三方图片或字体文件。`catalog.json` 和 `../../docs/template-research.md` 记录公开参考链接和“只参考设计原则”的使用范围，不复制学校名称、校徽、示例正文或其他品牌资产。

`tests/run-neutral-template-catalog.ps1` 会把真实正文填入全部十套模板；`tests/run-template-fidelity-corpus.ps1` 会把五种差异最大的版式当作用户上传模板再次填充，并核对页面、字体、段落、表格网格、页眉页脚和 DOCX 样式部件是否保持。
