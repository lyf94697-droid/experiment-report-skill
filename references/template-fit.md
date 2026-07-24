# Template Fit Mode

## 1. Template-first mapping

- 如果用户要生成 DOCX，但尚未说明是否有模板，先问一次：“你有老师、学校或自己认可的优秀 DOCX/DOC 模板吗？”
- 用户提供模板时，它拥有最高优先级，不与内置模板混搭。
- 用户明确没有模板时，从 `examples/report-templates/catalog.json` 的十套中性模板中选择，不再反复追问。
- 先识别模板里的固定标题、编号、表格字段、封面字段。
- 输出时按模板顺序组织，不要擅自改编号。
- 如果模板标题和常规模板不同，优先服从模板。

## 2. WPS or Word handling

- 本地 WPS 桌面自动化不是主路径，先完成内容，再尝试填写。
- 如果模板是本地 docx，优先先跑 `scripts/extract-docx-template.ps1 -Path <template.docx>`，用提取出的段落顺序、表格单元格和疑似字段做映射。
- 如果正文已经写完且需要机器生成字段映射，优先跑 `scripts/generate-docx-field-map.ps1`，不要手工拼大段 JSON。
- `scripts/generate-docx-field-map.ps1` 返回的 JSON 里如果有 `diagnostics`，优先根据 `code`、`message` 和 `suggestion` 判断是缺 metadata、缺章节内容，还是模板需要补 `fieldMapCompositeRules` / `sectionFields` alias。
- 如果你是在给一个新模板做 profile onboarding，优先改跑 `scripts/check-report-profile-template-fit.ps1`，它会把 field-map diagnostics 归并成 profile change suggestions 和 input gaps，方便直接补 profile JSON。
- 如果用户明确需要机器生成一个已填内容的 docx 副本，再用 `scripts/apply-docx-field-map.ps1` 执行回填。
- 标签键映射适合保守填空，只会填充空白位或占位符。
- 如果模板里是固定章节标题加空白正文段，优先给出 `paragraphs` 数组。
- 需要保留标题、把正文写到下一段时，用 `mode: "after"`。
- 位置键映射适合显式覆盖，比如 `P2`、`T1R1C2`。
- 如果模板是本地 docx 且可编辑，优先按字段填充。
- 如果只有截图或用户只给了空白模板界面，输出“字段 -> 内容”映射，便于直接粘贴。
- 若没有可靠自动化，就明确说明返回的是正文和字段映射，而不是虚构“已自动填好模板”。

## 3. Field mapping guidance

- 封面字段：课程名、实验名、姓名、学号、班级、日期分开写。
- 表格字段：优先输出单元格级内容，不要把整段话塞进短字段。
- 正文大段落：优先按段落数组输出，避免把多段正文压成一个字段。
- 对“实验目的 / 实验步骤 / 实验结果”这类固定标题，尽量保持标题原样，只映射后续正文。

## 4. Stop conditions

- 模板内容或实验要求缺失到无法保证真实性时，先补信息再继续。
- 若桌面模板填写能力不可用，仍需完成整份报告正文和字段映射，不得直接中断。

## 5. Tutorial article plus screenshots mode

- When the user gives a tutorial article plus their own screenshots or results, treat the article as procedural reference and the screenshots or results as factual evidence.
- Fill missing explanatory sections from the article, but keep the result section aligned with the user's actual outputs.
- If the article includes code and the user has not provided their own code, mark it as reference implementation instead of pretending it is the user's exact work.

## 6. Ten neutral built-in templates

- `neutral-classic-lab`：经典四列信息表，适合普通课程实验。
- `neutral-bordered-lab`：闭合页面外框，适合强调传统纸质版式的报告。
- `neutral-engineering-lab`：技术层级、代码和图注样式更明确，适合计算机与工程类实验。
- `neutral-course-design`：独立封面和正文分节，适合课程设计、综合实验和项目报告。
- `neutral-modern-minimal`：两列元数据和轻量层级，适合现代简洁风格。
- `neutral-compact-header-lab`：六列紧凑信息条，适合周实验、短实验和快速记录。
- `neutral-review-panel-lab`：成绩、教师评语和签名区齐全，适合纸质批阅与归档。
- `neutral-code-notebook-lab`：代码、测试、输出和调试记录层级明确，适合程序设计实验。
- `neutral-data-analysis-lab`：原始值、处理值、趋势和误差分析分离，适合测量与统计实验。
- `neutral-project-dossier`：封面、摘要、长篇正文和参考文献齐全，适合系统项目。

自动选择顺序：

1. 长篇系统项目 → `neutral-project-dossier`；普通课程设计 → `neutral-course-design`
2. 代码、测试或调试实验 → `neutral-code-notebook-lab`
3. 数据记录、测量、统计或误差分析 → `neutral-data-analysis-lab`
4. 教师评语、成绩或签名归档 → `neutral-review-panel-lab`
5. 周实验、短实验或紧凑记录 → `neutral-compact-header-lab`
6. 明确要求闭合外框 → `neutral-bordered-lab`
7. 明确要求现代或极简 → `neutral-modern-minimal`
8. 计算机网络、操作系统、数据库、Web、Android 或软件工程 → `neutral-engineering-lab`
9. 其他实验 → `neutral-classic-lab`

## 7. Neutrality and provenance gate

- 内置模板不得出现真实学校、学院、专业组织名称、校徽、校训、水印、地址、二维码、示例学生姓名或长数字学号。
- 内置模板不得携带第三方图片、字体文件或复制的示例正文。
- 公共模板只用于参考信息表比例、章节层级、分节和排版节奏。
- 每次修改或重建内置模板后运行：

```powershell
python -m universal_report audit-template-catalog --repo-root .
```

- 审计不通过时，不得把对应模板作为自动回退模板。
