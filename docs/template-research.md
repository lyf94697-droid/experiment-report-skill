# 实验报告模板调研与采用边界

## 目标

本调研用于回答两个问题：

1. 中文高校实验报告和课程设计报告中，哪些版式结构最常见、最值得做成内置模板。
2. 如何吸收公开模板的设计原则，同时不把学校名称、校徽、示例学生身份、示例正文或第三方资产带入仓库。

内置 DOCX 全部由 `scripts/build-neutral-templates.py` 从零生成。公开网页和 GitHub 项目只作为结构研究来源，不作为文件复制来源。

## 公开来源

高校公开页面：

- [宁夏大学信息工程学院实验报告模板](https://xxgc.nxu.edu.cn/info/1011/1208.htm)：常见字段包括课程、实验项目、地点、班级、姓名、学号、教师和日期；正文包含目的、环境、步骤、结果分析和教师审阅。
- [四川大学生物实验教学中心实验报告模板](https://biosci.lab.scu.edu.cn/info/2283/1087.htm)：目的、材料方法、结果记录和分析讨论相互分离。
- [中北大学通用实验报告模板](https://hjaq.nuc.edu.cn/info/1083/5956.htm)：展示通用实验报告的下载和教学归档场景。
- [北京航空航天大学课程设计报告模板](https://dnc.buaa.edu.cn/info/1051/1473.htm)：课程设计通常采用独立封面和长篇正文。
- [中南大学课程设计报告要求](https://dgdz.csu.edu.cn/pract_edu/kcsj/kcsj_report.html)：公开要求覆盖封面、前言或摘要、目录、章节和参考文献。
- [焦作大学信息工程学院实验报告说明](https://xxgc.jzu.edu.cn/info/1117/1612.htm)：教师评分、评语和签字是常见教学归档字段。
- [WPS 报告模板分类](https://template.wps.com/themes/report-94/)：用于确认通用报告对信息表、数据表和分析结论的常见版面分工；仓库不使用站点素材。

开源项目：

- [megrxu/zjureport](https://github.com/megrxu/zjureport)：研究紧凑正文页信息条。
- [xiaoxinganling/WHUExperiment](https://github.com/xiaoxinganling/WHUExperiment)：研究封面、元数据和正文分离。
- [LuminolT/SHU-Lab-Report-Template](https://github.com/LuminolT/SHU-Lab-Report-Template)：研究表格式实验记录与教师评阅结构。
- [StarHub-SPA/Experiment_Report_SPA_Template](https://github.com/StarHub-SPA/Experiment_Report_SPA_Template)：研究闭合页面框架和正文区域。
- [Su-anAcid/CQUPT-Course-Report](https://github.com/Su-anAcid/CQUPT-Course-Report)：研究程序设计类报告对实现、测试和结果的内容分工。
- [pisceskkk/NUDT_ExperimentReportTemplate](https://github.com/pisceskkk/NUDT_ExperimentReportTemplate)：研究长篇技术报告的封面与正文分离。
- [Typst simple-bupt-report](https://typst.app/universe/package/simple-bupt-report/)：研究中文技术报告的标题、正文、图注和代码字号层级。

## 归纳出的常见结构

调研中重复出现的结构不是单一“学校风格”，而是以下可组合的文档角色：

1. 独立封面：适合课程设计、综合项目和长篇技术报告。
2. 紧凑信息条：适合周实验、短实验和频繁提交。
3. 四列或六列元数据表：标签列固定宽度，值列承担长课程名和实验名。
4. 闭合页面外框：适合纸质提交和档案式实验记录。
5. 教师评阅区：成绩、评语、签名和日期与学生正文分离。
6. 代码与命令区：等宽字体、浅色底、测试和调试记录独立。
7. 数据记录表：原始值、处理值、备注与后续误差分析分离。
8. 长篇项目结构：摘要、关键词、需求、设计、实现、测试、总结、参考文献和附录。

常见排版参数集中在 A4、约 2–3 cm 页边距、宋体 12 pt 左右正文、黑体 12–15 pt 标题、1.2–1.5 倍行距、首行缩进约 2 个汉字、10.5 pt 左右图表题。仓库模板根据用途在该范围内调整，但不声称复刻某一学校的专有规范。

## 十套模板覆盖

| 模板 ID | 主要覆盖的常见结构 |
| --- | --- |
| `neutral-classic-lab` | 四列信息表、传统章节、通用实验 |
| `neutral-bordered-lab` | 闭合页面外框、纸质提交 |
| `neutral-engineering-lab` | 工程技术层级、图表与配置说明 |
| `neutral-course-design` | 普通课程设计、封面与正文分节 |
| `neutral-modern-minimal` | 低装饰、充分留白、轻量报告 |
| `neutral-compact-header-lab` | 六列紧凑信息条、周实验与短实验 |
| `neutral-review-panel-lab` | 成绩列、教师评语、签名与日期 |
| `neutral-code-notebook-lab` | 代码、命令输出、测试和调试 |
| `neutral-data-analysis-lab` | 数据表、趋势、误差与不确定性 |
| `neutral-project-dossier` | 长篇系统项目、摘要与参考文献 |

## 不采用的内容

- 不复制真实学校名称、学院名称、校徽、页眉品牌和水印。
- 不复制公开模板中的示例姓名、学号、课程名、实验数据和正文。
- 不复制来源仓库的图片、字体文件、Logo、LaTeX 类文件或 DOCX 文件。
- 不把单个学校的专有字段强行当作通用规则。
- 不承诺任意复杂模板“像素级零差异”。上传模板会优先进入保真模式，但浮动文本框、复杂域、宏、受保护文档和特殊 Word/WPS 渲染差异仍可能需要人工复核。

## 自动验收

- `python -m universal_report audit-template-catalog --repo-root .`：检查模板数量、来源字段、学校标识、长数字身份和嵌入媒体。
- `tests/run-neutral-template-catalog.ps1`：将真实正文填入十套内置模板，要求章节全部映射且格式契约通过。
- `tests/run-template-fidelity-corpus.ps1`：将五种结构差异最大的模板当作用户上传文件再次填充，要求用户模板优先、保真模式、页面/字体/段落/表格检查通过，并核对样式、编号、设置、主题、页眉和页脚部件未被改写。
- 严格模式：继续执行 DOCX 转 PDF、逐页渲染和视觉检查；缺少转换器时不得把未完成视觉验收的结果标为完全通过。
