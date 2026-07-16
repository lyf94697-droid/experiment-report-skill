# 模板示例

这个目录放的是仓库内置、可直接引用的示例模板，主要用于：

- 一键演示
- 新用户上手
- 自定义模板前的字段结构参考

## 当前五套内置模板

| ID | 展示名称 | 模板文件 | 用途 |
| --- | --- | --- | --- |
| `neutral-classic-lab` | 经典实验报告 | `neutral-classic-lab.docx` | 普通课程实验默认模板 |
| `neutral-bordered-lab` | 闭合外框实验报告 | `neutral-bordered-lab.docx` | 强调页面闭合外框和传统版式 |
| `neutral-engineering-lab` | 工程技术实验报告 | `neutral-engineering-lab.docx` | 计算机、编程和工程类实验 |
| `neutral-course-design` | 课程设计报告 | `neutral-course-design.docx` | 课程设计、综合实验和项目报告 |
| `neutral-modern-minimal` | 现代简洁实验报告 | `neutral-modern-minimal.docx` | 不要求传统表格式的简洁报告 |

`experiment-report-template.docx` 和 `course-design-report-template.docx` 是兼容旧命令的别名，分别指向经典实验报告和课程设计报告的当前版本。

## 使用建议

- 有用户模板时始终优先使用用户模板，并保持其格式。
- 无模板时可传 `-BuiltInTemplateId neutral-engineering-lab` 等 ID。
- 不指定 ID 时，普通实验默认经典模板，课程设计默认课程设计模板。
- 真实学校模板适配前，先跑 `scripts/check-report-profile-template-fit.ps1`
- 内置模板修改后运行 `python -m universal_report audit-template-catalog --repo-root .`

## 来源与许可

五份 DOCX 都由 `scripts/build-neutral-templates.py` 从零生成，按仓库 MIT 许可证分发，不包含第三方图片或字体文件。`catalog.json` 记录公开参考链接和“只参考设计原则”的使用范围，不复制学校名称、校徽、示例正文或其他品牌资产。
