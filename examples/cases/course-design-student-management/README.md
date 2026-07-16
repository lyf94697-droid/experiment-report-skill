# course-design-student-management：学生成绩管理系统课程设计

这个案例适合课程设计报告：题目不是单次验证实验，而是一个小型系统的需求分析、总体设计、模块实现、运行结果和总结。

## 输入文件

- `prompt.md`：课程设计报告生成约束
- `report.txt`：可直接进入课程设计模板填充的示例正文
- `metadata.json`：课程设计封面和基础字段
- `requirements.json`：课程设计报告校验规则
- `image-specs.json`：6 张架构、数据模型、界面、统计和测试图片的精确锚点
- `assets/`：无学校名称、校徽或第三方品牌的可交付视觉素材

## 可运行命令

课程设计必须使用 `course-design-report` profile：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -BuiltInTemplateId neutral-course-design `
  -ReportPath ".\examples\cases\course-design-student-management\report.txt" `
  -MetadataPath ".\examples\cases\course-design-student-management\metadata.json" `
  -ImageSpecsPath ".\examples\cases\course-design-student-management\image-specs.json" `
  -RequirementsPath ".\examples\cases\course-design-student-management\requirements.json" `
  -ReportProfileName course-design-report `
  -OutputDir ".\tests-output\course-design-student-management-case" `
  -PipelineMode full `
  -DetailLevel long `
  -TemplateStyleMode preserve
```

## 适配建议

案例已包含与正文数据一致的系统截图、统计图和测试记录，并自动生成与业务模型一致的功能模块表、数据库表和核心字段表。替换题目时，应同时更新正文、图片和数据库表结构。
