# os-process-scheduling：进程调度算法实验

这个案例适合操作系统课程里的算法实验：实现或模拟先来先服务、短作业优先、时间片轮转等调度策略，并分析等待时间、周转时间和调度结果。

## 输入文件

- `prompt.md`：生成正文时的实验事实和写作约束
- `report.txt`：可直接进入模板填充的示例正文
- `metadata.json`：学生和实验基础信息
- `requirements.json`：正文校验规则
- `image-specs.json`：5 张算法流程、甘特图、指标、控制台和测试素材的放置规则
- `assets/`：可直接随案例分发的无品牌图表与运行记录

## 可运行命令

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -BuiltInTemplateId neutral-engineering-lab `
  -ReportPath ".\examples\cases\os-process-scheduling\report.txt" `
  -MetadataPath ".\examples\cases\os-process-scheduling\metadata.json" `
  -ImageSpecsPath ".\examples\cases\os-process-scheduling\image-specs.json" `
  -RequirementsPath ".\examples\cases\os-process-scheduling\requirements.json" `
  -OutputDir ".\tests-output\os-process-scheduling-case" `
  -PipelineMode full `
  -DetailLevel long `
  -TemplateStyleMode preserve
```

## 适配建议

案例已经包含完整图表和运行记录，可直接生成成品。替换为个人实验数据时，应同步替换甘特图、性能指标和测试结论，不要把公开示例输出当作本次实验的真实结果。
