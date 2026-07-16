# network-dos：局域网搭建与 DOS 命令实验

这个案例适合计算机网络课程里的验证性实验：配置两台主机地址，使用 `ipconfig`、`ping`、`arp` 等命令记录网络连通性和邻居缓存。

## 输入文件

- `prompt.md`：从实验要求生成正文时使用的提示词
- `report.txt`：包含地址规划、命令记录、故障分析和结论的完整正文
- `metadata.json`：学生和实验基础信息
- `requirements.json`：正文校验规则和关键词
- `image-specs.json`：5 张拓扑图、终端记录和排障流程图的插入位置与图注
- `assets/`：可跨机器分发、无学校标识的确定性演示素材

## 可运行命令

这个案例可以直接复用仓库自带正文和截图跑完整 `docx` 流程：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-report.ps1 `
  -BuiltInTemplateId neutral-classic-lab `
  -ReportPath ".\examples\cases\network-dos\report.txt" `
  -MetadataPath ".\examples\cases\network-dos\metadata.json" `
  -ImageSpecsPath ".\examples\cases\network-dos\image-specs.json" `
  -RequirementsPath ".\examples\cases\network-dos\requirements.json" `
  -OutputDir ".\tests-output\network-dos-case" `
  -PipelineMode full `
  -DetailLevel long `
  -TemplateStyleMode preserve
```

## 预期输出

输出目录会包含最终 `docx`、字段映射、图片映射、图片放置计划和 layout check。打开最终文档时应重点检查：

- 顶部姓名、学号、课程名、实验名是否填入
- 5 张图是否插入到对应步骤、结果和故障分析附近
- 图注是否说明 `ipconfig`、`ping` 和 `arp` 的证据含义
- 并排截图是否清晰可读
