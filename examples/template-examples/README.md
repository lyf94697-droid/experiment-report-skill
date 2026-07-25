# 十套模板成品

这里的十份 DOCX 分别使用十套内置模板生成。每份都是正文、数据、结果证据和封面信息齐全的成品，不是空模板，也不是把同一篇报告重复套版。

姓名、学号、班级、教师、日期和实验数据均为虚构内容，使用时请替换成自己的真实信息与实验结果。成品正文不带“演示版”“示例版”字样，目录中也不使用学校名称、校徽或水印。

| 模板 | 报告主题 | 正文与信息 | 成品 |
| --- | --- | --- | --- |
| 经典实验报告 | 静态地址配置与双向连通性验证 | [`neutral-classic-lab`](neutral-classic-lab/) | [`01-经典实验报告-静态地址配置与双向连通性验证.docx`](docx/01-经典实验报告-静态地址配置与双向连通性验证.docx) |
| 闭合外框实验报告 | 进程调度算法比较 | [`neutral-bordered-lab`](neutral-bordered-lab/) | [`02-闭合外框实验报告-进程调度算法比较.docx`](docx/02-闭合外框实验报告-进程调度算法比较.docx) |
| 工程技术实验报告 | 复合索引对查询性能的影响 | [`neutral-engineering-lab`](neutral-engineering-lab/) | [`03-工程技术实验报告-复合索引对查询性能的影响.docx`](docx/03-工程技术实验报告-复合索引对查询性能的影响.docx) |
| 课程设计报告 | 学生成绩管理系统 | [正文](../cases/course-design-student-management/report.txt) · [信息](neutral-course-design/metadata.json) | [`04-课程设计报告-学生成绩管理系统.docx`](docx/04-课程设计报告-学生成绩管理系统.docx) |
| 现代简洁实验报告 | 响应式课程卡片布局 | [`neutral-modern-minimal`](neutral-modern-minimal/) | [`05-现代简洁实验报告-响应式课程卡片布局.docx`](docx/05-现代简洁实验报告-响应式课程卡片布局.docx) |
| 紧凑信息条实验报告 | Linux 文件权限与用户组管理 | [`neutral-compact-header-lab`](neutral-compact-header-lab/) | [`06-紧凑信息条实验报告-Linux文件权限与用户组管理.docx`](docx/06-紧凑信息条实验报告-Linux文件权限与用户组管理.docx) |
| 评阅记录实验报告 | JUnit 参数化测试与边界值验证 | [`neutral-review-panel-lab`](neutral-review-panel-lab/) | [`07-评阅记录实验报告-JUnit参数化测试与边界值验证.docx`](docx/07-评阅记录实验报告-JUnit参数化测试与边界值验证.docx) |
| 程序设计实验报告 | 归并排序实现与性能测试 | [`neutral-code-notebook-lab`](neutral-code-notebook-lab/) | [`08-程序设计实验报告-归并排序实现与性能测试.docx`](docx/08-程序设计实验报告-归并排序实现与性能测试.docx) |
| 实验数据分析报告 | 温度传感器标定与误差分析 | [`neutral-data-analysis-lab`](neutral-data-analysis-lab/) | [`09-实验数据分析报告-温度传感器标定与误差分析.docx`](docx/09-实验数据分析报告-温度传感器标定与误差分析.docx) |
| 项目技术报告 | 个人任务清单 Android 应用 | [`neutral-project-dossier`](neutral-project-dossier/) | [`10-项目技术报告-个人任务清单Android应用.docx`](docx/10-项目技术报告-个人任务清单Android应用.docx) |

重新生成全部成品：

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\build-template-examples.ps1
```

脚本会读取 [`catalog.json`](catalog.json)，先在 `tests-output/` 中完成填充和检查，再把十份最终 DOCX 复制到 `docx/`。
