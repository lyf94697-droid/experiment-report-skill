# Web UI Demo

This example shows how to start the local Web UI and generate DOCX, PDF, and preview PNG artifacts from browser-uploaded materials.

DOCX is the primary output. PDF and preview PNG are optional convenience outputs: the UI first tries LibreOffice / `soffice` for stable headless PDF export. WPS/Microsoft Word COM fallback is disabled by default because it can hang when Office is busy.

## Start

Install the optional UI dependencies:

```powershell
python -m pip install -r requirements-web.txt
```

Start the UI:

```powershell
python web_ui.py
```

Open:

```text
http://127.0.0.1:7860
```

## Inputs

Fill these fields:

- 报告类型
- 生成方式
- 正文长度
- 课程名称
- 常用学生资料
- 学生姓名
- 学号
- 班级
- 实验名称/题目名称
- 实验要求
- 参考链接或补充说明
- 对话式需求
- 本地截图文件夹/文件路径
- 本地代码文件夹/文件路径

Upload:

- an optional `.docx` or `.doc` template
- one or more result screenshots
- one or more code files

If no template is uploaded, experiment reports use `E:\实验报告\00-模板\实验报告模版1.docx` when available. Course-design reports use `E:\新建文件夹\课程设计-模板.doc` when available.

The student profile dropdown starts with `示例学生 / 20260001 / 计科2401`. Course name, experiment name, student name, ID, class, and output root are editable dropdowns: click them to pick local history, or type a new value. Successful generations update `outputs/web-ui/web-ui-history.json` for the next session.

The chat-style box can accept text like:

```text
CSDN链接：https://example.com/article
课程名称：计算机网络
实验名称：根据教程链接填充
姓名：示例学生
学号：20260001
班级：计科2401
截图材料："E:\实验报告\截图\计网实验六"
```

Manual fields take priority. Empty fields are filled from the chat-style text where possible.

## Output

Click `生成报告`. The page shows:

- generation status
- warnings or errors
- a DOCX download button
- a PDF download button
- a preview PNG download button
- an in-page preview image

Generated artifacts are copied to the selected output root, defaulting to:

```text
E:\实验报告\docx
E:\实验报告\pdf
E:\实验报告\预览图
```

The working files are also kept under:

```text
outputs/web-ui/
```

## PDF And Errors

- For stable PDF export, install LibreOffice and make sure `soffice` is available on `PATH`.
- If you want to allow WPS/Word fallback, close WPS/Word windows first and start the UI with:

```powershell
$env:EXPERIMENT_REPORT_ALLOW_OFFICE_COM = "1"
python web_ui.py
```

- If PDF export fails, the DOCX still succeeds and remains downloadable.
- Generation failures are shown with a short human-readable cause, the likely next step, and a log path under the working output directory.

## Notes

- `快速本地草稿` is the default mode for stable everyday runs.
- `智能长文（接近对话效果）` uses the local OpenClaw chat gateway when available.
- `质量模式` defaults to `快速生成`; choose `严格检查` for new or visually risky templates so the UI also checks metadata-table readability.
- If the chat gateway is unavailable, the UI falls back to `快速本地草稿` and shows the reason in the warning box.
- PDF export prefers LibreOffice. WPS/Microsoft Word fallback is opt-in with `EXPERIMENT_REPORT_ALLOW_OFFICE_COM=1`. Preview PNG rendering uses PyMuPDF.
