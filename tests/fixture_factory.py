from __future__ import annotations

import zipfile
from pathlib import Path


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>
"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

DOCUMENT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"""


def _styles_xml(
    *,
    title_font: str,
    title_size: int,
    body_font: str,
    body_size: int,
    body_line: int,
    first_line: int,
) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="{body_font}"/>
        <w:sz w:val="{body_size}"/>
        <w:szCs w:val="{body_size}"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:line="{body_line}" w:lineRule="auto"/>
        <w:ind w:firstLine="{first_line}" w:firstLineChars="200"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ReportTitle">
    <w:name w:val="Report Title"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="0" w:after="240"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{title_font}" w:hAnsi="{title_font}" w:eastAsia="{title_font}"/>
      <w:b/>
      <w:sz w:val="{title_size}"/>
      <w:szCs w:val="{title_size}"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="HeadingOne">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="{body_size + 4}"/></w:rPr>
  </w:style>
</w:styles>
"""


def make_template(
    path: Path,
    *,
    columns: list[int] | None = None,
    title_font: str = "黑体",
    title_size: int = 44,
    body_font: str = "宋体",
    body_size: int = 24,
    body_line: int = 360,
    first_line: int = 480,
    margins: tuple[int, int, int, int] = (1440, 1440, 1440, 1440),
    page_size: tuple[int, int] = (11906, 16838),
    page_border: bool = False,
    sections: int = 1,
    blank: bool = False,
    image_placeholder: bool = False,
    outer_table_border: bool = True,
    long_course_name: bool = False,
    include_metadata_table: bool = True,
    empty_body: bool = False,
    body_table_after: bool = False,
    misleading_cover_lines: bool = False,
) -> Path:
    columns = [1800, 3000, 1800, 4200] if columns is None else columns
    title_text = "" if blank else "信息学院实验报告"
    course_value = "面向复杂工程系统的软件体系结构与综合实践" if long_course_name else "计算机网络"
    border_xml = ""
    if page_border:
        border_xml = """
        <w:pgBorders w:offsetFrom="page">
          <w:top w:val="single" w:sz="8" w:space="12" w:color="000000"/>
          <w:left w:val="single" w:sz="8" w:space="12" w:color="000000"/>
          <w:bottom w:val="single" w:sz="8" w:space="12" w:color="000000"/>
          <w:right w:val="single" w:sz="8" w:space="12" w:color="000000"/>
        </w:pgBorders>"""

    table_border_xml = ""
    if outer_table_border:
        table_border_xml = """
          <w:tblBorders>
            <w:top w:val="single" w:sz="4" w:color="000000"/>
            <w:left w:val="single" w:sz="4" w:color="000000"/>
            <w:bottom w:val="single" w:sz="4" w:color="000000"/>
            <w:right w:val="single" w:sz="4" w:color="000000"/>
            <w:insideH w:val="single" w:sz="4" w:color="000000"/>
            <w:insideV w:val="single" w:sz="4" w:color="000000"/>
          </w:tblBorders>"""

    metadata_table_xml = ""
    if include_metadata_table:
        grid = "".join(f'<w:gridCol w:w="{width}"/>' for width in columns)
        labels = ["姓名", "示例学生", "学号", "20260001", "班级", "计科2401", "课程名称", course_value]
        cells = []
        for index, width in enumerate(columns):
            text = labels[index] if index < len(labels) else ""
            cells.append(
                f"""<w:tc>
                  <w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>
                  <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>
                </w:tc>"""
            )
        metadata_table_xml = f"""<w:tbl>
      <w:tblPr>
        <w:tblW w:w="{sum(columns)}" w:type="dxa"/>
        <w:tblLayout w:type="fixed"/>
        {table_border_xml}
      </w:tblPr>
      <w:tblGrid>{grid}</w:tblGrid>
      <w:tr>{''.join(cells)}</w:tr>
    </w:tbl>"""
    placeholder = (
        '<w:p><w:r><w:t>[[IMAGE:实验结果]]</w:t></w:r></w:p>' if image_placeholder else ""
    )
    body_paragraph = (
        "<w:p><w:r/></w:p>"
        if empty_body
        else "<w:p><w:r><w:t>掌握实验环境配置、关键操作步骤和结果验证方法。</w:t></w:r></w:p>"
    )
    body_table_xml = (
        """<w:tbl>
      <w:tblPr><w:tblW w:w="6000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>模块</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>说明</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p><w:r><w:rPr><w:rFonts w:eastAsia="微软雅黑"/><w:sz w:val="19"/></w:rPr><w:t>图1 系统结构图</w:t></w:r></w:p>"""
        if body_table_after
        else ""
    )
    cover_lines_xml = (
        """<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>COURSE DESIGN REPORT TEMPLATE</w:t></w:r></w:p>
    <w:p><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>课题名称：</w:t></w:r></w:p>"""
        if misleading_cover_lines
        else ""
    )
    section_xml = []
    for index in range(sections):
        title_page = "<w:titlePg/>" if index == 0 and sections > 1 else ""
        section_xml.append(
            f"""<w:p><w:pPr><w:sectPr>
              <w:type w:val="nextPage"/>
              <w:pgSz w:w="{page_size[0]}" w:h="{page_size[1]}"/>
              <w:pgMar w:top="{margins[0]}" w:right="{margins[1]}" w:bottom="{margins[2]}" w:left="{margins[3]}" w:header="720" w:footer="720"/>
              {title_page}
              {border_xml}
            </w:sectPr></w:pPr></w:p>"""
        )

    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="ReportTitle"/></w:pPr>
      <w:r><w:t>{title_text}</w:t></w:r>
    </w:p>
    {cover_lines_xml}
    {metadata_table_xml}
    <w:p><w:pPr><w:pStyle w:val="HeadingOne"/></w:pPr><w:r><w:t>实验目的</w:t></w:r></w:p>
    {body_paragraph}
    {body_table_xml}
    {placeholder}
    {''.join(section_xml)}
  </w:body>
</w:document>
"""

    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", CONTENT_TYPES)
        archive.writestr("_rels/.rels", ROOT_RELS)
        archive.writestr("word/_rels/document.xml.rels", DOCUMENT_RELS)
        archive.writestr(
            "word/styles.xml",
            _styles_xml(
                title_font=title_font,
                title_size=title_size,
                body_font=body_font,
                body_size=body_size,
                body_line=body_line,
                first_line=first_line,
            ),
        )
        archive.writestr("word/document.xml", document)
    return path
