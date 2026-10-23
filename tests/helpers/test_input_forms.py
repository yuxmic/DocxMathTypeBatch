"""核对不对称定界符、百分号和非 OLE 嵌入包，并保留同文档内的失败原式。"""
import json
import sys
from collections import Counter
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import Part
from docx.oxml import parse_xml
from docx.shared import Pt
from audit_docx import audit


def fixture(folder):
    document = Document()
    document.core_properties.author = 'Synthetic test fixture'
    document.core_properties.last_modified_by = ''
    for value in (r'[xy\]', r' [x\] ', r'[z_{1}\]', '$i_1$', '$$d_1$$', r'\(p_1\)', r'\[q_1\]', r'[\]', 'End stays exact.'):
        document.add_paragraph(value)
    document.save(folder / 'reverse-input-forms.docx')

    # 真正挂接图表与工作簿关系，避免 Word 清理没有引用的合成嵌入部件。
    spreadsheet = BytesIO()
    with ZipFile(spreadsheet, 'w', ZIP_DEFLATED) as package:
        package.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>')
        package.writestr('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        package.writestr('xl/workbook.xml', '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>')
        package.writestr('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>')
        package.writestr('xl/worksheets/sheet1.xml', '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Value</t></is></c><c r="B1"><v>25</v></c></row></sheetData></worksheet>')
    document = Document()
    document.core_properties.author = 'Synthetic test fixture'
    document.core_properties.last_modified_by = ''
    document.styles['Normal'].font.size = Pt(12)
    document.add_paragraph(r'\[50\%\]')
    document.add_paragraph(r'Keep unsupported $\text{汉字}$ exact.')
    document.add_paragraph(r'\[unfinished')
    document.add_table(rows=1, cols=1).cell(0, 0).text = r'Cell $25\%$ and $t_1$ stay here.'
    workbook = Part(PackURI('/word/embeddings/input-forms.xlsx'), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', spreadsheet.getvalue(), document.part.package)
    chart = Part(PackURI('/word/charts/chart1.xml'), 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml', b'''<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
      <c:chart><c:plotArea><c:pieChart><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Value</c:v></c:pt></c:strCache></c:strRef></c:cat><c:val><c:numRef><c:f>Sheet1!$B$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="1"/><c:pt idx="0"><c:v>25</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:pieChart></c:plotArea><c:plotVisOnly val="1"/></c:chart>
      <c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>''', document.part.package)
    assert chart.relate_to(workbook, RT.PACKAGE) == 'rId1'
    relation = document.part.relate_to(chart, RT.CHART)
    document.add_paragraph().add_run()._r.append(parse_xml(f'''<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="2743200" cy="1828800"/><wp:docPr id="1" name="Synthetic chart"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="{relation}"/></a:graphicData></a:graphic></wp:inline></w:drawing>'''))
    document.add_paragraph('End stays exact.')
    document.save(folder / 'embedded-input-forms.docx')


def check(folder):
    for delimiter in ('bracket', 'dollar'):
        path = folder / ('reverse-forms-' + delimiter + '.docx')
        paragraphs = [paragraph.text for paragraph in Document(path).paragraphs]
        opening, closing = (r'\[', r'\]') if delimiter == 'bracket' else ('$$', '$$')
        assert paragraphs == [opening+'xy'+closing, ' '+opening+'x'+closing+' ', opening+'z_{1}'+closing, '$i_1$', opening+'d_1'+closing, '$p_1$', opening+'q_1'+closing, r'[\]', 'End stays exact.'], paragraphs
        report = json.loads(path.with_suffix('.json').read_text(encoding='utf-8-sig'))
        assert report['Succeeded'] == 7 and report['RejectedPreserved'] == 1 and report['Status'] == 'PARTIAL'
    source = folder / 'embedded-input-forms.docx'
    with ZipFile(source) as package:
        embeddings = Counter(package.read(name) for name in package.namelist() if name.startswith('word/embeddings/'))
    for name in ('embedded-forms', 'embedded-forms-extra', 'embedded-forms-reverse'):
        path = folder / (name + '.docx')
        report = json.loads(path.with_suffix('.json').read_text(encoding='utf-8-sig'))
        document = Document(path)
        assert report['Status'] == 'PARTIAL' and report['WordReopenVerified'] and report['OriginalFileBytesUnchanged']
        assert document.paragraphs[1].text == r'Keep unsupported $\text{汉字}$ exact.'
        assert document.paragraphs[2].text == r'\[unfinished' and document.paragraphs[-1].text == 'End stays exact.'
        assert len(document.element.xpath('//*[local-name()="chart"]')) == 1
        with ZipFile(path) as package:
            assert Counter(package.read(part) for part in package.namelist() if part.startswith('word/embeddings/') and part.endswith('.xlsx')) == embeddings
        result = audit(path)
        if name == 'embedded-forms-reverse':
            assert result['MathType'] == 0 and report['ExportedMathType'] == 1
            assert r'50\%' in document.paragraphs[0].text and r'25\%' in document.tables[0].cell(0, 0).text
        else:
            assert report['FailedPreserved'] == 1 and report['RejectedPreserved'] == 1
            assert report['Succeeded'] == result['MathType'] == (1 if name == 'embedded-forms' else 3)
            if name == 'embedded-forms':
                assert document.tables[0].cell(0, 0).text == r'Cell $25\%$ and $t_1$ stay here.'
    print('INPUT_FORMS_AND_RETAINED_EMBEDDINGS_PASSED')


if __name__ == '__main__':
    mode, directory = sys.argv[1:3]
    (fixture if mode == 'fixture' else check)(Path(directory))
