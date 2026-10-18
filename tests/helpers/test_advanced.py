"""独立读取 DOCX，核对行距、编号可见结果、原有段距和表格。"""
import json
import sys
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt
from native_number_text import native_visible_numbers
from test_native_display import effective

mode, folder = sys.argv[1], Path(sys.argv[2])
if mode == 'fixture':
    doc = Document()
    doc.core_properties.author = "Synthetic test fixture"
    doc.core_properties.last_modified_by = ""
    doc.styles['Normal'].font.size = Pt(12)
    for text in ('普通正文不改。', '正文 $a+b$ 和 $c+d$ 同行。', r'\[\frac{x+1}{y+2}\]', '$$x+1$$', '结尾正文不改。'):
        p = doc.add_paragraph(text)
        f = p.paragraph_format
        f.line_spacing = Pt(20)
        f.space_before, f.space_after = Pt(6), Pt(9)
        f.left_indent, f.right_indent = Pt(12), Pt(6)
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = '$t_1$ 不转换'
    # Word 要求表格之后有末段，样本直接提供，避免把自动末段算作转换拆段。
    doc.paragraphs[-1]._p.addprevious(table._tbl)
    doc.save(folder / 'advanced-input.docx')
    print('ADVANCED_FIXTURE_PASSED')
    sys.exit()

source, target = Document(folder / 'advanced-input.docx'), Document(folder / (mode + '.docx'))
report = json.loads((folder / (mode + '.json')).read_text(encoding='utf-8-sig'))
assert report['Status'] == 'COMPLETE' and report['Succeeded'] == 4
assert len(source.paragraphs) == len(target.paragraphs) == 5
assert target.tables[0].cell(0, 0).text == source.tables[0].cell(0, 0).text

def spacing(p):
    value = p._p.find('./' + qn('w:pPr') + '/' + qn('w:spacing'))
    # 原生显示段落会将相同属性归入样式，按有效属性比较而非只读直接格式。
    style = p.style
    while value is None and style is not None:
        value = style.element.find('./' + qn('w:pPr') + '/' + qn('w:spacing'))
        style = style.base_style
    return value.get(qn('w:lineRule'), 'auto'), int(value.get(qn('w:line')))

rules = {
    'follow': ('exact', 400), 'single': ('auto', 240), 'onehalf': ('auto', 360),
    'double': ('auto', 480), 'exact:24.000000': ('exact', 480),
    'atleast:28.000000': ('atLeast', 560), 'multiple:1.250000': ('auto', 300),
}
for index, (before, after) in enumerate(zip(source.paragraphs, target.paragraphs)):
    for name in ('space_before', 'space_after', 'left_indent', 'right_indent'):
        assert effective(before, name) == effective(after, name), (mode, index, name)
    if index in (0, 4):
        assert before.text == after.text and spacing(before) == spacing(after)
    else:
        setting = report['InlineSpacing' if index == 1 else 'DisplaySpacing']
        assert spacing(after) == rules[setting], (mode, index, spacing(after), setting)
        if setting != 'follow':
            assert after._p.find('./' + qn('w:pPr') + '/' + qn('w:snapToGrid')).get(qn('w:val')) == '0'

expected = {
    'alphabetic': ['[IV.b-AA]', '[IV.b-BB]'],
    'roman': ['ix', 'x'], 'numeric': ['（3-2-9）', '（3-2-10）'],
}
visible = ''.join(native_visible_numbers(target.element.body))
for number in expected.get(mode, []):
    assert number in visible, (number, visible)
assert target.paragraphs[1].text.startswith('正文 ') and target.paragraphs[1].text.endswith(' 同行。')
print(json.dumps({'Status': 'PASSED', 'Case': mode, 'ParagraphSpacingVerified': True,
                  'UnselectedContentUnchanged': True, 'NumberTextVerified': expected.get(mode, [])}, ensure_ascii=False))
