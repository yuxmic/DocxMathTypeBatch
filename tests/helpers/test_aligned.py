"""核对 aligned 的实际 MathType 导出、完整行列与 PDF 中的红色残留。"""
import json
import re
import sys
from pathlib import Path
import fitz
from audit_docx import audit

target = Path(sys.argv[1])
forward = json.loads(target.with_suffix('.json').read_text(encoding='utf-8-sig'))
reverse = json.loads(target.with_name('aligned-reverse.json').read_text(encoding='utf-8-sig'))
assert forward['Status'] == reverse['Status'] == 'COMPLETE'
assert forward['Succeeded'] == reverse['Succeeded'] == 7
assert sum(item['Display'] for item in forward['Items']) == 6
assert audit(target)['MathType'] == 7
assert all(float(item['NativeFullSize']) == 12 for item in forward['Items'])
expected = [
    ('a_1', 'b_1', 'a_2', 'b_2'),
    ('P_i', 'Q_i', 'V_i', 'V_j', 'P_2', 'Q_2', 'P_3', 'Q_3', r'\sin', r'\cos'),
    ('c_1', 'd_1', 'c_2', 'd_2'),
    ('e_1', 'm_1', 'm_2', 'm_3', 'm_4'),
    ('f_1', *(f'n_{i}' for i in range(1, 9))),
    tuple(f'q_{i}' for i in range(1, 5)),
    tuple(f'u_{i}' for i in range(1, 5)),
]
for item, tokens in zip(reverse['Items'], expected):
    tex = re.sub(r'\s|[{}]', '', item['OutputText'])
    assert 'aligned' not in tex, tex
    for token in tokens:
        assert token in tex, (token, tex)

def red(color):
    return color is not None and color[0] > .6 and color[1] < .3 and color[2] < .3

# 使用 Word 导出的页面验证，不能仅靠 API 返回成功判断不再标红。
with fitz.open(str(target) + '.pdf') as doc:
    for page in doc:
        for block in page.get_text('dict')['blocks']:
            for line in block.get('lines', []):
                for span in line['spans']:
                    color = span['color']
                    assert not red(((color >> 16 & 255) / 255, (color >> 8 & 255) / 255, (color & 255) / 255)), span
        for drawing in page.get_drawings():
            assert not red(drawing['color']) and not red(drawing['fill']), '公式有红色残留'
print(json.dumps(dict(Status='PASSED', MathType=7, Inline=1, Display=6,
                     AllRowsAndColumnsPresent=True, RedResiduals=False), ensure_ascii=False))
