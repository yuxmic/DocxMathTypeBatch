"""定向核对四类小样及开发探针读出的实际 MTEF 内容，不进入生产路径。"""
import json
from pathlib import Path
import re
import sys
import zipfile
from lxml import etree as ET
from native_number_text import native_visible_numbers

source, output, native = map(Path, sys.argv[1:4])
expected = json.loads(source.with_suffix('.expected.json').read_text(encoding='utf-8'))
actual = json.loads(native.read_text(encoding='utf-8'))
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
      'o': 'urn:schemas-microsoft-com:office:office'}


def read(path):
    with zipfile.ZipFile(path) as archive:
        return ET.fromstring(archive.read('word/document.xml'))


def structure(node):
    return node.tag, sorted(node.attrib.items()), node.text or '', [structure(c) for c in node]


def normalized(tex):
    tex = re.sub(r'\\[\[\]()]|\$', '', tex).replace(r'\dfrac', r'\frac')
    return re.sub(r'[{}\s]', '', tex)


assert len(expected) == len(actual), (len(expected), len(actual))
for index, (want, got) in enumerate(zip(expected, actual), 1):
    assert normalized(want['tex']) == normalized(got['tex']), (index, want, got)
    assert want['display'] != got['inline'], (index, want, got)
before, after = read(source), read(output)
objects = after.xpath('/w:document/w:body/w:p/w:r[w:object]', namespaces=ns)
assert len(objects) == len(expected)
for index, (want, obj) in enumerate(zip(expected, objects), 1):
    offset = int(obj.xpath('string(w:rPr/w:position/@w:val)', namespaces=ns))
    assert offset < 0, (index, offset)
assert not after.xpath('/w:document/w:body/w:p//m:oMath', namespaces=ns)
assert [structure(t) for t in before.xpath('//w:tbl', namespaces=ns)] == [structure(t) for t in after.xpath('//w:tbl', namespaces=ns)]
numbers = native_visible_numbers(after)
assert numbers == [f'({i})' for i in range(1, sum(x['display'] for x in expected) + 1)], numbers
assert after.xpath('//w:r[w:t="前文1 "]/w:rPr/w:b', namespaces=ns)
text = ''.join(after.xpath('/w:document/w:body/w:p/w:r/w:t/text()', namespaces=ns))
assert '混排  与  末尾  结束' in text and '行外之前 ' in text and ' 行外之后' in text
cleaned = 0
for index, (want, obj) in enumerate(zip(expected, objects), 1):
    if 'manual_number' in want:
        # 旧号和原有 Tab 已清理，公式段落只保留本轮生成的两个定位 Tab。
        assert want['manual_number'] not in text, (index, want)
        paragraph = obj.getparent()
        assert not paragraph.xpath('w:r/w:t', namespaces=ns), (index, want)
        assert len(paragraph.xpath('w:r/w:tab', namespaces=ns)) == 2, (index, want)
        cleaned += 1
    if 'preserved_text' in want:
        assert want['preserved_text'] in text, (index, want, text)
ids = after.xpath('//o:OLEObject/@ObjectID', namespaces=ns)
assert len(ids) == len(set(ids)), '对象标识重复'
print(json.dumps({'Status': 'PASSED', 'ActualNativeFormulas': len(actual), 'Numbers': len(numbers), 'ManualNumbersCleaned': cleaned, 'TablePreserved': True, 'InlineBaselines': True, 'SurroundingTextAndFormatting': True}))
