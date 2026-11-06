"""生成极速四类公式的定向小样或千式输入，不启动 Word。"""
import argparse
import copy
import json
from pathlib import Path
import zipfile
from lxml import etree as ET

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS = {'w': W, 'm': M}


def run(text, bold=False):
    element = ET.Element(f'{{{W}}}r')
    if bold:
        ET.SubElement(ET.SubElement(element, f'{{{W}}}rPr'), f'{{{W}}}b')
    ET.SubElement(element, f'{{{W}}}t', {'{http://www.w3.org/XML/1998/namespace}space': 'preserve'}).text = text
    return element


def tab():
    # 真正的 Word Tab 节点与文本中的制表字符分别覆盖。
    element = ET.Element(f'{{{W}}}r')
    ET.SubElement(element, f'{{{W}}}tab')
    return element


def math(index):
    # 每式内容唯一，并交替使用下标和分式来覆盖基线深度。
    text = f'x_{{{index}}}+1'
    node = ET.Element(f'{{{M}}}oMath', nsmap={'m': M})
    script = ET.SubElement(node, f'{{{M}}}sSub')
    for name, value in [('e', 'x'), ('sub', str(index))]:
        ET.SubElement(ET.SubElement(ET.SubElement(script, f'{{{M}}}{name}'), f'{{{M}}}r'), f'{{{M}}}t').text = value
    ET.SubElement(ET.SubElement(node, f'{{{M}}}r'), f'{{{M}}}t').text = '+1'
    if index % 3 == 0:
        inner = node
        node = ET.Element(f'{{{M}}}oMath', nsmap={'m': M})
        fraction = ET.SubElement(node, f'{{{M}}}f')
        numerator = ET.SubElement(fraction, f'{{{M}}}num')
        for child in list(inner):
            numerator.append(child)
        ET.SubElement(ET.SubElement(ET.SubElement(fraction, f'{{{M}}}den'), f'{{{M}}}r'), f'{{{M}}}t').text = '2'
        text = r'\frac{' + text + '}{2}'
    return node, text


def create(base, output, count, edge_cases):
    with zipfile.ZipFile(base) as archive:
        parts = {item.filename: archive.read(item) for item in archive.infolist()}
    document = ET.fromstring(parts['word/document.xml'])
    body = document.find('w:body', NS)
    section = copy.deepcopy(body.find('w:sectPr', NS))
    body.clear()
    expected = []
    for index in range(1, count + 1):
        kind = (index - 1) % 4
        # 小样末尾复用前面的两式，验证缓存后的对象仍相互独立。
        source_index = index - 12 if edge_cases and index > count - 2 else index
        formula, tex = math(source_index)
        display = kind >= 2
        paragraph = ET.SubElement(body, f'{{{W}}}p')
        props = ET.SubElement(paragraph, f'{{{W}}}pPr')
        ET.SubElement(props, f'{{{W}}}spacing', {f'{{{W}}}after': '120'})
        if not display:
            paragraph.append(run(f'前文{index} ', True))
        if kind in (0, 2):
            start, end = (r'\[', r'\]') if display else ('$', '$')
            split = len(tex) // 2
            paragraph.extend([run(start + tex[:split], True), run(tex[split:] + end)])
        else:
            if display:
                group = ET.SubElement(paragraph, f'{{{M}}}oMathPara')
                group.append(formula)
            else:
                paragraph.append(formula)
        if not display:
            paragraph.append(run(f' 后文{index}'))
        expected.append({'tex': tex, 'display': display, 'source': 'LaTeX' if kind in (0, 2) else 'OMML'})
    if edge_cases:
        # 共段的两种来源、跨 run 定界符、正文中的行外式和共享 oMathPara。
        p = ET.SubElement(body, f'{{{W}}}p')
        p.extend([run('混排 $x_{21}'), run('+1$ 与 '), math(22)[0], run(r' 末尾 \(x_{23}+1\) 结束')])
        expected.extend({'tex': f'x_{i}+1', 'display': False, 'source': 'OMML' if i == 22 else 'LaTeX'} for i in (21, 22, 23))
        p = ET.SubElement(body, f'{{{W}}}p')
        p.append(run(r'行外之前 $$x_{25}+1$$ 行外之后', True))
        expected.append({'tex': 'x_25+1', 'display': True, 'source': 'LaTeX'})
        p = ET.SubElement(body, f'{{{W}}}p')
        group = ET.SubElement(p, f'{{{M}}}oMathPara')
        for index in (26, 28):
            formula, tex = math(index)
            group.append(formula)
            expected.append({'tex': tex, 'display': True, 'source': 'OMML'})
        # 六种单式来源均清理尾号，覆盖同一文本节点、跨 run 和真实 Tab。
        for index, start, end, number in [
                (31, '$', '$', '(801)'), (32, r'\(', r'\)', '（802）'),
                (33, '$$', '$$', '(8-03)'), (34, r'\[', r'\]', '(8.04)')]:
            p = ET.SubElement(body, f'{{{W}}}p')
            tex = math(index)[1]
            if index == 31:
                p.extend([run('\u3000'), run(start + tex[:3]), run(tex[3:] + end + '  ' + number + '  ', True)])
            else:
                # 定界符和旧号分别跨 run；首尾 Tab 不应留下空段落。
                p.extend([tab(), run(start[:1]), run(start[1:] + tex + end[:1]), run(end[1:]),
                          tab(), run(number[:2]), run(number[2:] + '\u00a0')])
            expected.append({'tex': tex, 'display': True, 'source': 'LaTeX', 'manual_number': number})
        for index, display, number in [(35, False, '(8－05)'), (36, True, '（806）')]:
            p = ET.SubElement(body, f'{{{W}}}p')
            p.extend([run(' '), tab()])
            formula, tex = math(index)
            if display:
                ET.SubElement(p, f'{{{M}}}oMathPara').append(formula)
            else:
                p.append(formula)
            p.extend([tab(), run(number[:2], True), run(number[2:] + ' ')])
            expected.append({'tex': tex, 'display': True, 'source': 'OMML', 'manual_number': number})
        # 正文里的数字、单位说明及多式段落不能作为整段旧编号清理。
        p = ET.SubElement(body, f'{{{W}}}p')
        p.append(run('正文引用 $x_{41}+1$ (901)'))
        expected.append({'tex': 'x_41+1', 'display': False, 'source': 'LaTeX', 'preserved_text': '正文引用  (901)'})
        p = ET.SubElement(body, f'{{{W}}}p')
        formula, tex = math(42)
        p.extend([run('正文原生引用 '), formula, run(' (902)')])
        expected.append({'tex': tex, 'display': False, 'source': 'OMML', 'preserved_text': '正文原生引用  (902)'})
        p = ET.SubElement(body, f'{{{W}}}p')
        p.append(run('$x_{43}+1$ (m)'))
        expected.append({'tex': 'x_43+1', 'display': False, 'source': 'LaTeX', 'preserved_text': ' (m)'})
        p = ET.SubElement(body, f'{{{W}}}p')
        p.append(run(r'\[x_{44}+1\] (说明)'))
        expected.append({'tex': 'x_44+1', 'display': True, 'source': 'LaTeX', 'preserved_text': ' (说明)'})
        p = ET.SubElement(body, f'{{{W}}}p')
        p.append(run('$x_{46}+1$ 与 $x_{47}+1$ (903)'))
        expected.extend({'tex': f'x_{i}+1', 'display': False, 'source': 'LaTeX', 'preserved_text': ' 与  (903)'} for i in (46, 47))
        p = ET.SubElement(body, f'{{{W}}}p')
        group = ET.SubElement(p, f'{{{M}}}oMathPara')
        for index in (48, 49):
            formula, tex = math(index)
            group.append(formula)
            expected.append({'tex': tex, 'display': True, 'source': 'OMML', 'preserved_text': ' (904)'})
        p.append(run(' (904)'))
        # 表格公式和文本属于排除范围，必须原样保留。
        table = ET.SubElement(body, f'{{{W}}}tbl')
        ET.SubElement(table, f'{{{W}}}tblPr')
        ET.SubElement(ET.SubElement(table, f'{{{W}}}tblGrid'), f'{{{W}}}gridCol', {f'{{{W}}}w': '8000'})
        p = ET.SubElement(ET.SubElement(ET.SubElement(table, f'{{{W}}}tr'), f'{{{W}}}tc'), f'{{{W}}}p')
        p.extend([run('表格保留 $z+1$ '), math(99)[0]])
    body.append(section)
    parts['word/document.xml'] = ET.tostring(document, xml_declaration=True, encoding='UTF-8', standalone=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, 'x', zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    output.with_suffix('.expected.json').write_text(json.dumps(expected, ensure_ascii=False, indent=2), encoding='utf-8')
    print(json.dumps({'input': str(output), 'targets': len(expected), 'numbers': sum(x['display'] for x in expected)}, ensure_ascii=False))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('base', type=Path)
    parser.add_argument('output', type=Path)
    parser.add_argument('--count', type=int, default=16)
    parser.add_argument('--edge-cases', action='store_true')
    args = parser.parse_args()
    create(args.base, args.output, args.count, args.edge_cases)
