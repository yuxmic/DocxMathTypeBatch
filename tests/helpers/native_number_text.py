"""开发核验：从嵌套域的缓存值还原 MTPlaceRef 可见编号，不执行 Word 域。"""
import re
from docx.oxml.ns import qn

def native_visible_numbers(element):
    stack, numbers = [], []
    for node in element.iter():
        if node.tag == qn('w:fldChar'):
            kind = node.get(qn('w:fldCharType'))
            if kind == 'begin':
                stack.append({'code': '', 'result': '', 'separate': False})
            elif kind == 'separate' and stack:
                stack[-1]['separate'] = True
            elif kind == 'end' and stack:
                field = stack.pop()
                text = field['result']
                if re.match(r'\s*MACROBUTTON\s+MTPlaceRef\b', field['code']):
                    text = re.sub(r'\s*\\\*\s+MERGEFORMAT\b', '', field['code'].split('MTPlaceRef', 1)[1]).strip()
                    numbers.append(text)
                if stack:
                    key = 'result' if stack[-1]['separate'] else 'code'
                    stack[-1][key] += text
        elif node.tag in (qn('w:t'), qn('w:instrText')) and stack:
            key = 'result' if stack[-1]['separate'] else 'code'
            stack[-1][key] += node.text or ''
    return numbers
