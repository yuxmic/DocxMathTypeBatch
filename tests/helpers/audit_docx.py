"""开发验收：用标准库独立检查 DOCX 容器、XML、关系和表格拓扑。"""
import hashlib
import json
import posixpath
import sys
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def audit(path):
    with zipfile.ZipFile(path) as package:
        names = package.namelist()
        assert len(names) == len(set(names)), "重复 ZIP 条目"
        assert package.testzip() is None, "ZIP CRC 错误"
        parts = {name: package.read(name) for name in names}
    # 全部 XML 都解析；不执行宏，不访问外部关系目标。
    xml = {name: ET.fromstring(data) for name, data in parts.items()
           if name.endswith((".xml", ".rels"))}
    types = xml["[Content_Types].xml"]
    defaults = {node.attrib["Extension"].lower(): node.attrib["ContentType"]
                for node in types if node.tag.endswith("}Default")}
    overrides = {urllib.parse.unquote(node.attrib["PartName"]).lstrip("/"):
                 node.attrib["ContentType"] for node in types if node.tag.endswith("}Override")}
    for name in parts:
        if name != "[Content_Types].xml":
            assert name in overrides or name.rsplit(".", 1)[-1].lower() in defaults, f"缺内容类型：{name}"
    relations = {}
    for name, root in xml.items():
        if not name.endswith(".rels"):
            continue
        source = "" if name == "_rels/.rels" else posixpath.join(
            posixpath.dirname(posixpath.dirname(name)), posixpath.basename(name)[:-5])
        assert not source or source in parts, f"关系源缺失：{source}"
        entries = {}
        for node in root:
            rid = node.attrib["Id"]
            assert rid not in entries, f"重复关系 ID：{name}/{rid}"
            target = None
            if node.attrib.get("TargetMode") != "External":
                uri = urllib.parse.urlsplit(node.attrib["Target"])
                assert not uri.scheme and not uri.netloc, "内部关系包含外部地址"
                target = posixpath.normpath(posixpath.join(posixpath.dirname(source),
                                                        urllib.parse.unquote(uri.path))).lstrip("/")
                assert target in parts, f"内部关系目标缺失：{target}"
            entries[rid] = (target, node.attrib["Type"])
        relations[source] = entries
    for name, root in xml.items():
        if name.endswith(".rels"):
            continue
        for node in root.iter():
            for attribute, rid in node.attrib.items():
                if attribute.startswith(R):
                    assert rid in relations.get(name, {}), f"XML 关系引用缺失：{name}/{rid}"
    document = xml["word/document.xml"]
    parents = {child: parent for parent in document.iter() for child in parent}
    equations = [node for node in document.iter() if node.tag.endswith("}OLEObject")
                 and node.attrib.get("ProgID", "").startswith("Equation.DSMT")]
    for node in equations:
        target, kind = relations["word/document.xml"][node.attrib[R + "id"]]
        assert target and kind.endswith("/oleObject"), "MathType 关系类型错误"
        # 结构化存储和 Equation Native 完整回读由原生转换器检查；这里独立核对 CFB 标识及预览绑定。
        assert len(parts[target]) >= 512 and parts[target][:8] == bytes.fromhex("d0cf11e0a1b11ae1"), "OLE 容器错误"
        previews = [image for image in parents[node].iter() if image.tag.endswith("}imagedata")]
        assert previews, "MathType 缺少预览"
        for image in previews:
            preview, kind = relations["word/document.xml"][image.attrib[R + "id"]]
            assert preview and parts[preview] and kind.endswith("/image"), "MathType 预览关系错误"
    # 忽略 Word 规范化的样式/坐标，比较每层表格的行、物理单元格、网格跨度和纵向合并。
    topology = []
    for table in document.iter(W + "tbl"):
        depth, ancestor = 0, parents.get(table)
        while ancestor is not None:
            depth += ancestor.tag == W + "tbl"
            ancestor = parents.get(ancestor)
        rows = []
        for row in table.findall(W + "tr"):
            cells = []
            for cell in row.findall(W + "tc"):
                span = cell.find(W + "tcPr/" + W + "gridSpan")
                merge = cell.find(W + "tcPr/" + W + "vMerge")
                cells.append([span.attrib.get(W + "val", "1") if span is not None else "1",
                              merge.attrib.get(W + "val", "continue") if merge is not None else ""])
            rows.append(cells)
        topology.append([depth, rows])
    return {"Status": "PASSED", "Parts": len(parts), "XmlParts": len(xml),
            "MathType": len(equations), "Tables": len(topology),
            "TableTopology": topology,
            "TableTopologySHA256": hashlib.sha256(json.dumps(topology).encode()).hexdigest(),
            "Text": "".join(node.text or "" for node in document.iter(W + "t")),
            "OMML": sum(node.tag.endswith("}oMath") for node in document.iter()),
            "Checks": ["ZIP_CRC", "ALL_XML_PARSE", "CONTENT_TYPES", "INTERNAL_RELATIONSHIPS",
                       "XML_RELATIONSHIP_REFERENCES", "OLE_CFB_HEADER_AND_PREVIEW"]}


if __name__ == "__main__":
    # 被 test-acceptance.ps1 调用；断言失败即非零退出，检查不会冒充完整数学/CFB 审计。
    print(json.dumps(audit(sys.argv[1]), ensure_ascii=False))
