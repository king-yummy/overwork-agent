# -*- coding: utf-8 -*-
"""xlsx 템플릿의 '데이터 셀만' 외과적으로 수정 — 차트·수식·서식 100% 보존.

openpyxl로 저장하면 차트 색상/스타일이 유실되므로, zip 내부의 해당 워크시트
XML만 직접 수정하고 나머지 파트(차트·드로잉·기타 시트)는 원본 그대로 둔다.
문자열은 inlineStr로 써서 sharedStrings 테이블을 건드리지 않는다.
"""
import zipfile
import shutil
from lxml import etree

M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": M, "r": R}


def col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + (ord(ch.upper()) - 64)
    return n


def sheet_xml_path(z: zipfile.ZipFile, sheet_name: str) -> str:
    wb = etree.fromstring(z.read("xl/workbook.xml"))
    rels = etree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    relmap = {r.get("Id"): r.get("Target") for r in rels}
    for s in wb.findall(".//m:sheets/m:sheet", NS):
        if s.get("name") == sheet_name:
            rid = s.get(f"{{{R}}}id")
            tgt = relmap[rid]
            return tgt if tgt.startswith("xl/") else "xl/" + tgt
    raise KeyError(f"시트 없음: {sheet_name}")


def _set_cell(row_el, col, rownum, value, kind):
    """kind: True=문자열(inlineStr), False=숫자, 'f'=수식."""
    ref = f"{col}{rownum}"
    cnum = col_to_num(col)
    # 기존 셀 찾기 / 삽입 위치
    target, insert_before = None, None
    for c in row_el.findall("m:c", NS):
        r = c.get("r")
        ccol = "".join(ch for ch in r if ch.isalpha())
        if r == ref:
            target = c; break
        if col_to_num(ccol) > cnum and insert_before is None:
            insert_before = c
    if target is None:
        target = etree.SubElement(row_el, f"{{{M}}}c")
        target.set("r", ref)
        if insert_before is not None:
            row_el.remove(target); insert_before.addprevious(target)
    # 비우기
    for ch in list(target):
        target.remove(ch)
    if kind == "f":
        target.attrib.pop("t", None)
        f = etree.SubElement(target, f"{{{M}}}f")
        f.text = str(value)[1:] if str(value).startswith("=") else str(value)
        return
    if value is None or value == "":
        target.attrib.pop("t", None)
        return
    if kind is True:
        target.set("t", "inlineStr")
        is_el = etree.SubElement(target, f"{{{M}}}is")
        t = etree.SubElement(is_el, f"{{{M}}}t")
        t.text = str(value)
        if str(value)[:1].isspace() or str(value)[-1:].isspace():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    else:
        target.attrib.pop("t", None)
        v = etree.SubElement(target, f"{{{M}}}v")
        v.text = repr(float(value)) if isinstance(value, float) else str(value)


def write_cells(sheet_bytes: bytes, cells: dict) -> bytes:
    """cells: {(col_letter, rownum): (value, kind)}  kind: True/False/'f'"""
    root = etree.fromstring(sheet_bytes)
    sd = root.find("m:sheetData", NS)
    rows = {int(r.get("r")): r for r in sd.findall("m:row", NS)}
    for (col, rownum), (value, kind) in cells.items():
        row_el = rows.get(rownum)
        if row_el is None:
            row_el = etree.SubElement(sd, f"{{{M}}}row")
            row_el.set("r", str(rownum))
            later = [rr for rn, rr in rows.items() if rn > rownum]
            if later:
                ib = min(later, key=lambda e: int(e.get("r")))
                sd.remove(row_el); ib.addprevious(row_el)
            rows[rownum] = row_el
        _set_cell(row_el, col, rownum, value, kind)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _force_recalc(wb_bytes: bytes) -> bytes:
    """엑셀이 열 때 모든 수식을 재계산하도록 calcPr fullCalcOnLoad=1 설정."""
    root = etree.fromstring(wb_bytes)
    calc = root.find("m:calcPr", NS)
    if calc is None:
        calc = etree.SubElement(root, f"{{{M}}}calcPr")
        # calcPr는 sheets 뒤 적절한 위치면 되지만 끝에 둬도 엑셀이 허용
    calc.set("fullCalcOnLoad", "1")
    calc.set("calcId", "0")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _strip_calcchain(name: str, raw: bytes) -> bytes:
    """[Content_Types].xml / workbook.xml.rels 에서 calcChain 참조 제거."""
    try:
        root = etree.fromstring(raw)
    except Exception:
        return raw
    for el in list(root):
        tag = etree.QName(el).localname
        if tag == "Override" and (el.get("PartName") or "").endswith("calcChain.xml"):
            root.remove(el)
        if tag == "Relationship" and (el.get("Target") or "").endswith("calcChain.xml"):
            root.remove(el)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def fill_template(src_path: str, dst_path: str, sheet_cells: dict):
    """sheet_cells: {sheet_name: {(col,row):(value,kind)}}  나머지 파트는 원본 보존."""
    shutil.copy(src_path, dst_path)
    with zipfile.ZipFile(src_path) as zin:
        paths = {sn: sheet_xml_path(zin, sn) for sn in sheet_cells}
        modified = {}
        for sn, cells in sheet_cells.items():
            modified[paths[sn]] = write_cells(zin.read(paths[sn]), cells)
        modified["xl/workbook.xml"] = _force_recalc(zin.read("xl/workbook.xml"))
        names = zin.namelist()
        infos = {i.filename: i for i in zin.infolist()}
        data = {n: zin.read(n) for n in names}
    # calcChain.xml 제거 (수식 셀을 새로 넣어 캐시가 어긋남 → 엑셀이 열 때 자동 재생성)
    for ct in ("[Content_Types].xml", "xl/_rels/workbook.xml.rels"):
        if ct in data:
            modified.setdefault(ct, _strip_calcchain(ct, data[ct]))
    for p, b in modified.items():
        data[p] = b
    names = [n for n in names if n != "xl/calcChain.xml"]
    with zipfile.ZipFile(dst_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zi = zipfile.ZipInfo(n, date_time=infos[n].date_time)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = infos[n].external_attr
            zout.writestr(zi, data[n])