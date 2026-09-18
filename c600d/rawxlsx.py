# -*- coding: utf-8 -*-
"""直接读写 xlsx 内部 XML 的最小工具。

为什么需要它
------------
1. openpyxl 打不开某些工作簿——例如样式表里带有非标准 `<fill>` 定义时会抛
   ``TypeError: Fill() takes no arguments``。本模块只解析 sharedStrings 与
   worksheet XML，完全不碰样式表，因此不受影响。
2. **openpyxl 默认模式读「内联字符串」（``t="inlineStr"``）会返回 None。**
   对这类单元格，本模块与 ``openpyxl.load_workbook(..., read_only=True)``
   都能正确读出。排查「数值莫名丢失」时优先用本模块交叉验证。
"""
import re
import zipfile
import xml.etree.ElementTree as ET

M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
R_NS = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
R_ID = R_NS + 'id'


def col_index(letters):
    """'A' -> 1, 'AA' -> 27"""
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letter(n):
    """1 -> 'A', 27 -> 'AA'"""
    s = ''
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def sheet_list(path):
    """返回 [(序号, 表名, xl 内部路径, 状态)]，顺序与工作簿一致。"""
    with zipfile.ZipFile(path) as z:
        rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rid2 = {r.get('Id'): r.get('Target') for r in rels}
        wb = ET.fromstring(z.read('xl/workbook.xml'))
        out = []
        for i, sh in enumerate(wb.find('{%s}sheets' % M), start=1):
            tgt = rid2.get(sh.get(R_ID), '')
            if tgt.startswith('/'):
                tgt = tgt[1:]
            elif not tgt.startswith('xl/'):
                tgt = 'xl/' + tgt.lstrip('./')
            out.append((i, sh.get('name'), tgt, sh.get('state') or 'visible'))
        return out


def read_sheet(path, target):
    """读单个工作表，返回 {(行, 列): 值(str)}（行、列均从 1 开始）。"""
    with zipfile.ZipFile(path) as z:
        shared = []
        if 'xl/sharedStrings.xml' in z.namelist():
            root = ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in root.findall('{%s}si' % M):
                shared.append(''.join(t.text or '' for t in si.iter('{%s}t' % M)))
        root = ET.fromstring(z.read(target))
        cells = {}
        for c in root.iter('{%s}c' % M):
            ref = c.get('r')
            t = c.get('t')
            v = c.find('{%s}v' % M)
            if t == 's' and v is not None:
                val = shared[int(v.text)]
            elif t == 'inlineStr':
                isn = c.find('{%s}is' % M)
                val = ''.join(x.text or '' for x in isn.iter('{%s}t' % M)) if isn is not None else ''
            elif v is not None:
                val = v.text
            else:
                continue
            if val is None or str(val).strip() == '':
                continue
            m = re.match(r'([A-Z]+)(\d+)', ref)
            cells[(int(m.group(2)), col_index(m.group(1)))] = str(val)
        return cells


def read_book(path, names=None):
    """读整个工作簿；names 为要读的表名集合（None 表示全部）。"""
    out = {}
    for _i, name, target, _st in sheet_list(path):
        if names is not None and name not in names:
            continue
        try:
            out[name] = read_sheet(path, target)
        except (KeyError, ET.ParseError) as exc:
            out[name] = {'__error__': str(exc)}
    return out


def find_header(cells, must_contain, max_scan=6):
    """在前若干行里找表头行。

    返回 ``(表头行号, {列号: 列名})``；找不到时返回 ``(None, {})``。
    must_contain 里的每个词都必须出现在该行某个单元格中。
    """
    rows = sorted(set(r for r, _c in cells))
    for r in rows[:max_scan]:
        row = {c: (cells.get((r, c)) or '').strip()
               for c in range(1, 61) if cells.get((r, c))}
        if all(any(word in v for v in row.values()) for word in must_contain):
            return r, row
    return None, {}
