# -*- coding: utf-8 -*-
"""xlsx 层的小工具。

openpyxl 保存工作簿时会把**公式单元格的缓存值（<v>）清空**，导致用
WPS/Excel 之外的读取方式（以及部分预览器）看不到数字。这里在保存后直接
改写 worksheet XML，把已知的缓存值注入回去。

注意：注入缓存值后，不要再调用 openpyxl 的 save()，否则缓存值会再次被清空。
"""
import os
import re
import shutil
import time
import zipfile

_CELL_RE = re.compile(
    r'<c r="(?P<ref>[A-Z]+\d+)"(?P<style>[^>]*?)>'
    r'<f>(?P<f>[^<]*)</f>\s*<v\s*/></c>')


def patch_cached_values(path, sheet_positions, values, backup=None):
    """把 values（{(表名, 单元格): 数值}）写进对应 sheet XML 的 <v>。

    sheet_positions: {表名: 该表在 xl/worksheets/sheetN.xml 中的序号 N}
    backup: 可选备份路径。
    返回成功注入的单元格数。
    """
    if backup:
        shutil.copyfile(path, backup)

    zin = zipfile.ZipFile(path)
    items = zin.infolist()
    data = {name: zin.read(name) for name in zin.namelist()}
    zin.close()

    # 单元格 -> (表名, 数值)，按 sheet 分组
    by_sheet = {}
    for (sheet, ref), val in values.items():
        by_sheet.setdefault(sheet, {})[ref] = val

    patched = 0
    for sheet, refs in by_sheet.items():
        idx = sheet_positions.get(sheet)
        if idx is None:
            continue
        key = 'xl/worksheets/sheet%d.xml' % idx
        if key not in data:
            continue
        xml = data[key].decode('utf-8')

        state = {'n': 0}

        def repl(m):
            ref = m.group('ref')
            if ref not in refs:
                return m.group(0)
            state['n'] += 1
            return '<c r="%s"%s><f>%s</f><v>%s</v></c>' % (
                ref, m.group('style'), m.group('f'), refs[ref])

        data[key] = _CELL_RE.sub(repl, xml).encode('utf-8')
        patched += state['n']

    tmp = path + '.tmp'
    zout = zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED)
    for it in items:
        zout.writestr(it, data[it.filename])
    zout.close()
    _replace_with_retry(tmp, path)
    return patched


def _replace_with_retry(src, dst, attempts=10):
    """在 Windows 上替换文件常被短暂占用（句柄未释放 / 杀毒扫描），重试几次。"""
    last = None
    for i in range(attempts):
        try:
            os.replace(src, dst)
            return
        except PermissionError as exc:      # WinError 5 / 32
            last = exc
            time.sleep(0.15 * (i + 1))
    try:
        os.remove(src)
    except OSError:
        pass
    raise PermissionError(
        '无法替换 %s（可能仍被 Excel/WPS 占用，请先关闭该文件）：%s' % (dst, last))


def count_formulas(path, sheet_names):
    """统计各表公式单元格总数与带缓存值的数量，用于自检。"""
    result = {}
    with zipfile.ZipFile(path) as z:
        for i, name in enumerate(sheet_names, start=1):
            key = 'xl/worksheets/sheet%d.xml' % i
            try:
                xml = z.read(key).decode('utf-8')
            except KeyError:
                continue
            total = withv = 0
            for m in re.finditer(
                    r'<c r="[A-Z]+\d+"[^>]*><f>[^<]*</f>(?:<v>[^<]*</v>)?', xml):
                total += 1
                if '<v>' in m.group(0):
                    withv += 1
            result[name] = (total, withv)
    return result


def sheet_positions(path):
    """返回 {表名: sheetN 的 N}，按 workbook.xml 中的顺序。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    names = list(wb.sheetnames)
    wb.close()
    return {name: i for i, name in enumerate(names, start=1)}
