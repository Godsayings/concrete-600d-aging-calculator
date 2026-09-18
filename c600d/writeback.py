# -*- coding: utf-8 -*-
"""回写：把算好的龄期与累计温度写进台账。

只改三个地方：
  1. 平均温度列（仅当提供了外部温度补充时）
  2. 累计温度列：首格写数值，其余写公式
  3. 龄期列：首格写数值，其余写公式

合计温度/龄期两列的公式构成**一条跨块、跨表的连续链**。
"""
import openpyxl

from . import calc, xmlfix


def prepare_values(sheet_positions, results):
    """算出需要注入 XML 的缓存值。"""
    values = {}
    for cell, age, cum in results:
        values[(cell.sheet, '%s%d' % (cell.cum_col, cell.row))] = cum
        values[(cell.sheet, '%s%d' % (cell.age_col, cell.row))] = age
    return values


def write_back(path, cells, layout, base_date, results, temp_changes=(),
               backup=None):
    """执行回写。

    path        -- 目标 xlsx
    cells       -- 已按日期排序的 Cell 列表
    results     -- calc.compute_series 的结果
    temp_changes-- [(日期, 原值, 新值), ...]，需写回的平均温度
    backup      -- 备份文件路径

    返回统计信息 dict。
    """
    wb = openpyxl.load_workbook(path, data_only=False)

    # ---- 1. 平均温度（外部来源）
    by_date = {c.date: c for c in cells}
    for d, _old, new in temp_changes:
        c = by_date[d]
        wb[c.sheet]['%s%d' % (c.temp_col, c.row)].value = _num(new)

    # ---- 2. 累计温度 + 3. 龄期：一条链
    stats = {'cum_formulas': 0, 'age_formulas': 0, 'cum_first': 0, 'age_first': 0}
    prev = None
    for cell, age, cum in results:
        ws = wb[cell.sheet]
        if prev is None:
            ws['%s%d' % (cell.cum_col, cell.row)].value = calc.first_value_formula(
                cell, base_date)
            ws['%s%d' % (cell.age_col, cell.row)].value = calc.first_age_formula(
                cell, base_date)
            stats['cum_first'] += 1
            stats['age_first'] += 1
        else:
            ws['%s%d' % (cell.cum_col, cell.row)].value = calc.cum_formula(prev, cell)
            ws['%s%d' % (cell.age_col, cell.row)].value = calc.age_formula(prev, cell)
            stats['cum_formulas'] += 1
            stats['age_formulas'] += 1
        prev = cell

    wb.save(path)
    wb.close()          # 必须关闭，否则 Windows 上无法替换文件

    # ---- 4. openpyxl 会清空公式缓存值，这里补回去
    positions = xmlfix.sheet_positions(path)
    values = prepare_values(positions, results)
    stats['cached'] = xmlfix.patch_cached_values(path, positions, values, backup=backup)
    return stats


def _num(v):
    """整数不显示小数位，与台账既有写法一致。"""
    return int(v) if abs(v - round(v)) < 1e-9 else round(v, 1)
