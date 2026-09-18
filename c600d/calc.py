# -*- coding: utf-8 -*-
"""计算核心：龄期与累计温度。

口径
----
龄期(d)
    自起算日（浇注日计为第 1 天）起的自然日差 + 1。
    例：起算日 2025-11-11，则 2026-04-06 的龄期为 (2026-04-06 - 2025-11-11).days + 1 = 147。

累计温度(℃·d)
    cum(n) = cum(n-1) + MAX(T(n), 0)
    即 T > 0 时累加当日平均温度，T <= 0 时按 0 计（累计值冻结、只增不减）。
    该口径下累计值单调不减，且不会因负温日倒退。

写入方式
--------
每个数据序列只把**第一格**写成数值，其余全部写成公式：
    同表相邻行   =J19+MAX(F20,0)
    跨表相邻行   =Sheet1!J27+MAX(T6,0)
    （龄期同为 =M6+1 / =Sheet1!AA27+1）
这样一来跨数据块、跨工作表都是**一条链**，不会出现分块重复起算。
"""

MAX_FN = 'MAX({ref},0)'


def compute_age(day, base_date):
    """龄期 = 与起算日的自然日差 + 1（浇注当天算第 1 天）。"""
    return (day - base_date).days + 1


def cum_rule(temp):
    """累计温度的日增量。"""
    return temp if temp and temp > 0 else 0.0


def compute_series(cells, base_date):
    """按顺序算出生效值。

    返回 (results, cross_info)：
      results -- [(cell, age:int, cum:float), ...]
      cross_info -- {'first_cross_date','first_cross_age','first_cross_cum'} 或 None
    """
    results = []
    run = 0.0
    cross = None
    for c in cells:
        run = round(run + cum_rule(c.temp), 1)
        age = compute_age(c.date, base_date)
        results.append((c, age, run))
        if cross is None and run >= 600:
            cross = {
                'date': c.date,
                'age': age,
                'cum': run,
                'ref': c.cum_ref,
            }
    return results, cross


def age_formula(prev_cell, cur_cell):
    """生成龄期公式（引用上一行的龄期格）。"""
    return _ref_formula(prev_cell, cur_cell, prev_cell.age_col, '1')


def first_age_formula(first_cell, base_date):
    """第一个数据格的龄期公式：起算日龄期恒为 1。"""
    return '=1'


def cum_formula(prev_cell, cur_cell):
    """生成累计温度公式（引用上一行的累计格 + 本行温度格取 MAX）。"""
    return _ref_formula(prev_cell, cur_cell, prev_cell.cum_col,
                        MAX_FN.format(ref=_local_ref(cur_cell, cur_cell.temp_col)))


def _ref_formula(prev_cell, cur_cell, prev_col, suffix):
    """构造 `=上一格 + suffix`；同表省略表名，跨表带上表名。"""
    if prev_cell.sheet == cur_cell.sheet:
        prev_ref = '%s%d' % (prev_col, prev_cell.row)
    else:
        prev_ref = '%s!%s%d' % (prev_cell.sheet, prev_col, prev_cell.row)
    return '=%s+%s' % (prev_ref, suffix)


def _local_ref(cell, col):
    """本表内的引用（不带表名），用于 MAX() 里的温度格。"""
    return '%s%d' % (col, cell.row)


def first_value_formula(first_cell, base_date):
    """第一个数据格的累计温度公式：=MAX(本行平均温度,0)。

    写成公式而不是纯数值，是为了让整列每一格都是公式，口径自解释、
    插入行时也不会被漏掉。
    """
    return '=%s' % MAX_FN.format(ref=_local_ref(first_cell, first_cell.temp_col))
