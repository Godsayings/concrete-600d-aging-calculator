# -*- coding: utf-8 -*-
"""从「试验台账汇总」生成「等效龄期计算表」（表 C5-19-2）工作表。

流程
----
1. 在台账汇总里按关键字（默认 ``600``）**自动发现**同条件试验台账工作表，
   每个工作表在目标工作簿里建一张同名的工作表。
2. 每张新表按目标工作簿模板表（默认 Sheet1 = 表 C5-19-2 等效龄期计算表）
   的格式复制：标题、表头、列宽、行高、合并区、打印设置一并继承。
3. 试件数据取自台账汇总；温度值到「逐日温度台账」里按日期取值。

计算规则
--------
============  ==========================================================
D 制作日期     ← 台账汇总的「成型日期」
E 检验日期     ← 台账汇总的「送检日期」
F 制作日期累计温度值  ← 逐日温度台账里**制作日当天**的平均温度（日均温列）
G 检验日期累计温度值  ← 逐日温度台账里**检验日当天**的平均温度（日均温列）
H 600℃·d计算值 = 累计温度(检验日) − 累计温度(制作日)（累计温度列）
I 等效龄期（d） = (送检日期 − 成型日期).days（自然天数）
============  ==========================================================

台账里查不到的日期，单元格留空并标黄，便于人工补齐。

> ⚠️ **关键：F/G 与 H 取自台账的**不同列**
>
> - **F/G** 是「当天平均温度」→ 取**日均温列**（默认左块 `F`、右块 `T`），
>   量级为 −20~35 ℃。
> - **H** 是「两个日期的累计温度之差」→ 取**累计温度列**（默认左块 `J`、右块 `X`），
>   量级为几十到几千 ℃·d。
>
> 两者列由配置 `ledger.temp_columns` 与 `ledger.cum_columns` 分别指定，不要混用。
> 表格模板的 F/G 表头文字写的是「制作日期累计温度值 / 检验日期累计温度值」，
> 与 F/G 实际填的内容**不一致**（表头是既有模板自带的，不能改）——
> **以本说明与配置为准，不要只看表头。**
"""
import datetime
import io
import os
import re
import shutil

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

from . import rawxlsx

SER = datetime.date(1899, 12, 30)
CN_DATE = re.compile(r'(\d{4})年(\d{1,2})月(\d{1,2})日')

DEFAULT_CONFIG = {
    # ---- 台账汇总 ----
    'summary': {
        'sheet_pattern': '600',        # 表名包含该关键字即为 600° 同条件试验台账
        'exclude_pattern': None,       # 需要排除的表名正则
        'header_row': None,            # None = 自动探测第 1 行
        'header_keywords': ['试件编号', '浇筑部位'],
        # 语义 -> 表头候选文字（按顺序取第一个命中的列）
        'columns': {
            'no': ['试件编号'],
            'part': ['浇筑部位'],
            'made': ['成型日期', '制作日期'],
            'sent': ['送检日期', '送检时间'],
        },
        # 缺失即为致命错误的列；其余列缺失只告警
        'required_columns': ['no'],
        'first_data_row': 2,
    },
    # ---- 逐日温度台账（如 11.12.xlsx）----
    'ledger': {
        'temp_columns': {              # 数据块 -> 该日「平均温度」列（供 F/G）
            'left': 'F', 'right': 'T',
        },
        'cum_columns': {               # 数据块 -> 该日「累计温度」列（供 H）
            'left': 'J', 'right': 'X',
        },
        'date_columns': {              # 数据块 -> 日期列
            'left': 'B', 'right': 'P',
        },
        'first_data_row': 6,
        'last_data_row': 27,
    },
    # ---- 目标工作簿 ----
    'target': {
        'template_sheet': 'Sheet1',    # 作为格式模板的表名或序号
        'keep_sheets': ['Sheet1'],     # 保留不动的表
        'title': '600°C·d实体检验等效龄期计算表\n表C5-19-2',
        'data_number': '02-01-C5-001',
        'date_format': 'yyyy/m/d;@',
        'blank_fill': 'FFFF00',        # 查不到数据的单元格底色
    },
}


def deep_merge(base, over):
    out = dict(base)
    for k, v in (over or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config(path):
    import json
    cfg = dict(DEFAULT_CONFIG)
    if path:
        with open(path, 'r', encoding='utf-8') as fh:
            cfg = deep_merge(cfg, json.load(fh))
    return cfg


def ser2date(v):
    try:
        return SER + datetime.timedelta(days=int(float(v)))
    except (TypeError, ValueError):
        return None


def parse_date_any(v):
    """支持 Excel 序列号、'2026年4月6日'、'2026-04-06' 等。"""
    if v is None:
        return None
    d = ser2date(v)
    if d and 1900 < d.year < 2200:
        return d
    m = CN_DATE.search(str(v))
    if m:
        try:
            return datetime.date(*(int(x) for x in m.groups()))
        except ValueError:
            return None
    m = re.match(r'^\s*(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})\s*$', str(v))
    if m:
        try:
            return datetime.date(*(int(x) for x in m.groups()))
        except ValueError:
            return None
    return None


# ------------------------------------------------------------------ 读取
def discover_sheets(summary_path, cfg):
    """在台账汇总里发现所有匹配的 600° 同条件试验台账工作表名。"""
    sc = cfg['summary']
    pat = re.compile(sc['sheet_pattern'])
    exc = re.compile(sc['exclude_pattern']) if sc.get('exclude_pattern') else None
    found = []
    for _i, name, _t, state in rawxlsx.sheet_list(summary_path):
        if state != 'visible':
            continue
        if not pat.search(name):
            continue
        if exc and exc.search(name):
            continue
        found.append(name)
    return found


def read_ledger_series(ledger_path, cfg):
    """读逐日温度台账，返回 ``(日均温, 累计温度)`` 两个 {日期: 值} 字典。

    F/G 填的是**当天平均温度**（取日均温列，默认左块 F、右块 T）；
    H 填的是**两个日期的累计温度之差**（取累计温度列，默认左块 J、右块 X）。
    两者取自台账的不同列，所以这里要同时读出来。

    同一数据块内同时有日均温和累计温度，因此用 (侧, 行) 连读两列即可。
    """
    lc = cfg['ledger']
    book = rawxlsx.read_book(ledger_path)
    avg, cum = {}, {}
    for side in lc['temp_columns']:
        tcol = rawxlsx.col_index(lc['temp_columns'][side])
        ccol = rawxlsx.col_index(lc['cum_columns'][side])
        dcol = rawxlsx.col_index(lc['date_columns'][side])
        for _name, cells in book.items():
            for r in range(lc['first_data_row'], lc['last_data_row'] + 1):
                raw_date = cells.get((r, dcol))
                if not raw_date:
                    continue
                d = parse_date_any(raw_date)
                if d is None:
                    continue
                tv = cells.get((r, tcol))
                if tv is not None:
                    try:
                        avg[d] = float(tv)
                    except ValueError:
                        pass
                cv = cells.get((r, ccol))
                if cv is not None:
                    try:
                        cum[d] = float(cv)
                    except ValueError:
                        pass
    return avg, cum


def read_specimens(summary_path, sheet_names, cfg):
    """读台账汇总里每个 600° 表的试件列表。

    返回 {表名: [{'no','part','made','sent'}, ...]}，顺序保留台账行序。
    """
    sc = cfg['summary']
    candidates = {k: (list(v) if isinstance(v, (list, tuple)) else [v])
                  for k, v in sc['columns'].items()}      # 语义 -> 候选表头
    required = set(sc.get('required_columns') or ['no'])
    book = rawxlsx.read_book(summary_path, names=set(sheet_names))
    result = {}
    fatal = []
    for name in sheet_names:
        cells = book.get(name, {})
        if not cells or '__error__' in cells:
            result[name] = []
            fatal.append((name, ['读取失败']))
            continue
        # 定位表头行并映射列号
        hrow = sc.get('header_row')
        if hrow is None:
            hrow, hdr_cells = rawxlsx.find_header(cells, sc['header_keywords'])
        else:
            hdr_cells = {c: (cells.get((hrow, c)) or '').strip()
                         for c in range(1, 61) if cells.get((hrow, c))}
        colmap = {}
        if hdr_cells:
            # 先按候选词精确匹配，再退化为包含匹配
            for key, wants in candidates.items():
                for want in wants:
                    for c, text in hdr_cells.items():
                        if str(text).strip() == want:
                            colmap[key] = c
                            break
                    if key in colmap:
                        break
            for key, wants in candidates.items():
                if key in colmap:
                    continue
                for want in wants:
                    for c, text in hdr_cells.items():
                        if want in str(text):
                            colmap[key] = c
                            break
                    if key in colmap:
                        break

        missing = [k for k in required if k not in colmap]
        if missing:
            fatal.append((name, ['、'.join(candidates[k]) for k in missing]))
            result[name] = []
            continue

        recs = []
        rows = sorted({r for (r, _c) in cells if r > hrow})
        for r in rows:
            no = cells.get((r, colmap.get('no', 0)))
            if not no or not str(no).strip():
                continue
            recs.append({
                'no': str(no).strip(),
                'part': (cells.get((r, colmap.get('part', 0))) or '').strip(),
                'made': parse_date_any(cells.get((r, colmap.get('made', 0)))),
                'sent': parse_date_any(cells.get((r, colmap.get('sent', 0)))),
            })
        result[name] = recs
    if fatal:
        raise ValueError(
            '以下工作表的必需列没找到（可用 summary.columns 覆盖列名、'
            'summary.header_row 指定表头行）：\n' +
            '\n'.join('  %s -> 缺 %s' % (n, '、'.join(m)) for n, m in fatal))
    return result


# ------------------------------------------------------------------ 建表
def build(target_path, summary_path, ledger_path, cfg, write=False, backup=None):
    """生成工作表。write=False 时只返回统计信息，不改文件。"""
    tc = cfg['target']
    names = discover_sheets(summary_path, cfg)
    if not names:
        raise SystemExit('台账汇总里没有找到匹配的 600° 同条件试验台账工作表')

    avg, cum = read_ledger_series(ledger_path, cfg)
    specs = read_specimens(summary_path, names, cfg)

    wb = openpyxl.load_workbook(target_path)
    tpl = wb[tc['template_sheet']] if isinstance(tc['template_sheet'], str) \
        else wb.worksheets[tc['template_sheet']]
    fill = PatternFill('solid', fgColor=tc['blank_fill'])
    date_fmt = tc['date_format']

    stats = []
    for name in names:
        if name in wb.sheetnames:
            del wb[name]
        ws = wb.copy_worksheet(tpl)
        ws.title = name

        # 清空模板数据区
        for r in range(3, ws.max_row + 1):
            for c in range(1, 10):
                cell = ws.cell(row=r, column=c)
                cell.value = None
                cell.fill = PatternFill(fill_type=None)
        ws['A1'] = tc['title']
        ws['H1'] = tc['data_number']

        n_full = n_blank = n_empty = 0
        r = 3
        for k, rec in enumerate(specs[name]):
            made, sent = rec['made'], rec['sent']
            if made is None and sent is None and not rec['part']:
                # 台账里只有编号的空占位行：只写编号
                ws.cell(row=r, column=1).value = k + 1
                ws.cell(row=r, column=3).value = rec['no']
                n_empty += 1
                r += 1
                continue

            tF = avg.get(made) if made else None
            tG = avg.get(sent) if sent else None
            cF = cum.get(made) if made else None
            cG = cum.get(sent) if sent else None
            ws.cell(row=r, column=1).value = k + 1
            ws.cell(row=r, column=2).value = rec['part']
            ws.cell(row=r, column=3).value = rec['no']
            if made:
                cell = ws.cell(row=r, column=4)
                cell.value = datetime.datetime(made.year, made.month, made.day)
                cell.number_format = date_fmt
            if sent:
                cell = ws.cell(row=r, column=5)
                cell.value = datetime.datetime(sent.year, sent.month, sent.day)
                cell.number_format = date_fmt

            # F/G 取日均温；H 取两个日期的累计温度之差。缺哪个标黄哪个。
            blanks = []
            for col, val in ((4, made), (5, sent)):
                if val is None:
                    blanks.append(col)
            for col, val in ((6, tF), (7, tG)):
                if val is None:
                    blanks.append(col)
                else:
                    ws.cell(row=r, column=col).value = val
            if cF is not None and cG is not None:
                ws.cell(row=r, column=8).value = round(cG - cF, 1)
            else:
                blanks.append(8)
            if made and sent:
                ws.cell(row=r, column=9).value = (sent - made).days
            else:
                blanks.append(9)
            for col in blanks:
                ws.cell(row=r, column=col).fill = fill
            if blanks:
                n_blank += 1
            else:
                n_full += 1
            r += 1
        ws.cell(row=r + 1, column=1).value = '制表人：'
        stats.append({'sheet': name, 'specimens': len(specs[name]),
                      'complete': n_full, 'incomplete': n_blank,
                      'placeholder': n_empty})

    if write:
        if backup:
            shutil.copyfile(target_path, backup)
        wb.save(target_path)
    return stats
