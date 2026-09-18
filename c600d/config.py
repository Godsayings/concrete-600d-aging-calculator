# -*- coding: utf-8 -*-
"""配置与输入解析。

工程名称、施工单位等涉密信息一律由用户通过 profile 传入，
代码中不内置任何具体工程信息。
"""
import csv
import datetime
import json
import os
import re

DATE_RE = re.compile(r'^\s*(\d{4})\s*[年\-/.]\s*(\d{1,2})\s*[月\-/.]\s*(\d{1,2})\s*日?\s*$')

#: 默认表格布局：数据块内的列号（0 基，可用列字母自行换算）
DEFAULT_LAYOUT = {
    'first_data_row': 6,      # 数据块首行
    'block_rows': 22,         # 每个数据块行数
    'left': {'date': 'B', 'temp': 'F', 'cum': 'J', 'age': 'M'},
    'right': {'date': 'P', 'temp': 'T', 'cum': 'X', 'age': 'AA'},
    'skip_sheets': [],
}


class ConfigError(ValueError):
    """配置或输入不合法。"""


def parse_date(value):
    """把单元格里的日期解析为 date。

    支持 datetime/date 对象，以及 '2026年04月06日'、'2026-04-06' 等文本。
    """
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    m = DATE_RE.match(str(value))
    if not m:
        return None
    y, mo, d = (int(x) for x in m.groups())
    try:
        return datetime.date(y, mo, d)
    except ValueError:
        return None


def parse_temp(value):
    """把单元格里的平均温度解析为 float；空值返回 None。"""
    if value is None or value == '':
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip())
    except ValueError:
        return None


def load_profile(path):
    """读取 JSON 配置。未指定路径时返回空配置。"""
    if not path:
        return {}
    if not os.path.exists(path):
        raise ConfigError('配置文件不存在：%s' % path)
    with open(path, 'r', encoding='utf-8') as fh:
        return json.load(fh)


def load_temp_overrides(path):
    """读取「日期,平均温度」两列 CSV，用于补齐或替换指定日期的温度。

    空温度单元格会被跳过（不覆盖）。气象网等外部来源的数据用这种方式注入，
    不进代码，便于按工程替换。
    """
    if not path:
        return {}
    if not os.path.exists(path):
        raise ConfigError('温度补充文件不存在：%s' % path)
    out = {}
    with open(path, 'r', encoding='utf-8-sig', newline='') as fh:
        for lineno, row in enumerate(csv.reader(fh), start=1):
            if not row or not str(row[0]).strip() or str(row[0]).lstrip().startswith('#'):
                continue
            d = parse_date(row[0])
            if d is None:
                if lineno == 1:
                    continue          # 表头
                raise ConfigError('温度补充文件第 %d 行日期无法解析：%r' % (lineno, row[0]))
            if len(row) < 2 or not str(row[1]).strip():
                continue              # 只有日期、没有温度 → 跳过
            out[d] = float(str(row[1]).strip())
    return out


def build_layout(profile):
    """合并默认布局与用户覆盖项。"""
    layout = json.loads(json.dumps(DEFAULT_LAYOUT))   # 深拷贝
    user = (profile or {}).get('layout') or {}
    for key in ('first_data_row', 'block_rows', 'skip_sheets'):
        if key in user:
            layout[key] = user[key]
    for side in ('left', 'right'):
        if side in user:
            layout[side].update(user[side])
    return layout
