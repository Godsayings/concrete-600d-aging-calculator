# -*- coding: utf-8 -*-
"""命令行入口。

    python -m c600d.cli --xlsx 台账.xlsx --base 2025-11-11 --dry-run
    python -m c600d.cli --xlsx 台账.xlsx --base 2025-11-11 --write

默认只做核算并打印结果（dry-run），加 --write 才修改文件。
"""
import argparse
import datetime
import sys

from . import __version__, calc, config, ledger, writeback


def build_parser():
    p = argparse.ArgumentParser(
        prog='c600d',
        description='600℃·d 结构实体检验温度记录：龄期与累计温度自动计算',
    )
    p.add_argument('--xlsx', required=True, help='台账 .xlsx 路径')
    p.add_argument('--base', required=True,
                   help='起算日（浇注日，计为第 1 天），如 2025-11-11')
    p.add_argument('--profile', help='可选 JSON 配置（表头布局等）')
    p.add_argument('--temps', help='可选 CSV：日期,平均温度 —— 补齐或替换指定日期的温度')
    p.add_argument('--sheets', help='可选：参与计算的工作表，逗号分隔')
    p.add_argument('--write', action='store_true',
                   help='写回台账（不加则只核算打印）')
    p.add_argument('--backup', help='写回前的备份路径')
    p.add_argument('--quiet', action='store_true', help='只打印汇总')
    p.add_argument('--version', action='version', version='c600d %s' % __version__)
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)

    base_date = config.parse_date(args.base)
    if base_date is None:
        print('起算日无法解析：%r' % args.base, file=sys.stderr)
        return 2

    profile = config.load_profile(args.profile)
    layout = config.build_layout(profile)
    order = [s.strip() for s in args.sheets.split(',')] if args.sheets else None

    cells = ledger.read_ledger(args.xlsx, layout, sheet_order=order)
    problems = ledger.validate(cells)
    if problems:
        print('台账校验发现问题：', file=sys.stderr)
        for p in problems:
            print('  - %s' % p, file=sys.stderr)
        return 3

    changed = ledger.apply_temp_overrides(cells, config.load_temp_overrides(args.temps))
    results, cross = calc.compute_series(cells, base_date)

    if not args.quiet:
        print('%-12s %-8s %-12s %-10s %-8s %s' % (
            '日期', '平均温度', '累计温度', '龄期(d)', '工作表', '单元格'))
        for cell, age, cum in results:
            print('%-12s %-8s %-12s %-10s %-8s %s' % (
                cell.date, cell.temp, cum, age, cell.sheet, cell.cum_ref))

    print()
    print('日记录数：%d' % len(cells))
    print('日期范围：%s ~ %s' % (cells[0].date, cells[-1].date))
    print('起算日  ：%s（该日龄期 = 1）' % base_date)
    print('累计温度：%s ~ %s ℃·d' % (results[0][2], results[-1][2]))
    print('龄期    ：%s ~ %s d' % (results[0][1], results[-1][1]))
    if changed:
        print('外部温度替换 %d 天：' % len(changed))
        for d, old, new in changed:
            print('  %s  %s -> %s' % (d, old, new))
    if cross:
        print('600℃·d 达标日：%s（龄期 %d，当日累计 %.1f）' % (
            cross['date'], cross['age'], cross['cum']))

    if not args.write:
        print('\n（未写入。加 --write 才会修改台账。）')
        return 0

    stats = writeback.write_back(
        args.xlsx, cells, layout, base_date, results,
        temp_changes=changed, backup=args.backup)
    print('\n已写回：%s' % args.xlsx)
    if args.backup:
        print('备份  ：%s' % args.backup)
    print('累计温度公式 %d 个 + 初值 %d 个；龄期公式 %d 个 + 初值 %d 个；注入缓存值 %d 个' % (
        stats['cum_formulas'], stats['cum_first'],
        stats['age_formulas'], stats['age_first'], stats['cached']))
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
