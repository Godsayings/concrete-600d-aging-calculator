# -*- coding: utf-8 -*-
"""等效龄期计算表生成命令。

    python -m c600d eqage --summary 台账汇总.xlsx --ledger 逐日温度台账.xlsx \
                          --target 大气测温统计表.xlsx --write

默认只核算打印，加 --write 才修改目标文件（并自动备份）。
"""
import argparse
import sys

from . import __version__, eqage


def build_parser():
    p = argparse.ArgumentParser(
        prog='c600d eqage',
        description='从试验台账汇总生成 600℃·d 等效龄期计算表工作表',
    )
    p.add_argument('--summary', required=True, help='试验台账汇总 .xlsx')
    p.add_argument('--ledger', required=True, help='逐日温度台账 .xlsx（提供累计温度）')
    p.add_argument('--target', required=True, help='目标工作簿 .xlsx（在其中建表）')
    p.add_argument('--profile', help='可选 JSON 配置（表名关键字、列名等）')
    p.add_argument('--write', action='store_true', help='写回目标文件')
    p.add_argument('--backup', help='写回前的备份路径')
    p.add_argument('--list', action='store_true', help='只列出发现的工作表名后退出')
    p.add_argument('--version', action='version', version='c600d %s' % __version__)
    return p


def main(argv=None):
    args = build_parser().parse_args(argv)
    cfg = eqage.load_config(args.profile)

    try:
        names = eqage.discover_sheets(args.summary, cfg)
    except FileNotFoundError as exc:
        print('文件不存在：%s' % exc, file=sys.stderr)
        return 2

    print('在台账汇总里发现 %d 个匹配 %r 的工作表：' % (
        len(names), cfg['summary']['sheet_pattern']))
    for n in names:
        print('  - %s' % n)
    if args.list:
        return 0

    stats = eqage.build(args.target, args.summary, args.ledger, cfg,
                        write=args.write, backup=args.backup)

    print()
    print('%-40s %6s %6s %8s %8s' % ('工作表', '试件', '完整', '待补数据', '空占位'))
    for s in stats:
        print('%-40s %6d %6d %8d %8d' % (
            s['sheet'], s['specimens'], s['complete'],
            s['incomplete'], s['placeholder']))

    total = sum(s['specimens'] for s in stats)
    print()
    print('合计 %d 张表、%d 个试件；完整 %d、待补数据 %d' % (
        len(stats), total, sum(s['complete'] for s in stats),
        sum(s['incomplete'] for s in stats)))
    if args.write:
        print('已写回：%s' % args.target)
        if args.backup:
            print('备份  ：%s' % args.backup)
    else:
        print('（未写入。加 --write 才会修改目标文件。）')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
