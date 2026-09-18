# -*- coding: utf-8 -*-
"""等效龄期计算表生成的自测（全部使用合成数据）。

    python -m unittest discover -s tests -t . -v
"""
import datetime
import os
import shutil
import sys
import unittest

import openpyxl

REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, REPO_ROOT)

from c600d import eqage, rawxlsx  # noqa: E402

TMP_ROOT = os.path.join(REPO_ROOT, '.tmp_tests')
BASE = datetime.date(2025, 11, 11)

def _cfg(**summary_over):
    """在默认配置上覆盖 summary 部分（保留 columns 默认值）。"""
    return eqage.deep_merge(eqage.DEFAULT_CONFIG, {'summary': summary_over})


CFG = _cfg()


def fresh_dir(name):
    path = os.path.join(TMP_ROOT, name)
    shutil.rmtree(path, ignore_errors=True)
    os.makedirs(path, exist_ok=True)
    return path


def make_summary(path):
    """台账汇总：一张 600° 表 + 一张无关表。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '2#楼600°同条件试验台账'
    ws.append(['序号', '试件编号', '浇筑部位', '强度等级', '成型日期', '送检日期', '备注'])
    ws.append([1, 'T-001', '某工程一层柱', 'C30',
               datetime.datetime(2025, 11, 12), datetime.datetime(2025, 11, 20), ''])
    ws.append([2, 'T-002', '某工程一层梁', 'C30',
               datetime.datetime(2025, 11, 13), datetime.datetime(2025, 12, 1), ''])
    ws.append(['', 'T-003', '', '', None, None, ''])          # 空占位行
    ws.append(['', 'T-004', '某工程二层板', 'C30',
               datetime.datetime(2025, 11, 14), datetime.datetime(2025, 12, 5), ''])
    wb.create_sheet('无关台账').append(['x'])
    wb.save(path)
    wb.close()


def make_ledger(path):
    """逐日温度台账：与真实台账同构（左 B/F/J/M，右 P/T/X/AA）。

    注意 J 列是「累计温度」、F 列是「日均温」，两者刻意取不同值，
    用于验证脚本取的是日均温而不是累计温度。
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    temps = [10.0, -3.0, 0.0, 4.0, 6.0, 8.0, 12.5, 5.0, 7.5, 9.0]
    run = 0.0
    for i, t in enumerate(temps):
        d = BASE + datetime.timedelta(days=i)
        run = round(run + max(t, 0), 1)
        r = 6 + i
        ws['B%d' % r] = '%d年%02d月%02d日' % (d.year, d.month, d.day)
        ws['F%d' % r] = t            # 日均温
        ws['J%d' % r] = run          # 累计温度（应被忽略）
    wb.save(path)
    wb.close()


def make_target(path):
    """目标工作簿：Sheet1 作格式模板。"""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    ws['A1'] = '原表标题'
    for j, h in enumerate(['序号', '施工部位', '试块组数及编号', '试块制作日期',
                           '试块检验日期', '制作日期累计温度值', '检验日期累计温度值',
                           '600℃·d计算值', '等效龄期（d）'], start=1):
        ws.cell(row=2, column=j).value = h
    ws['A3'] = 1
    ws['C3'] = '旧数据'
    ws.merge_cells('A1:F1')
    ws.merge_cells('H1:I1')
    ws.column_dimensions['B'].width = 40
    ws.row_dimensions[5].height = 22
    wb.save(path)
    wb.close()


class TestRawXlsx(unittest.TestCase):

    def setUp(self):
        self.tmp = fresh_dir('raw')
        self.path = os.path.join(self.tmp, 'book.xlsx')
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = '数据'
        ws['A1'] = '编号'
        ws['B1'] = 'T-001'
        wb.save(self.path)
        wb.close()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_sheet_list_and_read(self):
        sheets = rawxlsx.sheet_list(self.path)
        self.assertEqual([s[1] for s in sheets], ['数据'])
        cells = rawxlsx.read_sheet(self.path, sheets[0][2])
        self.assertEqual(cells[(1, 1)], '编号')
        self.assertEqual(cells[(1, 2)], 'T-001')

    def test_col_helpers(self):
        self.assertEqual(rawxlsx.col_index('A'), 1)
        self.assertEqual(rawxlsx.col_index('AA'), 27)
        self.assertEqual(rawxlsx.col_letter(1), 'A')
        self.assertEqual(rawxlsx.col_letter(27), 'AA')


class TestEqage(unittest.TestCase):

    def setUp(self):
        self.tmp = fresh_dir('eqage')
        self.summary = os.path.join(self.tmp, 'summary.xlsx')
        self.ledger = os.path.join(self.tmp, 'ledger.xlsx')
        self.target = os.path.join(self.tmp, 'target.xlsx')
        make_summary(self.summary)
        make_ledger(self.ledger)
        make_target(self.target)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_discover_only_600_sheets(self):
        names = eqage.discover_sheets(self.summary, CFG)
        self.assertEqual(names, ['2#楼600°同条件试验台账'])

    def test_read_ledger_series(self):
        """日均温列与累计温度列分别读出，两者不可混用。"""
        avg, cum = eqage.read_ledger_series(self.ledger, CFG)
        self.assertEqual(avg[BASE], 10.0)                                # 日均温
        self.assertEqual(avg[BASE + datetime.timedelta(days=1)], -3.0)   # 负温原样保留
        self.assertEqual(avg[BASE + datetime.timedelta(days=2)], 0.0)
        self.assertEqual(avg[BASE + datetime.timedelta(days=4)], 6.0)
        # 累计温度列：只增不减，11-15 是 16.5
        self.assertEqual(cum[BASE], 10.0)
        self.assertEqual(cum[BASE + datetime.timedelta(days=1)], 10.0)   # 负温冻结
        self.assertEqual(cum[BASE + datetime.timedelta(days=4)], 20.0)
        self.assertNotEqual(avg[BASE + datetime.timedelta(days=4)],
                            cum[BASE + datetime.timedelta(days=4)])

    def test_read_specimens(self):
        specs = eqage.read_specimens(self.summary,
                                     ['2#楼600°同条件试验台账'], CFG)
        recs = specs['2#楼600°同条件试验台账']
        self.assertEqual([r['no'] for r in recs], ['T-001', 'T-002', 'T-003', 'T-004'])
        self.assertEqual(recs[0]['made'], datetime.date(2025, 11, 12))
        self.assertEqual(recs[0]['part'], '某工程一层柱')
        self.assertIsNone(recs[2]['made'])          # 空占位行
        self.assertEqual(recs[3]['sent'], datetime.date(2025, 12, 5))

    def test_parse_date_any(self):
        self.assertEqual(eqage.parse_date_any('2025年11月12日'), datetime.date(2025, 11, 12))
        self.assertEqual(eqage.parse_date_any('2025/11/12'), datetime.date(2025, 11, 12))
        self.assertEqual(eqage.parse_date_any('45973'), datetime.date(2025, 11, 12))
        self.assertIsNone(eqage.parse_date_any(''))
        self.assertIsNone(eqage.parse_date_any('不是日期'))

    def test_build_and_verify(self):
        before = openpyxl.load_workbook(self.target)
        s1_before = [before['Sheet1'].cell(row=2, column=c).value for c in range(1, 10)]

        stats = eqage.build(self.target, self.summary, self.ledger, CFG, write=True)
        self.assertEqual(len(stats), 1)
        st = stats[0]
        self.assertEqual(st['specimens'], 4)
        self.assertEqual(st['placeholder'], 1)

        book = rawxlsx.read_book(self.target)
        self.assertIn('2#楼600°同条件试验台账', book)
        self.assertIn('Sheet1', book)                 # 原表保留
        cells = book['2#楼600°同条件试验台账']

        # 表头与标题
        self.assertEqual(cells[(1, 1)], '600°C·d实体检验等效龄期计算表\n表C5-19-2')
        self.assertEqual(cells[(2, 1)], '序号')
        self.assertEqual(cells[(2, 9)], '等效龄期（d）')
        self.assertEqual(cells[(1, 8)], '02-01-C5-001')

        # T-001：制作 11-12，检验 11-20
        #   F = 制作日日均温 = -3.0 ；G = 检验日日均温 = 9.0
        #   H = 累计温(11-20) − 累计温(11-12) = 62.0 − 10.0 = 52.0 ；I = 8
        self.assertEqual(cells[(3, 3)], 'T-001')
        self.assertEqual(float(cells[(3, 6)]), -3.0)
        self.assertEqual(float(cells[(3, 7)]), 9.0)
        self.assertEqual(float(cells[(3, 8)]), 52.0)
        self.assertEqual(float(cells[(3, 9)]), 8)

        # T-002：制作 11-13（日均温 0.0），检验 12-01 超出温度范围 → G/H 留空，I = 18
        self.assertEqual(float(cells[(4, 6)]), 0.0)
        self.assertNotIn((4, 7), cells)
        self.assertNotIn((4, 8), cells)
        self.assertEqual(float(cells[(4, 9)]), 18)

        # T-003 空占位：只有序号与编号
        self.assertEqual(cells[(5, 3)], 'T-003')
        self.assertNotIn((5, 4), cells)
        self.assertEqual(cells[(5, 1)], '3')

        # T-004：制作 11-14（日均温 4.0），检验 12-05 超出范围
        self.assertEqual(float(cells[(6, 6)]), 4.0)
        self.assertEqual(float(cells[(6, 9)]), 21)

        # 制表人行
        self.assertEqual(cells[(8, 1)], '制表人：')

        # Sheet1 未被改动
        after = openpyxl.load_workbook(self.target)
        self.assertEqual(
            [after['Sheet1'].cell(row=2, column=c).value for c in range(1, 10)],
            s1_before)
        self.assertEqual(after['Sheet1']['C3'].value, '旧数据')

    def test_H_is_cumulative_difference(self):
        """H = 累计温(检验日) − 累计温(制作日)，取自累计温度列。"""
        avg, cum = eqage.read_ledger_series(self.ledger, CFG)
        d1 = datetime.date(2025, 11, 12)     # 累计 10.0
        d2 = datetime.date(2025, 11, 18)     # 累计 45.5
        self.assertEqual(cum[d1], 10.0)
        self.assertEqual(cum[d2], 45.5)
        self.assertAlmostEqual(round(cum[d2] - cum[d1], 1), 35.5)
        # H 不应等于日均温之差（那样只有 12.0）
        self.assertNotAlmostEqual(round(avg[d2] - avg[d1], 1), 35.5)

    def test_F_G_are_daily_mean_not_cumulative(self):
        """F/G 是当天平均温度；若误取累计温度列，量级会从个位数跳到几十上百。"""
        avg, cum = eqage.read_ledger_series(self.ledger, CFG)
        for d in (BASE, BASE + datetime.timedelta(days=4)):
            self.assertLess(abs(avg[d]), 50, 'F/G 出现了累计温度量级，说明取错列了')
        self.assertGreater(cum[BASE + datetime.timedelta(days=4)], avg[BASE + datetime.timedelta(days=4)])

    def test_dry_run_does_not_write(self):
        mtime = os.path.getmtime(self.target)
        eqage.build(self.target, self.summary, self.ledger, CFG, write=False)
        self.assertEqual(os.path.getmtime(self.target), mtime)
        book = rawxlsx.read_book(self.target)
        self.assertNotIn('2#楼600°同条件试验台账', book)


if __name__ == '__main__':
    unittest.main(verbosity=2)
