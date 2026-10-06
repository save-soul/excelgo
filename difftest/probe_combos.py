import os
import zipfile
import warnings

warnings.simplefilter("ignore")
import openpyxl

# 探测 openpyxl 对各种"组合场景"的实际接受度。
# 目的：找出 excelgo 侧尚未验证、但 openpyxl 会当场拒收的组合。
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def probe(name, fn):
    try:
        fn()
        print("OK    %s" % name)
    except Exception as e:
        print("FAIL  %s -> %s: %s" % (name, type(e).__name__, str(e)[:110]))


# 1) 图表 + 命名样式 + 条件格式 + 冻结窗格 同时存在
def combo1():
    from openpyxl import Workbook
    from openpyxl.chart import BarChart, Reference
    from openpyxl.styles import NamedStyle
    from openpyxl.formatting.rule import ColorScaleRule

    wb = Workbook()
    ws = wb.active
    ws.title = "组合"
    ws.append(["类", "值"])
    for i in range(1, 5):
        ws.append(["C%d" % i, i * 10])
    from openpyxl.styles import Font
    ns = NamedStyle(name="我的样式")
    ns.font = Font(bold=True)
    wb.add_named_style(ns)
    ws["A1"].style = "我的样式"
    ws.conditional_formatting.add(
        "B2:B5",
        ColorScaleRule(start_type="min", start_color="FFFFFFFF",
                       end_type="max", end_color="FFFF0000"),
    )
    ch = BarChart()
    ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=5), titles_from_data=True)
    ws.add_chart(ch, "D2")
    ws.freeze_panes = "B2"
    p = os.path.join(BASE, "probe_combo1.xlsx")
    wb.save(p)
    openpyxl.load_workbook(p)


# 2) 数据验证 + 条件格式 + 合并单元格 + 表格 同一区域
def combo2():
    from openpyxl import Workbook
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "组合2"
    ws.append(["名称", "数量"])
    for i in range(1, 6):
        ws.append(["项%d" % i, i])
    dv = DataValidation(type="list", formula1='"甲,乙,丙"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("A2:A6")
    t = Table(displayName="表1", ref="A1:B6")
    t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium9", showRowStripes=True)
    ws.add_table(t)
    ws.conditional_formatting.add(
        "B2:B6", CellIsRule(operator="greaterThan", formula=["3"],
                            fill=PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")))
    p = os.path.join(BASE, "probe_combo2.xlsx")
    wb.save(p)
    openpyxl.load_workbook(p)


# 3) 图表 + 图片 + 打印设置 三者共存
def combo3():
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, Reference
    from openpyxl.drawing.image import Image

    wb = Workbook()
    ws = wb.active
    ws.title = "组合3"
    ws.append(["x", "y"])
    for i in range(1, 5):
        ws.append([i, i * 2])
    ch = LineChart()
    ch.add_data(Reference(ws, min_col=2, min_row=1, max_row=5))
    ws.add_chart(ch, "F2")
    ws.print_area = "A1:B5"
    ws.page_setup.orientation = "landscape"
    ws.oddHeader.center.text = "第 &P 页"
    p = os.path.join(BASE, "probe_combo3.xlsx")
    wb.save(p)
    openpyxl.load_workbook(p)


# 4) 多个图表 + 多个表格 + 多个数据验证 叠加
def combo4():
    from openpyxl import Workbook
    from openpyxl.chart import LineChart, BarChart, PieChart, Reference
    from openpyxl.worksheet.table import Table, TableStyleInfo

    wb = Workbook()
    ws = wb.active
    ws.title = "多"
    ws.append(["类", "v1", "v2", "v3"])
    for i in range(1, 6):
        ws.append(["C%d" % i, i, i * 2, i * 3])
    for k, cls in enumerate([LineChart, BarChart, PieChart]):
        c = cls()
        c.add_data(Reference(ws, min_col=2 + (k % 3), min_row=1, max_row=6),
                   titles_from_data=True)
        ws.add_chart(c, "F%d" % (2 + k * 16))
    for i, ref in enumerate(["A1:B6", "A1:C6", "A1:D6"]):
        t = Table(displayName="T%d" % (i + 1), ref=ref)
        t.tableStyleInfo = TableStyleInfo(name="TableStyleLight%d" % (i + 1))
        ws.add_table(t)
    p = os.path.join(BASE, "probe_combo4.xlsx")
    wb.save(p)
    openpyxl.load_workbook(p)


# 5) 隐藏行/列 + 大纲折叠 + 条件格式
def combo5():
    from openpyxl import Workbook
    from openpyxl.formatting.rule import FormulaRule
    from openpyxl.styles import PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "折叠"
    for i in range(1, 21):
        ws.cell(row=i, column=1, value=i)
    ws.row_dimensions.group(5, 12, outline_level=1, hidden=True)
    ws.column_dimensions.group("C", "D", outline_level=1, hidden=True)
    ws.conditional_formatting.add(
        "A1:A20",
        FormulaRule(formula=["$A1>10"],
                    fill=PatternFill(start_color="FFFFFF00", end_color="FFFFFF00", fill_type="solid")),
    )
    p = os.path.join(BASE, "probe_combo5.xlsx")
    wb.save(p)
    openpyxl.load_workbook(p)


# 6) 超链接 + 批注 + 保护 三者共存
def combo6():
    from openpyxl import Workbook
    from openpyxl.comments import Comment
    from openpyxl.worksheet.protection import SheetProtection

    wb = Workbook()
    ws = wb.active
    ws.title = "批注"
    ws["A1"] = "链接"
    ws["A1"].hyperlink = "https://example.com"
    ws["A1"].comment = Comment("这是批注", "作者")
    ws.protection = SheetProtection(sheet=True, password="1234")
    p = os.path.join(BASE, "probe_combo6.xlsx")
    wb.save(p)
    w2 = openpyxl.load_workbook(p)
    w2["批注"].protection.sheet = True


for nm, fn in [("图表+命名样式+条件格式+冻结", combo1),
               ("数据验证+表格+条件格式+合并", combo2),
               ("图表+打印设置+页眉", combo3),
               ("多图表+多表格", combo4),
               ("隐藏行列+折叠+条件格式", combo5),
               ("超链接+批注+保护", combo6)]:
    probe(nm, fn)