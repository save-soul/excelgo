"""生成含图表的 xlsx，供 Go 侧测试复制/合并是否正确搬运图表部件。

openpyxl 能创建图表并写出完整部件链：
    xl/worksheets/sheet1.xml  -> <drawing r:id="..."/>
    xl/worksheets/_rels/sheet1.xml.rels -> ../drawings/drawing1.xml
    xl/drawings/drawing1.xml  -> <c:chart r:id="..."/>
    xl/drawings/_rels/drawing1.xml.rels -> ../charts/chart1.xml
    xl/charts/chart1.xml      -> 图表定义（数据缓存 c:numCache 等）

这是**多级关系链**（sheet -> drawing -> chart），比单层的 media 复杂。
"""
import sys

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference


def build(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["月份", "产值", "成本"])
    for i in range(1, 7):
        ws.append([f"2026-0{i}", i * 100, i * 80])

    # 条形图
    bar = BarChart()
    bar.title = "产值成本对比"
    data = Reference(ws, min_col=2, max_col=3, min_row=1, max_row=7)
    cats = Reference(ws, min_col=1, min_row=2, max_row=7)
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    ws.add_chart(bar, "E2")

    # 折线图（第二个图表，验证同一 sheet 多图表）
    line = LineChart()
    line.title = "趋势"
    data2 = Reference(ws, min_col=2, min_row=1, max_row=7)
    line.add_data(data2, titles_from_data=True)
    line.set_categories(cats)
    ws.add_chart(line, "E18")

    # 第二个 sheet 也带图
    ws2 = wb.create_sheet("Second")
    ws2.append(["项目", "金额"])
    for i in range(1, 4):
        ws2.append([f"项目{i}", i * 500])
    bar2 = BarChart()
    bar2.title = "Second"
    d = Reference(ws2, min_col=2, min_row=1, max_row=4)
    c = Reference(ws2, min_col=1, min_row=2, max_row=4)
    bar2.add_data(d, titles_from_data=True)
    bar2.set_categories(c)
    ws2.add_chart(bar2, "D2")

    wb.save(path)
    return path


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "charts.xlsx"
    build(out)
    print("BUILD_OK " + out)
