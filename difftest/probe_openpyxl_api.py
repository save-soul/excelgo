"""提取 openpyxl 的能力面，用于与 excelgo 做能力对比。

用类字典反射而非 dir() 实例 —— dir(实例) 会触发属性求值，
openpyxl 某些属性求值会抛异常（merged_cell_ranges 曾在空表上报错）。

输出：按类别分组的能力清单，供 Go 侧或人工比对。
"""
import inspect
import warnings

warnings.simplefilter("ignore")

import openpyxl
from openpyxl.chart import BarChart
from openpyxl.drawing.image import Image
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.workbook.workbook import Workbook


def methods_of(cls):
    out = []
    for name, obj in vars(cls).items():
        if name.startswith("_"):
            continue
        if inspect.isfunction(obj) or inspect.ismethod(obj):
            out.append(name)
        elif isinstance(obj, property):
            out.append(name + "()")
    return sorted(out)


def props_of(cls):
    out = []
    for klass in cls.__mro__:
        for name, obj in vars(klass).items():
            if name.startswith("_") or not isinstance(obj, property):
                continue
            if name not in out:
                out.append(name)
    return sorted(out)


wb_methods = methods_of(Workbook)
ws_methods = methods_of(Worksheet)
wb_props = props_of(Workbook)
ws_props = props_of(Worksheet)

print("### Workbook")
print("方法(%d): %s" % (len(wb_methods), ", ".join(wb_methods)))
print()
print("属性(%d): %s" % (len(wb_props), ", ".join(wb_props)))
print()
print("### Worksheet")
print("方法(%d): %s" % (len(ws_methods), ", ".join(ws_methods)))
print()
print("属性(%d): %s" % (len(ws_props), ", ".join(ws_props)))
print()

# 关键子模块是否存在
mods = {
    "chart": ["BarChart", "LineChart", "PieChart", "ScatterChart", "BubbleChart",
              "AreaChart", "RadarChart", "DoughnutChart", "StockChart",
              "SurfaceChart", "BarChart3D"],
    "drawing": ["Image"],
    "pivot": ["TableDefinition", "CacheDefinition"],
    "worksheet.table": ["Table", "TableStyleInfo"],
    "worksheet.datavalidation": ["DataValidation"],
    "worksheet.protection": ["SheetProtection"],
    "formatting.rule": ["CellIsRule", "FormulaRule", "ColorScaleRule",
                        "DataBarRule", "IconSetRule"],
    "comments": ["Comment"],
    "workbook.properties": ["CalcProperties"],
    "utils": ["get_column_letter", "column_index_from_string"],
    "writerex": ["Workbook", "WriteOnlyCell"],
    "reader": ["load_workbook", "read_only"],
    "writer": ["save_workbook"],
}
print("### 子模块能力")
for m, items in mods.items():
    try:
        mod = __import__("openpyxl." + m, fromlist=items)
        have = [i for i in items if hasattr(mod, i)]
        print("  %-28s %d/%d  %s" % (m, len(have), len(items), ", ".join(have)))
    except Exception as e:
        print("  %-28s ERR %s" % (m, e))
