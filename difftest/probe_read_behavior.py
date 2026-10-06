"""针对本库 excelgo 的高风险行为做差分探测。

这里不测"API 有没有"，而是测**行为对不对** —— 因为短板往往不是缺功能，
而是同一件事的语义与 openpyxl 不一致，用户会踩坑。

每个探针对应一类高风险：
  P1 读空单元格 / 不存在的单元格 —— 是否能区分
  P2 读公式的结果值 —— data_only 语义
  P3 共享字符串里的富文本 —— 是否能还原
  P4 数字格式化的显示值 —— 库返回原始值还是格式化值
  P5 行/列越界的读行为
  P6 表名大小写与特殊字符
  P7 日期时间的读回类型
"""
import sys
import warnings

warnings.simplefilter("ignore")

import openpyxl

path = sys.argv[1]
wb = openpyxl.load_workbook(path, data_only=False)
out = {}

for name in wb.sheetnames:
    ws = wb[name]
    info = {}

    def enc(v):
        """把 openpyxl 的值编码成可与 Go 侧直接比较的形式。

        必须区分 bool 与数字：Python 里 True == 1，Go 里亦然，
        不显式编码就会把 bool 与 int 混为一谈。
        """
        if v is None:
            return {"k": "nil"}
        if isinstance(v, bool):
            return {"k": "bool", "v": "TRUE" if v else "FALSE"}
        if isinstance(v, (int, float)):
            return {"k": "num", "v": str(int(v)) if v == int(v) else repr(v)}
        if isinstance(v, str):
            if v.startswith("="):
                return {"k": "formula", "v": v}
            return {"k": "str", "v": v}
        return {"k": "other", "v": str(v)}

    # P1 空单元格
    for coord in ("A1", "B1", "C1", "A2", "Z99", "E1", "F1", "G1", "G2", "B2"):
        c = ws[coord]
        info["P1_" + coord] = {"value": enc(c.value), "data_type": c.data_type}
    # P2 公式
    for coord in ("D1", "D2"):
        c = ws[coord]
        info["P2_" + coord] = {
            "value": enc(c.value),
            "has_formula": isinstance(c.value, str) and c.value.startswith("="),
        }
    # P7 日期
    for coord in ("E1", "E2"):
        c = ws[coord]
        info["P7_" + coord] = {
            "value": enc(c.value),
            "is_date": c.is_date,
            "numfmt": c.number_format,
        }
    # P4 数字格式化的显示值
    for coord in ("F1", "F2"):
        c = ws[coord]
        info["P4_" + coord] = {"raw": str(c.value), "numfmt": c.number_format}
    # P3 共享字符串富文本
    for coord in ("G1", "G2"):
        c = ws[coord]
        info["P3_" + coord] = {"value": c.value, "type": type(c.value).__name__}
    out[name] = info

import json
print("PROBE_OK " + json.dumps(out, ensure_ascii=False, sort_keys=True))
