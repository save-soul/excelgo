"""用 openpyxl 执行一组操作，产出与 excelgo 侧同构的 xlsx，供差分比较。

操作以 JSON 从 stdin 读入，格式：
    {"out": "path.xlsx", "ops": [{"op": "set_str", "sheet": "S1", "cell": "A1",
                                "value": "x"}, ...]}

设计原则：**每个操作的语义必须与 Go 侧的 excelgo 调用一一对应**。
若某操作 openpyxl 无法等价表达，在 ops 里标记 "skip_reason"，
Go 侧会把它记为"不可比"而非"差异"—— 区分"真差异"与"能力差异"很重要。

用法：
    python op.py <ops.json>
"""

import json
import sys

import openpyxl
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
import openpyxl.worksheet.protection
import openpyxl.workbook.protection


def apply_ops(spec):
    out = spec["out"]
    ops = spec["ops"]

    # 决定从新建还是已有文件开始
    base = spec.get("base")
    if base:
        wb = openpyxl.load_workbook(base)
    else:
        wb = openpyxl.Workbook()
        # openpyxl 新建工作簿的默认表名是 "Sheet"（excelgo 是 "Sheet1"）。
        # 两侧统一到 "Sheet1"，否则表名本身就会成为假差异。
        wb.active.title = spec.get("first_sheet", "Sheet1")

    for o in ops:
        op = o["op"]
        fn = HANDLERS.get(op)
        if fn is None:
            raise SystemExit("unknown op: %s" % op)
        try:
            fn(wb, o)
        except Exception as e:
            # 操作失败也要记录，让 Go 侧看到 openpyxl 侧的真实行为
            print("OP_FAIL %s: %s: %s" % (op, type(e).__name__, e),
                  file=sys.stderr)

    wb.save(out)


# ---------- 单元格 ----------

def _ws(wb, name):
    return wb[name]


def op_create_sheet(wb, o):
    wb.create_sheet(title=o["name"])


def op_set_str(wb, o):
    _ws(wb, o["sheet"])[o["cell"]] = o["value"]


def op_set_num(wb, o):
    _ws(wb, o["sheet"])[o["cell"]] = o["value"]


def op_set_bool(wb, o):
    _ws(wb, o["sheet"])[o["cell"]] = o["value"]


def op_set_formula(wb, o):
    # openpyxl 公式需带前导 =
    f = o["formula"]
    if not f.startswith("="):
        f = "=" + f
    _ws(wb, o["sheet"])[o["cell"]] = f


def op_set_value(wb, o):
    """按运行时类型写值（对应 excelgo SetCellValue）。"""
    v = o["value"]
    if isinstance(v, bool):
        _ws(wb, o["sheet"])[o["cell"]] = v
    elif isinstance(v, (int, float)):
        _ws(wb, o["sheet"])[o["cell"]] = v
    else:
        _ws(wb, o["sheet"])[o["cell"]] = str(v)


def op_set_range(wb, o):
    ws = _ws(wb, o["sheet"])
    for i, row in enumerate(o["values"]):
        for j, v in enumerate(row):
            if v is None:
                continue
            cell = "%s%d" % (get_column_letter(j + 1), i + 1)
            if isinstance(v, bool):
                ws[cell] = v
            elif isinstance(v, (int, float)):
                ws[cell] = v
            else:
                ws[cell] = str(v)


# ---------- 样式 ----------

def _argb(c):
    """把 6 位 RGB 或 8 位 ARGB 统一成 8 位 ARGB。

    两侧用例必须传入**完全相同**的颜色字面量，否则测的是用例不对称
    而非实现差异。统一在这里做 alpha 补全，避免用例里出现 "FF" + 8位
    这种重复前缀。
    """
    c = str(c).strip().upper().lstrip("#")
    if len(c) == 6:
        return "FF" + c
    if len(c) == 8:
        return c
    raise SystemExit("bad color: %r" % c)


def op_set_style(wb, o):
    ws = _ws(wb, o["sheet"])
    st = o["style"]
    cell = ws[o["cell"]]
    f = {}
    if "bold" in st:
        f["bold"] = st["bold"]
    if "underline" in st and st["underline"]:
        f["underline"] = st["underline"]
    if "strike" in st and st["strike"]:
        f["strike"] = st["strike"]
    if "italic" in st:
        f["italic"] = st["italic"]
    if "size" in st:
        f["size"] = st["size"]
    if "name" in st:
        f["name"] = st["name"]
    if "color" in st and st["color"]:
        f["color"] = _argb(st["color"])
    if f:
        cell.font = Font(**f)
    if "fill" in st and st["fill"]:
        cell.fill = PatternFill(
            start_color=_argb(st["fill"]),
            end_color=_argb(st["fill"]), fill_type="solid")
    if "numfmt" in st and st["numfmt"]:
        cell.number_format = st["numfmt"]
    al = {}
    if "halign" in st and st["halign"]:
        al["horizontal"] = st["halign"]
    if "valign" in st and st["valign"]:
        al["vertical"] = st["valign"]
    if "wrap" in st:
        al["wrap_text"] = st["wrap"]
    if al:
        cell.alignment = Alignment(**al)
    if "border" in st and st["border"]:
        sides = {}
        for name, sv in st["border"].items():
            sides[name] = Side(style=sv)
        cell.border = Border(**sides)


def op_copy_style(wb, o):
    ws = _ws(wb, o["sheet"])
    src = ws[o["from"]]
    dst = ws[o["to"]]
    dst.font = src.font.copy()
    dst.fill = src.fill.copy()
    dst.border = src.border.copy()
    dst.alignment = src.alignment.copy()
    dst.number_format = src.number_format


# ---------- 行列 ----------

def op_col_width(wb, o):
    ws = _ws(wb, o["sheet"])
    # min/max 优先；缺省时用 col。两侧语义必须对齐，否则会出现假差异。
    lo = o.get("min", o.get("col"))
    hi = o.get("max", o.get("col"))
    if lo is None or hi is None:
        raise SystemExit("col_width 需要 col 或 min/max")
    for c in range(lo, hi + 1):
        ws.column_dimensions[get_column_letter(c)].width = o["width"]


def op_row_height(wb, o):
    ws = _ws(wb, o["sheet"])
    lo = o.get("min", o.get("row"))
    hi = o.get("max", o.get("row"))
    if lo is None or hi is None:
        raise SystemExit("row_height 需要 row 或 min/max")
    for r in range(lo, hi + 1):
        ws.row_dimensions[r].height = o["height"]


def op_insert_rows(wb, o):
    ws = _ws(wb, o["sheet"])
    ws.insert_rows(o["row"], o.get("count", 1))


def op_insert_cols(wb, o):
    ws = _ws(wb, o["sheet"])
    ws.insert_cols(o["col"], o.get("count", 1))


def op_remove_rows(wb, o):
    ws = _ws(wb, o["sheet"])
    ws.delete_rows(o["row"], o.get("count", 1))


def op_remove_cols(wb, o):
    ws = _ws(wb, o["sheet"])
    ws.delete_cols(o["col"], o.get("count", 1))


def op_group_rows(wb, o):
    ws = _ws(wb, o["sheet"])
    ws.row_dimensions.group(o["r1"], o["r2"], outline_level=o.get("level", 1))


def op_group_cols(wb, o):
    ws = _ws(wb, o["sheet"])
    ws.column_dimensions.group(
        get_column_letter(o["c1"]), get_column_letter(o["c2"]),
        outline_level=o.get("level", 1))


# ---------- 布局 ----------

def op_merge(wb, o):
    _ws(wb, o["sheet"]).merge_cells(o["ref"])


def op_unmerge(wb, o):
    _ws(wb, o["sheet"]).unmerge_cells(o["ref"])


def op_freeze(wb, o):
    _ws(wb, o["sheet"]).freeze_panes = o["ref"]


def op_autofilter(wb, o):
    _ws(wb, o["sheet"]).auto_filter.ref = o["ref"]


def op_print_area(wb, o):
    _ws(wb, o["sheet"]).print_area = o["ref"]


def op_print_titles(wb, o):
    _ws(wb, o["sheet"]).print_title_rows = o["rows"]


# ---------- 其它功能 ----------

def op_comment(wb, o):
    ws = _ws(wb, o["sheet"])
    c = Comment(o["text"], o.get("author", ""))
    ws[o["cell"]].comment = c


def op_hyperlink(wb, o):
    ws = _ws(wb, o["sheet"])
    cell = ws[o["cell"]]
    cell.value = o.get("display", o["url"])
    cell.hyperlink = o["url"]


def op_table(wb, o):
    ws = _ws(wb, o["sheet"])
    t = Table(displayName=o["name"], ref=o["ref"])
    t.tableStyleInfo = TableStyleInfo(
        name="TableStyleLight1", showRowStripes=True)
    ws.add_table(t)


def op_datavalidation(wb, o):
    ws = _ws(wb, o["sheet"])
    if "values" in o:
        f1 = '"' + ",".join(o["values"]) + '"'
    else:
        f1 = o.get("formula1", "")
    dv = DataValidation(
        type=o.get("type", "list"), formula1=f1,
        allow_blank=o.get("allow_blank", True))
    ws.add_data_validation(dv)
    dv.add(o["ref"])


def op_defined_name(wb, o):
    from openpyxl.workbook.defined_name import DefinedName
    wb.defined_names[o["name"]] = DefinedName(o["name"], attr_text=o["refers_to"])


def op_rename_sheet(wb, o):
    wb[o["old"]].title = o["new"]


def op_delete_sheet(wb, o):
    del wb[o["name"]]


def op_sheet_visible(wb, o):
    wb[o["name"]].sheet_state = o["state"]


def op_tab_color(wb, o):
    wb[o["name"]].sheet_properties.tabColor = _argb(o["color"])


# ---------- 复合操作（对标 excelgo 的 CopySheet / MergeWorkbook） ----------


def op_copy_sheet(wb, o):
    """同工作簿内复制工作表（对标 excelgo CopySheet）。

    openpyxl 的 copy_worksheet 会带上值、样式、合并、列宽/行高，
    但**不搬图片/图表/条件格式/数据验证/批注** —— openpyxl 已知限制。
    这些差异在 Go 侧按字段路径豁免，并写明理由。
    """
    src = wb[o["sheet"]]
    dst = wb.copy_worksheet(src)
    dst.title = o.get("new_name") or (src.title + "_copy")
    # copy_worksheet 不搬打印设置，手动补齐以保证可比
    if "print_area" in o:
        dst.print_area = o["print_area"]
    if "print_titles" in o:
        dst.print_title_rows = o["print_titles"]


def op_move_sheet(wb, o):
    wb.move_sheet(o["sheet"], offset=o.get("offset", 0))





def op_protect(wb, o):
    """保护工作表（对标 excelgo ProtectSheet）。

    openpyxl 的 password 传**明文**，它内部会算 OOXML legacy hash；
    excelgo 侧则要求调用方自己传已哈希值。故两侧最终 XML 应一致。
    """
    ws = _ws(wb, o["sheet"])
    kw = {"sheet": True, "objects": True, "scenarios": True}
    if "password" in o and o["password"]:
        kw["password"] = o["password"]
    ws.protection = openpyxl.worksheet.protection.SheetProtection(**kw)


def op_protect_workbook(wb, o):
    """工作簿结构保护（对标 excelgo ProtectWorkbook）。"""
    # 注意：WorkbookProtection 的参数名是 workbookPassword，不是 password
    kw = {"lockStructure": True}
    if "password" in o and o["password"]:
        kw["workbookPassword"] = o["password"]
    wb.security = openpyxl.workbook.protection.WorkbookProtection(**kw)


HANDLERS = {
    "copy_sheet": op_copy_sheet,
    "move_sheet": op_move_sheet,
    "create_sheet": op_create_sheet,
    "set_str": op_set_str,
    "set_num": op_set_num,
    "set_bool": op_set_bool,
    "set_formula": op_set_formula,
    "set_value": op_set_value,
    "set_range": op_set_range,
    "set_style": op_set_style,
    "copy_style": op_copy_style,
    "col_width": op_col_width,
    "row_height": op_row_height,
    "insert_rows": op_insert_rows,
    "insert_cols": op_insert_cols,
    "remove_rows": op_remove_rows,
    "remove_cols": op_remove_cols,
    "group_rows": op_group_rows,
    "group_cols": op_group_cols,
    "merge": op_merge,
    "unmerge": op_unmerge,
    "freeze": op_freeze,
    "autofilter": op_autofilter,
    "print_area": op_print_area,
    "print_titles": op_print_titles,
    "comment": op_comment,
    "hyperlink": op_hyperlink,
    "table": op_table,
    "datavalidation": op_datavalidation,
    "defined_name": op_defined_name,
    "rename_sheet": op_rename_sheet,
    "delete_sheet": op_delete_sheet,
    "sheet_visible": op_sheet_visible,
    "tab_color": op_tab_color,
    "protect": op_protect,
    "protect_workbook": op_protect_workbook,
}


def main():
    spec = json.load(sys.stdin)
    apply_ops(spec)
    print("OPS_OK")


if __name__ == "__main__":
    main()
