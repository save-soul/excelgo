"""把 xlsx 归一化为可跨实现比较的 JSON 快照。

设计要点：
1. **归一化到语义层**，不比较 XML 字节。两侧（excelgo / openpyxl）的 XML 排布
   必然不同，比字节毫无意义；比"用户能观察到的语义"才有意义。
2. **只输出有值/有设置的部分**，跳过默认值。openpyxl 读回时会补出大量默认值
   （如 sheetView 默认选中、pageMargins 默认值），若不剔除会产生假差异。
3. **数值统一用 repr(float)**，避免 int/float 抖动导致的假差异。
4. **警告不污染快照**。openpyxl 会对"缺少 default style"之类的结构差异发
   UserWarning；那是诊断信息（写 stderr），不是快照内容。快照只走 stdout，
   以免 Go 侧把警告误判成失败。

用法：
    python snapshot.py <file.xlsx>

输出：JSON 到 stdout。以 "SNAPSHOT_OK " 开头，Go 侧据此判定成功。
"""

import json
import sys
import warnings

import openpyxl
from openpyxl.utils import get_column_letter

# openpyxl 对"与它自己产出的结构差异"会发一堆 UserWarning（缺 default style、
# 未支持的扩展等）。这些是诊断而非错误：收集后写 stderr，退出码仍为 0。
_warnings = []


def _collect_warning(message, category, filename, lineno, file=None, line=None):
    _warnings.append("%s: %s" % (category.__name__, message))


warnings.showwarning = _collect_warning


def norm_num(v):
    """数值归一化：整数不带 .0，浮点保留 6 位有效小数。"""
    if isinstance(v, bool):
        return v
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        if v == int(v):
            return int(v)
        return round(v, 6)
    return v


def cell_value(c):
    """把单元格值归一化为可比较形式。

    区分类型：数字 / 字符串 / 布尔 / 公式 / None。
    **公式保留前导 "="**：openpyxl 读回是 "=SUM(...)"，excelgo 读回也是
    "=SUM(...)"，两侧一致。统一剥掉反而会让本库读回值与快照无法直接比对，
    且掩盖"某侧漏了 = 号"这类真实问题。
    """
    v = c.value
    if v is None:
        return None
    if isinstance(v, bool):
        return {"t": "bool", "v": v}
    if isinstance(v, (int, float)):
        return {"t": "num", "v": norm_num(v)}
    if isinstance(v, str):
        if v.startswith("="):
            return {"t": "formula", "v": v}
        return {"t": "str", "v": v}
    # 日期时间等
    return {"t": "other", "v": str(v)}


def sheet_snapshot(ws):
    """单表快照。"""
    s = {}

    # --- 单元格（只收有值的）---
    cells = {}
    for row in ws.iter_rows():
        for c in row:
            cv = cell_value(c)
            if cv is not None:
                cells[c.coordinate] = cv
    if cells:
        s["cells"] = cells

    # --- 合并区域 ---
    merges = sorted(str(m) for m in ws.merged_cells.ranges)
    if merges:
        s["merges"] = merges

    # --- 列宽（只收显式设置过的）---
    widths = {}
    for letter, dim in ws.column_dimensions.items():
        if dim.width is not None:
            w = round(float(dim.width), 4)
            # openpyxl 默认列宽约 13.0，非默认才记录
            if abs(w - 13.0) > 1e-6:
                widths[letter] = w
    if widths:
        s["col_widths"] = widths

    # --- 行高 ---
    heights = {}
    for idx, dim in ws.row_dimensions.items():
        if dim.height is not None:
            heights[str(idx)] = round(float(dim.height), 4)
    if heights:
        s["row_heights"] = heights

    # --- 冻结窗格 ---
    if ws.freeze_panes:
        s["freeze_panes"] = str(ws.freeze_panes)

    # --- 自动筛选 ---
    try:
        if ws.auto_filter and ws.auto_filter.ref:
            s["auto_filter"] = str(ws.auto_filter.ref)
    except Exception:
        pass

    # --- 打印设置 ---
    pa = str(ws.print_area) if ws.print_area else ""
    if pa:
        # openpyxl 读回时可能带 sheet 前缀与引号：'Sheet1'!$A$1:$D$20
        # 归一化掉前缀与引号，只留区域本体，两侧才能对齐
        s["print_area"] = pa.split("!")[-1].replace("'", "").replace("$", "")
    ptr = str(ws.print_title_rows) if ws.print_title_rows else ""
    if ptr:
        s["print_title_rows"] = ptr.replace("$", "")

    # --- 表格 ---
    # 记录 displayName 与 ref 而不只是名字：名字在复制时会被避让改名（X_1），
    # 只比名字会造成假差异；但"表格是否跟着搬走、区域是否一致"是真正的保真指标。
    if getattr(ws, "tables", None):
        tbls = []
        for tname, tref in sorted(ws.tables.items()):
            entry = {"name": tname, "ref": str(tref)}
            tbls.append(entry)
        s["tables"] = tbls

    # --- 图片数量 ---
    # openpyxl 读 WPS 的 x14 内嵌图片时可能识别不出（它只认传统 drawing），
    # 故只记录数量、不记录细节，作为"有图"的粗粒度信号；精确性由 Go 侧
    # 直接查 zip 部件表保证。
    imgs = getattr(ws, "_images", None)
    if imgs:
        s["image_count"] = len(imgs)

    # --- 数据验证 ---
    dvs = []
    for dv in ws.data_validations.dataValidation:
        f1 = str(dv.formula1) if dv.formula1 else ""
        f1 = f1.lstrip("=")
        dvs.append({"type": dv.type, "sqref": str(dv.sqref), "f1": f1})
    if dvs:
        dvs.sort(key=lambda d: (d["sqref"], d["type"]))
        s["data_validations"] = dvs

    # --- 条件格式 ---
    cfs = []
    for rng in ws.conditional_formatting:
        for rule in rng.rules:
            cfs.append({
                "range": str(rng.sqref),
                "type": rule.type,
                "formula": list(rule.formula) if rule.formula else [],
            })
    if cfs:
        cfs.sort(key=lambda c: (c["range"], c["type"]))
        s["conditional_formats"] = cfs

    # --- 超链接 ---
    hls = {}
    for row in ws.iter_rows():
        for c in row:
            if c.hyperlink is not None:
                tgt = c.hyperlink.target or c.hyperlink.location or ""
                hls[c.coordinate] = tgt
    if hls:
        s["hyperlinks"] = hls

    # --- 批注 ---
    cms = {}
    for row in ws.iter_rows():
        for c in row:
            if c.comment is not None:
                cms[c.coordinate] = {
                    "text": c.comment.text,
                    "author": c.comment.author or "",
                }
    if cms:
        s["comments"] = cms

    # --- 单元格样式（仅收有实际格式的）---
    styles = {}
    for row in ws.iter_rows():
        for c in row:
            st = c._style
            if st is None:
                continue
            sig = {
                "bold": bool(c.font and c.font.bold),
                "italic": bool(c.font and c.font.italic),
                "underline": (c.font.underline if c.font and c.font.underline else None),
                "strike": bool(c.font and c.font.strike),
                "size": norm_num(c.font.size) if c.font and c.font.size else None,
                "name": (c.font.name if c.font else None),
                "color": _rgb(c.font.color) if c.font else None,
                "fill": _fill_rgb(c.fill),
                "numfmt": c.number_format if c.number_format else None,
                "halign": c.alignment.horizontal if c.alignment else None,
                "valign": c.alignment.vertical if c.alignment else None,
                "wrap": bool(c.alignment and c.alignment.wrap_text),
                "border": _border_sig(c),
            }
            # 只记录与默认不同的项。字体名/字号是 openpyxl 的默认值，
            # 未显式设置样式时也会带出来，会造成假差异，故按默认值剔除。
            diff = {k: v for k, v in sig.items() if v not in (None, False, "General")}
            if diff.get("name") in (None, "Calibri"):
                diff.pop("name", None)
            if diff.get("size") in (None, 11):
                diff.pop("size", None)
            if diff:
                styles[c.coordinate] = diff
    if styles:
        s["styles"] = styles

    return s


def _rgb(color):
    """安全取颜色的 ARGB 值，统一为 8 位大写十六进制。

    两个必须处理的坑：
    1. openpyxl 的 Color.rgb 在颜色为 theme/indexed 时是一个**描述错误的对象**
       （打印出来是 "Values must be of type <class 'str'>"），直接 str() 会产出
       无意义的假差异，必须按类型严格校验。
    2. 6 位与 8 位（有无 alpha）都要归一到 8 位，否则 excelgo 写 ARGB、openpyxl
       经 op.py 补 FF 前缀，两侧会因格式不同产生假差异。
    """
    if color is None:
        return None
    try:
        v = color.rgb
    except Exception:
        return None
    if not isinstance(v, str):
        return None
    v = v.strip().upper()
    if len(v) == 8 and all(c in "0123456789ABCDEF" for c in v):
        return v
    if len(v) == 6 and all(c in "0123456789ABCDEF" for c in v):
        return "FF" + v
    return None


def _fill_rgb(fill):
    """取填充色；仅 solid 填充才算。"""
    if fill is None or fill.patternType != "solid":
        return None
    return _rgb(fill.fgColor)


def _border_sig(c):
    """边框签名：只记录设置了样式的边。"""
    b = c.border
    if b is None:
        return None
    out = {}
    for side in ("left", "right", "top", "bottom"):
        s = getattr(b, side, None)
        if s is not None and s.style:
            out[side] = s.style
    return out or None


def snapshot(path):
    """整个工作簿的快照。"""
    wb = openpyxl.load_workbook(path, data_only=False)
    out = {"sheets": {}}
    for name in wb.sheetnames:
        out["sheets"][name] = sheet_snapshot(wb[name])
    # 工作表顺序（复制/移动操作会改）
    out["order"] = list(wb.sheetnames)
    # 定义名称
    dn = {}
    for k, v in wb.defined_names.items():
        dn[k] = str(v.value).replace("$", "")
    if dn:
        out["defined_names"] = dn
    return out


def main():
    if len(sys.argv) < 2:
        print("usage: snapshot.py <file.xlsx>", file=sys.stderr)
        sys.exit(2)
    path = sys.argv[1]
    try:
        snap = snapshot(path)
    except Exception as e:
        print("SNAPSHOT_FAIL %s: %s: %s" % (
            path, type(e).__name__, e), file=sys.stderr)
        sys.exit(1)
    # 诊断信息走 stderr，绝不与 stdout 的快照 JSON 混在一起
    for w in _warnings:
        print("SNAPSHOT_WARN " + w, file=sys.stderr)
    print("SNAPSHOT_OK " + json.dumps(snap, ensure_ascii=False, sort_keys=True))


if __name__ == "__main__":
    main()
