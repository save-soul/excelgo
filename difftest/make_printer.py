# 生成一份含打印机设置二进制的工作簿，用于测试 excelgo 的搬运保真。
#
# 打印机设置（xl/printerSettings/printerSettingsN.bin）是 Excel/WPS 保存工作簿时
# 写下的打印机专属二进制块（纸张、驱动、边距微调等）。它由 worksheet 的
# <pageSetup r:id="rIdN"> 引用，是 worksheet rels 里最容易在复制/合并时丢掉的一条链：
# 丢了 r:id 就悬空（文件损坏），留着 r:id 但部件没搬同样悬空。
#
# openpyxl 不提供 API 写它，只能构造一份等价文件来验证 excelgo。
import sys
import zipfile

SRC = sys.argv[1] if len(sys.argv) > 1 else "base.xlsx"
OUT = sys.argv[2] if len(sys.argv) > 2 else "prn.xlsx"

# 合法的打印机设置 DEVMODE 结构：dmDeviceName(32) + dmSpecVersion 等字段
# 这里只需是一个非空二进制块；Excel 遇到无法识别的 DEVMODE 会退回默认打印机，
# 不会判文件损坏 —— 所以测试只关心"部件在不在、关系通不通"。
DM = bytearray(b"\x00" * 220)
name = "Microsoft Print to PDF".encode("utf-16-le")
DM[0:len(name)] = name                      # dmDeviceName
DM[68:70] = (0x0401).to_bytes(2, "little")   # dmSpecVersion
DM[70:72] = (0x0600).to_bytes(2, "little")   # dmDriverVersion
DM[72:74] = (220).to_bytes(2, "little")      # dmSize

z = zipfile.ZipFile(SRC)
names = z.namelist()

# 1) 改写 sheet1.xml：在 pageSetup 上加 r:id，并声明 r 命名空间
sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8")
if "xmlns:r=" not in sheet.split(">", 1)[0]:
    sheet = sheet.replace(
        "<worksheet ",
        '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
        1,
    )
if "<pageSetup" in sheet:
    # 已有 pageSetup：加 r:id
    import re
    sheet = re.sub(r"<pageSetup\b([^>]*?)/>", r'<pageSetup\1 r:id="rIdPrn"/>', sheet, count=1)
    if 'r:id="rIdPrn"' not in sheet:  # pageSetup 不是自闭合的形态
        sheet = re.sub(r"<pageSetup\b([^>]*?)>",
                       r'<pageSetup\1 r:id="rIdPrn">', sheet, count=1)
else:
    sheet = sheet.replace(
        "</worksheet>",
        '<pageSetup paperSize="9" orientation="portrait" r:id="rIdPrn"/></worksheet>',
    )

# 2) 改写 sheet1.xml.rels：加入指向 printerSettings 的关系
rel_name = "xl/worksheets/_rels/sheet1.xml.rels"
if rel_name in names:
    rels = z.read(rel_name).decode("utf-8")
else:
    rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        "</Relationships>"
    )
if "rIdPrn" not in rels:
    rels = rels.replace(
        "</Relationships>",
        '<Relationship Id="rIdPrn" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" '
        'Target="../printerSettings/printerSettings1.bin"/>'
        "</Relationships>",
    )

# 3) [Content_Types].xml：bin 扩展的 Default 声明
ct = z.read("[Content_Types].xml").decode("utf-8")
if 'Extension="bin"' not in ct:
    ct = ct.replace(
        "<Types ",
        '<Types ',
        1,
    ).replace(
        ">",
        '><Default Extension="bin" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"/>',
        1,
    )

out = zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED)
written = set()
for n in names:
    written.add(n)
    if n == "xl/worksheets/sheet1.xml":
        out.writestr(n, sheet)
    elif n == rel_name:
        out.writestr(n, rels)
    elif n == "[Content_Types].xml":
        out.writestr(n, ct)
    else:
        out.writestr(n, z.read(n))
# 上面循环只在 rels 原本就存在时才会写；新建的必须补上，否则关系文件丢失
if rel_name not in written:
    out.writestr(rel_name, rels)
out.writestr("xl/printerSettings/printerSettings1.bin", bytes(DM))
out.close()
z.close()
print("BUILD_OK")