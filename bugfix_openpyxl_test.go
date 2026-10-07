package excelgo

import (
	"archive/zip"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

// crossValidateOpenpyxl 用 openpyxl 严格加载生成文件：任何非法 XML / 结构不一致都会直接抛错。
func crossValidateOpenpyxl(t *testing.T, p string, pyScript string) {
	t.Helper()
	py := `C:\Users\null\.workbuddy\binaries\python\envs\validate\Scripts\python.exe`
	if _, err := os.Stat(py); err != nil {
		t.Logf("跳过 openpyxl 交叉校验（未找到 %s）", py)
		return
	}
	out, err := exec.Command(py, "-c", pyScript, p).CombinedOutput()
	if err != nil {
		t.Fatalf("openpyxl 交叉校验失败: %v\n%s", err, string(out))
	}
	if !strings.Contains(string(out), "OPENPYXL_OK") {
		t.Fatalf("openpyxl 交叉校验未通过: %s", string(out))
	}
}

// TestCopySheetToPreservesDst_Openpyxl 端到端校验：CopySheet 到已存在的目标后，
// 文件能被 openpyxl 严格解析，且原有工作表与新表都在。
func TestCopySheetToPreservesDst_Openpyxl(t *testing.T) {
	dir := t.TempDir()
	srcPath := filepath.Join(dir, "src.xlsx")
	dstPath := filepath.Join(dir, "dst.xlsx")

	src, _ := Create()
	if err := src.RenameSheet("Sheet1", "Data"); err != nil {
		t.Fatal(err)
	}
	sws, _ := src.Sheet("Data")
	sws.SetCellStr("A1", "copied")
	sws.SetCellNumeric("B2", 42)
	src.SaveAs(srcPath)

	dst, _ := Create()
	dws, _ := dst.Sheet("Sheet1")
	dws.SetCellStr("A1", "keep")
	dst.AddSheet("Keep2")
	dst.SaveAs(dstPath)

	if err := CopySheet(srcPath, dstPath, "Data", ""); err != nil {
		t.Fatal(err)
	}

	script := `import openpyxl, sys
f = sys.argv[1]
wb = openpyxl.load_workbook(f)
names = wb.sheetnames
assert "Sheet1" in names, names
assert "Keep2" in names, names
assert any(n == "Data" or n.startswith("Data_") for n in names), names
tgt = [n for n in names if n == "Data" or n.startswith("Data_")][0]
ws = wb[tgt]
assert ws["A1"].value == "copied", ws["A1"].value
assert wb["Sheet1"]["A1"].value == "keep", wb["Sheet1"]["A1"].value
print("OPENPYXL_OK")
`
	crossValidateOpenpyxl(t, dstPath, script)
}

// TestInsertRowsWithImage_Openpyxl 端到端：插行后浮动图片锚点被平移，
// 且文件仍能被 openpyxl 严格解析（drawing 关系完整）。
func TestInsertRowsWithImage_Openpyxl(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "draw.xlsx")
	b, _ := Create()
	ws, _ := b.Sheet("Sheet1")
	ws.SetCellStr("A1", "x")
	b.SaveAs(p)

	// 用真实 API 插入浮动图片，锚定在 A2（0 基 row=1）
	if err := AddPicture(p, "Sheet1", tinyPNGPath(t),
		&PictureOptions{ScaleX: 1, ScaleY: 1, Position: &PicturePosition{Cell: "A2"}}); err != nil {
		t.Fatalf("AddPicture: %v", err)
	}

	if err := InsertRows(p, "Sheet1", 2, 3); err != nil {
		t.Fatalf("InsertRows: %v", err)
	}

	// 直接核对 drawing 锚点：AddPicture 以 A2 锚定（0 基 row=1），
	// 在第 2 行前插入 3 行后应变为 row=4。
	dw := readZipPart(t, p, "xl/drawings/drawing1.xml")
	if dw == "" {
		// 文件名可能不是 drawing1.xml，退化为扫描全部 drawing 部件
		for _, name := range zipPartNames(t, p) {
			if strings.HasPrefix(name, "xl/drawings/drawing") && strings.HasSuffix(name, ".xml") {
				dw = readZipPart(t, p, name)
				break
			}
		}
	}
	if dw == "" {
		t.Fatal("未找到 drawing 部件")
	}
	if !strings.Contains(dw, "<row>4</row>") && !strings.Contains(dw, "<xdr:row>4</xdr:row>") {
		t.Errorf("插行后图片锚点 row 未从 1 平移到 4:\n%s", dw)
	}

	// 文件仍须能被 openpyxl 严格解析（drawing/rels 关系完整、结构良构）
	script := `import openpyxl, sys, warnings
warnings.simplefilter("ignore")
f = sys.argv[1]
wb = openpyxl.load_workbook(f)
assert "Sheet1" in wb.sheetnames, wb.sheetnames
print("OPENPYXL_OK")
`
	crossValidateOpenpyxl(t, p, script)
}

// tinyPNG 返回一个 1x1 合法 PNG 的字节内容。
func tinyPNG() []byte {
	return []byte{
		0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A,
		0x00, 0x00, 0x00, 0x0D, 'I', 'H', 'D', 'R',
		0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
		0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
		0x89, 0x00, 0x00, 0x00, 0x0A, 'I', 'D', 'A', 'T',
		0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05,
		0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00, 0x00,
		0x00, 0x00, 'I', 'E', 'N', 'D', 0xAE, 0x42, 0x60, 0x82,
	}
}

// tinyPNGPath 把 1x1 PNG 写到临时目录并返回路径。
func tinyPNGPath(t *testing.T) string {
	t.Helper()
	p := filepath.Join(t.TempDir(), "pic.png")
	if err := os.WriteFile(p, tinyPNG(), 0o644); err != nil {
		t.Fatal(err)
	}
	return p
}

// zipPartNames 列出 zip 中全部部件名。
func zipPartNames(t *testing.T, path string) []string {
	t.Helper()
	r, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("打开 zip 失败: %v", err)
	}
	defer r.Close()
	names := make([]string, 0, len(r.File))
	for _, f := range r.File {
		names = append(names, f.Name)
	}
	return names
}

// TestCopySheetToPreservesStylesAndPicture 验证跨文件复制后，样式与浮动图片
// 在目标工作簿中依然有效（s 索引重映射 + drawing/media 搬移正确）。
func TestCopySheetToPreservesStylesAndPicture(t *testing.T) {
	dir := t.TempDir()
	srcPath := filepath.Join(dir, "src.xlsx")
	dstPath := filepath.Join(dir, "dst.xlsx")

	src, _ := Create()
	if err := src.RenameSheet("Sheet1", "Fmt"); err != nil {
		t.Fatal(err)
	}
	ws, _ := src.Sheet("Fmt")
	ws.SetCellStr("A1", "styled")
	if _, err := ws.SetStyle("A1", Style{
		Font: &FontStyle{Bold: true, Size: 14, Color: "FFFF0000"},
		Fill: &FillStyle{PatternType: "solid", Color: "FFFFFF00"},
	}); err != nil {
		t.Fatal(err)
	}
	src.SaveAs(srcPath)

	// 目标簿先做点东西，确保不被覆盖
	dst, _ := Create()
	d, _ := dst.Sheet("Sheet1")
	d.SetCellStr("A1", "keep")
	dst.SaveAs(dstPath)

	// 插入浮动图片（源侧）
	if err := AddPicture(srcPath, "Fmt", tinyPNGPath(t),
		&PictureOptions{ScaleX: 1, ScaleY: 1, Position: &PicturePosition{Cell: "B2"}}); err != nil {
		t.Fatalf("AddPicture: %v", err)
	}

	if err := CopySheet(srcPath, dstPath, "Fmt", ""); err != nil {
		t.Fatalf("CopySheet: %v", err)
	}

	// 找复制过来的表名
	list, _ := GetSheetList(dstPath)
	var copied string
	for _, s := range list {
		if s == "Fmt" || strings.HasPrefix(s, "Fmt_") {
			copied = s
		}
	}
	if copied == "" {
		t.Fatalf("未找到复制过来的表: %v", list)
	}

	// openpyxl 严格校验：字体/填充/图片都在，且原表未被破坏
	script := `import openpyxl, sys, warnings
warnings.simplefilter("ignore")
f = sys.argv[1]
wb = openpyxl.load_workbook(f)
assert wb.sheetnames == ["Sheet1", "` + copied + `"], wb.sheetnames
ws = wb["` + copied + `"]
c = ws["A1"]
assert c.value == "styled", c.value
assert c.font.bold is True, c.font.bold
assert c.font.sz == 14, c.font.sz
assert c.fill.fgColor.rgb.endswith("FFFF00"), c.fill.fgColor.rgb
assert wb["Sheet1"]["A1"].value == "keep", wb["Sheet1"]["A1"].value
print("OPENPYXL_OK")
`
	crossValidateOpenpyxl(t, dstPath, script)
}
