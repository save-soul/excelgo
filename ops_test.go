package excelgo

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

const opsTestSrc = "../testdata_tmp/full.xlsx"

// copyTestFile 把源复制到临时文件，避免污染样本。
func copyTestFile(t *testing.T, src, dst string) {
	t.Helper()
	data, err := os.ReadFile(src)
	if err != nil {
		t.Fatalf("读源文件失败: %v", err)
	}
	if err := os.WriteFile(dst, data, 0644); err != nil {
		t.Fatalf("写临时文件失败: %v", err)
	}
}

func TestSheetLifecycle(t *testing.T) {
	tmp := filepath.Join(t.TempDir(), "lifecycle.xlsx")
	copyTestFile(t, opsTestSrc, tmp)

	// 新建
	idx, err := NewSheet(tmp, "新建表")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	if idx != 4 {
		t.Errorf("NewSheet 返回序号应为 4，实际 %d", idx)
	}
	// 改名（避开同名自动加后缀）
	if err := RenameSheet(tmp, "新建表", "改名表"); err != nil {
		t.Fatalf("RenameSheet: %v", err)
	}
	// 移动：把 Sheet2 移到最前
	if err := MoveSheet(tmp, "Sheet2", 1); err != nil {
		t.Fatalf("MoveSheet: %v", err)
	}
	// 删除 Sheet3（带图片，应清理其专属 media）
	if err := DeleteSheet(tmp, "Sheet3"); err != nil {
		t.Fatalf("DeleteSheet: %v", err)
	}
	// 校验可再次打开且结构合法（交给 openpyxl；此处仅确认无panic）
	if _, err := os.Stat(tmp); err != nil {
		t.Fatalf("产物不存在: %v", err)
	}
}

func TestCellReadWrite(t *testing.T) {
	tmp := filepath.Join(t.TempDir(), "cell.xlsx")
	copyTestFile(t, opsTestSrc, tmp)

	// 写多种格式
	if err := SetCellStr(tmp, "Sheet1", "F1", "hello"); err != nil {
		t.Fatal(err)
	}
	if err := SetCellInt(tmp, "Sheet1", "F2", 42); err != nil {
		t.Fatal(err)
	}
	if err := SetCellNumeric(tmp, "Sheet1", "F3", 3.14); err != nil {
		t.Fatal(err)
	}
	if err := SetCellBool(tmp, "Sheet1", "F4", true); err != nil {
		t.Fatal(err)
	}
	if err := SetCellFormula(tmp, "Sheet1", "F5", "F2*2", "84"); err != nil {
		t.Fatal(err)
	}
	if err := SetCellValue(tmp, "Sheet1", "F6", "auto-string"); err != nil {
		t.Fatal(err)
	}

	// 读回
	checks := map[string]string{"F1": "hello", "F2": "42", "F3": "3.14", "F4": "1", "F6": "auto-string"}
	for c, want := range checks {
		got, err := GetCell(tmp, "Sheet1", c)
		if err != nil {
			t.Fatalf("GetCell %s: %v", c, err)
		}
		if got != want {
			t.Errorf("单元格 %s 读写不一致: 期望 %q，实际 %q", c, want, got)
		}
	}
	// 公式单元格读回结果
	if got, _ := GetCell(tmp, "Sheet1", "F5"); got != "84" {
		t.Errorf("公式结果应为 84，实际 %q", got)
	}
	// 原有数据不应被破坏
	if got, _ := GetCell(tmp, "Sheet1", "A1"); got != "标题" {
		t.Errorf("原有 A1 被破坏: %q", got)
	}
}

func TestAddPictures(t *testing.T) {
	tmp := filepath.Join(t.TempDir(), "pic.xlsx")
	copyTestFile(t, opsTestSrc, tmp)
	picRed := "../testdata_tmp/ph_red.png"
	picBlue := "../testdata_tmp/ph_blue.png"

	// 浮动图片（带位置与缩放）
	if err := AddPicture(tmp, "Sheet1", picRed, &PictureOptions{
		ScaleX: 0.5, ScaleY: 0.5,
		Position: &PicturePosition{Cell: "H1", ColOffset: 10, RowOffset: 5},
	}); err != nil {
		t.Fatalf("AddPicture 浮动: %v", err)
	}
	// WPS 单元格内嵌图片
	if err := AddCellPicture(tmp, "Sheet2", "B2", picBlue); err != nil {
		t.Fatalf("AddCellPicture: %v", err)
	}
	// 再次在 Sheet1 追加一张浮动图（验证复用已有 drawing）
	if err := AddPicture(tmp, "Sheet1", picBlue, nil); err != nil {
		t.Fatalf("AddPicture 复用 drawing: %v", err)
	}
}

// TestWriteVerifyArtifacts 把各功能产物写到 testdata_tmp 固定路径，供 openpyxl 离线校验。
// 该测试专注于生成产物，不在此断言；断言在外部 Python 校验脚本进行。
func TestWriteVerifyArtifacts(t *testing.T) {
	base := "../testdata_tmp"
	gen := func(name string) string {
		p := filepath.Join(base, name)
		copyTestFile(t, opsTestSrc, p)
		return p
	}

	// 生命周期
	lc := gen("verify_lifecycle.xlsx")
	NewSheet(lc, "新建表")
	RenameSheet(lc, "新建表", "改名表")
	MoveSheet(lc, "Sheet2", 1)
	DeleteSheet(lc, "Sheet3")

	// 单元格
	cl := gen("verify_cell.xlsx")
	SetCellStr(cl, "Sheet1", "F1", "hello")
	SetCellInt(cl, "Sheet1", "F2", 42)
	SetCellNumeric(cl, "Sheet1", "F3", 3.14)
	SetCellBool(cl, "Sheet1", "F4", true)
	SetCellFormula(cl, "Sheet1", "F5", "F2*2", "84")
	SetCellValue(cl, "Sheet1", "F6", "auto-string")

	// 图片
	pic := gen("verify_pic.xlsx")
	AddPicture(pic, "Sheet1", "../testdata_tmp/ph_red.png", &PictureOptions{
		ScaleX: 0.5, ScaleY: 0.5,
		Position: &PicturePosition{Cell: "H1", ColOffset: 10, RowOffset: 5},
	})
	AddCellPicture(pic, "Sheet2", "B2", "../testdata_tmp/ph_blue.png")
	AddPicture(pic, "Sheet1", "../testdata_tmp/ph_blue.png", nil)
}

// TestLocateSheetResolvesByRID 回归测试：验证工作表文件按 RID（经 workbook.xml.rels）定位，
// 而非按 "第 N 个 sheet 即 sheet{N+1}.xml" 的脆弱假设。
//
// 真实（尤其 WPS 删改过）工作簿中，xl/worksheets/sheet{N}.xml 的文件名编号与 <sheet> 列表顺序
// 解耦。这里人为构造一个解耦样本：
//
//	SheetA(rId1) -> sheet3.xml
//	SheetB(rId2) -> sheet1.xml
//	SheetC(rId3) -> sheet2.xml
//
// 旧实现会因 sheet{idx+1}.xml 恰好存在而直接返回错误文件（如 SheetA 返回 sheet1.xml）；
// 修复后应一律按 rels 解析为目标文件。
func TestLocateSheetResolvesByRID(t *testing.T) {
	fm := map[string][]byte{
		"xl/workbook.xml": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets>
<sheet name="SheetA" sheetId="1" r:id="rId1"/>
<sheet name="SheetB" sheetId="2" r:id="rId2"/>
<sheet name="SheetC" sheetId="3" r:id="rId3"/>
</sheets>
</workbook>`),
		"xl/_rels/workbook.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
</Relationships>`),
		"xl/worksheets/sheet1.xml": []byte(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>`),
		"xl/worksheets/sheet2.xml": []byte(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>`),
		"xl/worksheets/sheet3.xml": []byte(`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData></sheetData></worksheet>`),
	}

	wb := &Workbook{}
	if err := getXMLFromMap(fm, "xl/workbook.xml", wb); err != nil {
		t.Fatalf("解析 workbook 失败: %v", err)
	}

	wantBy := map[string]string{
		"SheetA": "xl/worksheets/sheet3.xml",
		"SheetB": "xl/worksheets/sheet1.xml",
		"SheetC": "xl/worksheets/sheet2.xml",
	}
	for name, wantFile := range wantBy {
		f, gotName, _, _, err := locateSheetInMap(fm, wb, name)
		if err != nil {
			t.Fatalf("按名定位 %s 失败: %v", name, err)
		}
		if gotName != name {
			t.Errorf("定位 %s 返回的名字应为 %s，实际 %s", name, name, gotName)
		}
		if f != wantFile {
			t.Errorf("按名定位 %s 应映射到 %s（依据 rels RID），实际 %s", name, wantFile, f)
		}
	}

	// 按 1 基索引定位同样应遵循 rels：第 1 个 <sheet>(SheetA) -> sheet3.xml，而非 sheet1.xml
	f, _, _, _, err := locateSheetInMap(fm, wb, "1")
	if err != nil {
		t.Fatalf("按索引定位失败: %v", err)
	}
	if f != "xl/worksheets/sheet3.xml" {
		t.Errorf("按索引 1 定位应解析到 rels 指向的 sheet3.xml，实际 %s", f)
	}
}

// decoupledWorkbookMap 构造一个“文件名编号与列表顺序解耦”的最小合法 xlsx（模拟 WPS 删改后）：
//
//	SheetA(rId1) -> worksheets/sheet2.xml
//	SheetB(rId2) -> worksheets/sheet1.xml
//
// 即第 1 个 <sheet>(SheetA) 实际指向 sheet2.xml，旧实现会错误地当 sheet1.xml 处理。
func decoupledWorkbookMap() map[string][]byte {
	const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
		`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Default Extension="xml" ContentType="application/xml"/>` +
		`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
		`<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
		`<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
		`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
		`<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` +
		`</Types>`
	const rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
		`</Relationships>`
	const wb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<sheets>` +
		`<sheet name="SheetA" sheetId="1" r:id="rId1"/>` +
		`<sheet name="SheetB" sheetId="2" r:id="rId2"/>` +
		`</sheets></workbook>`
	const xlRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>` +
		`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
		`<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
		`<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
		`</Relationships>`
	const sheetA = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		`<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>SheetA-ORIG</t></is></c></row></sheetData></worksheet>`
	const sheetB = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		`<sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>SheetB-ORIG</t></is></c></row></sheetData></worksheet>`
	const styles = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></styleSheet>`
	const ss = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>`
	return map[string][]byte{
		"[Content_Types].xml":        []byte(ct),
		"_rels/.rels":                []byte(rootRels),
		"xl/workbook.xml":            []byte(wb),
		"xl/_rels/workbook.xml.rels": []byte(xlRels),
		"xl/worksheets/sheet1.xml":   []byte(sheetB), // SheetB
		"xl/worksheets/sheet2.xml":   []byte(sheetA), // SheetA
		"xl/styles.xml":              []byte(styles),
		"xl/sharedStrings.xml":       []byte(ss),
	}
}

func mustReadMap(t *testing.T, filename string) map[string][]byte {
	t.Helper()
	fm, err := readZipToMap(filename)
	if err != nil {
		t.Fatalf("读盘失败: %v", err)
	}
	return fm
}

// TestSheetOpsOnDecoupledFiles 端到端回归测试：在“文件名编号与列表顺序解耦”的工作簿上，
// 验证 GetCell / SetCellInline / DeleteSheet 都作用于正确的 worksheet 文件，
// 而非按位置默认 sheet{N+1}.xml（旧实现会写错/读错/删错表）。
func TestSheetOpsOnDecoupledFiles(t *testing.T) {
	tmp := filepath.Join(t.TempDir(), "decoupled.xlsx")
	if err := writeMapToZip(tmp, decoupledWorkbookMap()); err != nil {
		t.Fatalf("写测试文件失败: %v", err)
	}

	// 向 SheetB（实际是 sheet1.xml）写入标记
	if err := SetCellInline(tmp, "SheetB", "A1", "MARKER_B_999"); err != nil {
		t.Fatalf("SetCellInline SheetB: %v", err)
	}
	afterB := mustReadMap(t, tmp)
	if !strings.Contains(string(afterB["xl/worksheets/sheet1.xml"]), "MARKER_B_999") {
		t.Errorf("写入 SheetB 未落到正确的 sheet1.xml")
	}
	if strings.Contains(string(afterB["xl/worksheets/sheet2.xml"]), "MARKER_B_999") {
		t.Errorf("写入 SheetB 错误地落到了 sheet2.xml（SheetA 的文件）")
	}

	// 向 SheetA（实际是 sheet2.xml）写入标记
	if err := SetCellInline(tmp, "SheetA", "A1", "MARKER_A_888"); err != nil {
		t.Fatalf("SetCellInline SheetA: %v", err)
	}
	afterA := mustReadMap(t, tmp)
	if !strings.Contains(string(afterA["xl/worksheets/sheet2.xml"]), "MARKER_A_888") {
		t.Errorf("写入 SheetA 未落到正确的 sheet2.xml")
	}
	if strings.Contains(string(afterA["xl/worksheets/sheet1.xml"]), "MARKER_A_888") {
		t.Errorf("写入 SheetA 错误地落到了 sheet1.xml（SheetB 的文件）")
	}

	// GetCell 端到端读回各自正确的内容
	if got, _ := GetCell(tmp, "SheetB", "A1"); got != "MARKER_B_999" {
		t.Errorf("GetCell(SheetB) 应返回 MARKER_B_999，实际 %q", got)
	}
	if got, _ := GetCell(tmp, "SheetA", "A1"); got != "MARKER_A_888" {
		t.Errorf("GetCell(SheetA) 应返回 MARKER_A_888，实际 %q", got)
	}

	// 删除 SheetA（sheet2.xml）后：sheet2.xml 应消失，sheet1.xml 保留，workbook.xml 不再含 SheetA
	if err := DeleteSheet(tmp, "SheetA"); err != nil {
		t.Fatalf("DeleteSheet SheetA: %v", err)
	}
	afterDel := mustReadMap(t, tmp)
	if _, ok := afterDel["xl/worksheets/sheet2.xml"]; ok {
		t.Errorf("删除 SheetA 后 sheet2.xml 应被移除")
	}
	if _, ok := afterDel["xl/worksheets/sheet1.xml"]; !ok {
		t.Errorf("删除 SheetA 后 sheet1.xml（SheetB）应保留")
	}
	if strings.Contains(string(afterDel["xl/workbook.xml"]), "SheetA") {
		t.Errorf("删除 SheetA 后 workbook.xml 不应再含 SheetA")
	}
}

// TestLocateSheetErrorsWhenRelsUnresolvable 锁定“无法解析直接报错、绝不兜底”的硬约束：
// 当 RID 指向的工作表文件不存在、或 RID 在 rels 中根本不存在时，locateSheetInMap 必须返回
// 错误，绝不能回退到位置命名的 sheet{idx+1}.xml（即便该文件恰好存在，也是别的表，错！）。
func TestLocateSheetErrorsWhenRelsUnresolvable(t *testing.T) {
	const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
		`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
		`<Default Extension="xml" ContentType="application/xml"/>` +
		`</Types>`
	const wbXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<sheets>` +
		`<sheet name="SheetA" sheetId="1" r:id="rId1"/>` +
		`<sheet name="SheetB" sheetId="2" r:id="rId2"/>` +
		`</sheets></workbook>`
	// rId1 指向一个不存在的文件；rels 中根本不存在 rId2。
	// 同时 sheet1.xml / sheet2.xml 都在，旧实现会“兜底”错误地返回它们。
	const xlRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/MISSING.xml"/>` +
		`<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
		`</Relationships>`
	fm := map[string][]byte{
		"[Content_Types].xml":        []byte(ct),
		"xl/workbook.xml":            []byte(wbXML),
		"xl/_rels/workbook.xml.rels": []byte(xlRels),
		"xl/worksheets/sheet1.xml":   []byte(`<worksheet/>`),
		"xl/worksheets/sheet2.xml":   []byte(`<worksheet/>`),
		"xl/styles.xml":              []byte(`<styleSheet/>`),
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fm, "xl/workbook.xml", wb); err != nil {
		t.Fatalf("解析 workbook 失败: %v", err)
	}

	// 情况1：RID 解析出的目标文件不存在 -> 必须报错，绝不兜底 sheet1.xml
	f, _, _, _, err := locateSheetInMap(fm, wb, "SheetA")
	if err == nil {
		t.Fatalf("SheetA: RID 指向缺失文件时应当报错，却返回 %q", f)
	}
	if f == "xl/worksheets/sheet1.xml" {
		t.Errorf("SheetA: 不应兜底返回 sheet1.xml（那是错文件）")
	}

	// 情况2：RID 在 rels 中不存在 -> 必须报错，绝不兜底 sheet2.xml
	f, _, _, _, err = locateSheetInMap(fm, wb, "SheetB")
	if err == nil {
		t.Fatalf("SheetB: RID 不在 rels 中时应当报错，却返回 %q", f)
	}
	if f == "xl/worksheets/sheet2.xml" {
		t.Errorf("SheetB: 不应兜底返回 sheet2.xml（那是错文件）")
	}
}
