package excelgo

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"testing"
	"time"
)

// readZipPart 读取 zip 中指定部件的内容，部件不存在返回 ""。
// readZipPart 读出 zip 内指定部件的内容；部件不存在返回空串。
//
// 读取委托给 readAllFromZip（与 difftest_composite_test.go 共用同一实现）。
func readZipPart(t *testing.T, path, part string) string {
	t.Helper()
	r, err := zip.OpenReader(path)
	if err != nil {
		t.Fatalf("打开 zip 失败 %s: %v", path, err)
	}
	defer r.Close()
	for _, f := range r.File {
		if f.Name == part {
			return readAllFromZip(f)
		}
	}
	return ""
}

// ---------- 回归 #1：MergeWorkbook 部分失败后仍写出完整 rels ----------

// TestMergePartialFailureKeepsRelsConsistent 验证合并中途出错时，落盘文件的
// workbook.xml 与 workbook.xml.rels 保持一致（不会出现「workbook.xml 引用了新表，
// 但 rels 里没有对应 Relationship」导致的文件损坏）。
func TestMergePartialFailureKeepsRelsConsistent(t *testing.T) {
	dir := t.TempDir()

	// 目标工作簿：2 张表
	dstPath := filepath.Join(dir, "dst.xlsx")
	dst, err := Create()
	if err != nil {
		t.Fatal(err)
	}
	dst.AddSheet("S1")
	dst.AddSheet("S2")
	if err := dst.SaveAs(dstPath); err != nil {
		t.Fatal(err)
	}

	// 源 1：正常
	src1 := filepath.Join(dir, "src1.xlsx")
	s1, _ := Create()
	ws1, _ := s1.Sheet("Sheet1")
	ws1.SetCellStr("A1", "from-src1")
	s1.SaveAs(src1)

	// 源 2：请求一个不存在的表 → mergeOneSheet 必然失败，用于触发部分合并失败路径
	src2 := filepath.Join(dir, "src2.xlsx")
	s2, _ := Create()
	s2.SaveAs(src2)

	merr := MergeWorkbook(dstPath, []SourceRef{
		{Workbook: src1, Sheet: "Sheet1"},
		{Workbook: src2, Sheet: "NoSuchSheet"}, // 该源会失败 → 触发部分合并失败路径
	})
	if merr == nil {
		t.Fatal("含失败源时应返回错误（用以确认部分失败路径被覆盖）")
	}
	t.Logf("如期返回部分失败错误: %v", merr)

	// 关键校验：无论后续合并是否失败，已写入 workbook.xml 的 sheet 必须都能在 rels 中找到
	checkConsistent := func(tag string) {
		wb := readZipPart(t, dstPath, "xl/workbook.xml")
		rels := readZipPart(t, dstPath, "xl/_rels/workbook.xml.rels")
		if wb == "" || rels == "" {
			t.Fatalf("[%s] 缺少 workbook.xml 或 rels", tag)
		}
		for _, rid := range extractAttrValues(wb, "r:id") {
			if !strings.Contains(rels, `Id="`+rid+`"`) {
				t.Errorf("[%s] workbook.xml 引用 r:id=%s 但 rels 中缺失 → 文件不一致", tag, rid)
			}
		}
	}
	checkConsistent("合并成功后")

	// 目标表应仍在，且新增了 src1 的表
	list, err := GetSheetList(dstPath)
	if err != nil {
		t.Fatal(err)
	}
	if len(list) < 3 {
		t.Errorf("合并后工作表数 = %d（%v），期望 >= 3", len(list), list)
	}
}

// extractAttrValues 提取 XML 中所有指定属性名的值。
func extractAttrValues(xml, attr string) []string {
	re := regexp.MustCompile(`\b` + regexp.QuoteMeta(attr) + `="([^"]*)"`)
	ms := re.FindAllStringSubmatch(xml, -1)
	out := make([]string, 0, len(ms))
	for _, m := range ms {
		out = append(out, m[1])
	}
	return out
}

// ---------- 回归 #2：DeleteSheet 重映射所有表级 definedName ----------

// TestDeleteSheetReindexesAllDefinedNames 验证删除工作表后，
// _xlnm.Print_Titles 与用户自建表级命名区域也会被正确重映射/清理，
// 而不是只处理 Print_Area。
func TestDeleteSheetReindexesAllDefinedNames(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "dn.xlsx")
	b, _ := Create() // 含默认 Sheet1
	b.AddSheet("S1")
	b.AddSheet("S2")
	b.AddSheet("S3")
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}

	// Create() 已建 Sheet1，故实际 4 张表：Sheet1(0) S1(1) S2(2) S3(3)。
	// 给每张表都加 Print_Area / Print_Titles / 自定义表级名，再加一个工作簿级名。
	bb, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	names := ""
	for i := 0; i < 4; i++ {
		n := strconv.Itoa(i)
		names += `<definedName name="_xlnm.Print_Area" localSheetId="` + n + `">Sheet!$A$1:$B$2</definedName>`
		names += `<definedName name="_xlnm.Print_Titles" localSheetId="` + n + `">Sheet!$1:$1</definedName>`
		names += `<definedName name="MyRange` + n + `" localSheetId="` + n + `">Sheet!$C$1</definedName>`
	}
	names += `<definedName name="GlobalName">Sheet!$D$1</definedName>`
	bb.fileMap["xl/workbook.xml"] = injectDefinedNames(bb.fileMap["xl/workbook.xml"], names)
	if err := bb.Save(); err != nil {
		t.Fatal(err)
	}

	// 删除中间的 S1（序号 1）
	if err := DeleteSheet(p, "S1"); err != nil {
		t.Fatalf("DeleteSheet: %v", err)
	}

	xml := readZipPart(t, p, "xl/workbook.xml")
	// 被删表（序号 1）的所有区域都应消失
	for _, gone := range []string{`name="MyRange1"`, `name="MyRange0" localSheetId="1"`} {
		if strings.Contains(xml, gone) {
			t.Errorf("已删除表的 definedName 未清理: %s\n%s", gone, xml)
		}
	}
	if strings.Contains(xml, `localSheetId="1">Sheet!$A$1:$B$2`) &&
		strings.Contains(xml, `<definedName name="MyRange1"`) {
		t.Errorf("已删除表的区域未清理\n%s", xml)
	}
	// 原序号 2 的 S2 应重映射为 1
	if !strings.Contains(xml, `<definedName name="_xlnm.Print_Titles" localSheetId="1">Sheet!$1:$1</definedName>`) {
		t.Errorf("S2 的 Print_Titles 未重映射到 localSheetId=1\n%s", xml)
	}
	if !strings.Contains(xml, `name="MyRange2" localSheetId="1"`) {
		t.Errorf("S2 的自定义表级命名区域未重映射到 localSheetId=1\n%s", xml)
	}
	// 原序号 3 的 S3 应重映射为 2
	if !strings.Contains(xml, `name="MyRange3" localSheetId="2"`) {
		t.Errorf("S3 的自定义表级命名区域未重映射到 localSheetId=2\n%s", xml)
	}
	// 未受影响的表保持原样
	if !strings.Contains(xml, `name="MyRange0" localSheetId="0"`) {
		t.Errorf("Sheet1 的区域不应被改动\n%s", xml)
	}
	// 工作簿级命名区域不受影响
	if !strings.Contains(xml, `<definedName name="GlobalName">`) {
		t.Errorf("工作簿级命名区域不应受影响\n%s", xml)
	}
}

// injectDefinedNames 把 names 插入到 workbook.xml 的 </sheets> 之后（合法位置）。
func injectDefinedNames(wbXML []byte, names string) []byte {
	s := string(wbXML)
	block := "<definedNames>" + names + "</definedNames>"
	if i := strings.Index(s, "<definedNames"); i != -1 {
		j := strings.Index(s[i:], "</definedNames>")
		if j != -1 {
			return []byte(s[:i] + block + s[i+j+len("</definedNames>"):])
		}
	}
	if i := strings.Index(s, "</sheets>"); i != -1 {
		return []byte(s[:i+len("</sheets>")] + block + s[i+len("</sheets>"):])
	}
	return []byte(s + block)
}

// ---------- 回归 #3：布尔单元格两条读路径一致 ----------

// TestBoolReadConsistency 验证同一布尔单元格经 GetCellValue / Cell.Get / GetRows
// 读出的表示一致（统一为 TRUE / FALSE）。
func TestBoolReadConsistency(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "bool.xlsx")
	b, _ := Create()
	ws, _ := b.Sheet("Sheet1")
	if err := ws.SetCellBool("A1", true); err != nil {
		t.Fatal(err)
	}
	if err := ws.SetCellBool("A2", false); err != nil {
		t.Fatal(err)
	}
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}

	b2, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	ws2, _ := b2.Sheet("Sheet1")

	if got := ws2.Cell("A1").Get(); got != "TRUE" {
		t.Errorf("Cell.Get(A1) = %q，期望 TRUE", got)
	}
	if got := ws2.Cell("A2").Get(); got != "FALSE" {
		t.Errorf("Cell.Get(A2) = %q，期望 FALSE", got)
	}
	v, err := GetCell(p, "Sheet1", "A1")
	if err != nil {
		t.Fatal(err)
	}
	if v != "TRUE" {
		t.Errorf("GetCell(A1) = %q，期望 TRUE", v)
	}

	rows, err := GetRows(p, "Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	if len(rows) < 2 || len(rows[0]) < 1 || rows[0][0] != "TRUE" {
		t.Errorf("GetRows 首行 = %v，期望首格为 TRUE", rows)
	}
	if rows[1][0] != "FALSE" {
		t.Errorf("GetRows 第二行首格 = %q，期望 FALSE", rows[1][0])
	}
}

// ---------- 回归 #4：RenameSheet 找不到源表时报错 ----------

func TestRenameSheetMissingNameReturnsError(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "rn.xlsx")
	b, _ := Create()
	b.SaveAs(p)

	b2, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	if err := b2.RenameSheet("NotExist", "X"); err == nil {
		t.Error("RenameSheet 对不存在的表名应返回错误，而不是静默 nil")
	}
	// 正常改名仍应成功
	if err := b2.RenameSheet("Sheet1", "Y"); err != nil {
		t.Errorf("RenameSheet 正常改名失败: %v", err)
	}
}

// ---------- 回归 #5：全局命名区域不误改表级同名区域 ----------

func TestDefinedNameGlobalDoesNotClobberSheetScoped(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "dn2.xlsx")
	b, _ := Create()
	b.AddSheet("S1")
	b.SaveAs(p)

	bb, _ := Open(p)
	names := `<definedName name="Dup" localSheetId="0">S1!$A$1</definedName>`
	bb.fileMap["xl/workbook.xml"] = injectDefinedNames(bb.fileMap["xl/workbook.xml"], names)
	if err := bb.Save(); err != nil {
		t.Fatal(err)
	}

	bb2, _ := Open(p)
	if err := bb2.SetDefinedName("Dup", "S1!$B$2"); err != nil {
		t.Fatal(err)
	}
	xml := string(bb2.fileMap["xl/workbook.xml"])
	if !strings.Contains(xml, `<definedName name="Dup" localSheetId="0">S1!$A$1</definedName>`) {
		t.Errorf("设置全局同名区域时误改了表级区域:\n%s", xml)
	}
	if !strings.Contains(xml, `<definedName name="Dup">S1!$B$2</definedName>`) {
		t.Errorf("全局命名区域未写入:\n%s", xml)
	}
}

// ---------- 回归 #6：absRange 产出真正的绝对引用 ----------

func TestAbsRangeProducesAbsoluteRefs(t *testing.T) {
	cases := map[string]string{
		"A1:D20":       "$A$1:$D$20",
		" A1 : D20 ":   "$A$1:$D$20",
		"$A$1:$D$20":   "$A$1:$D$20", // 已有 $ 必须保留
		"$A1:$D20":     "$A$1:$D$20",
		"A1":           "$A$1",
		"1:1":          "$1:$1",
		"SUM(A1:B2)":   "SUM(A1:B2)", // 非纯引用不改写
		"A1:D20:E30":   "A1:D20:E30", // 三段不是合法范围
		"":             "",
		"Sheet1!A1:B2": "Sheet1!A1:B2", // 带表名前缀不改写
	}
	for in, want := range cases {
		if got := absRange(in); got != want {
			t.Errorf("absRange(%q) = %q，期望 %q", in, got, want)
		}
	}
	if got := absRangeRows("1:1"); got != "$1:$1" {
		t.Errorf("absRangeRows(\"1:1\") = %q，期望 $1:$1", got)
	}
}

// ---------- 回归 #7：unescapeXML 解开双重转义 ----------

// TestUnescapeXMLSinglePass 验证实体解码。
// 注意：按 XML 语义，"&amp;lt;" 表示字面文本 "&lt;"（而非 "<"），
// 单次扫描解码即正确，不可循环解码（否则会破坏合法数据）。
func TestUnescapeXMLSinglePass(t *testing.T) {
	cases := map[string]string{
		"a&amp;b":       "a&b",
		"&lt;b&gt;":     "<b>",
		"&quot;x&quot;": `"x"`,
		"&amp;lt;":      "&lt;", // 字面文本 "&lt;"，不是 "<"
		"&amp;":         "&",
		"plain":         "plain",
		"a &lt; b":      "a < b",
	}
	for in, want := range cases {
		if got := unescapeXML(in); got != want {
			t.Errorf("unescapeXML(%q) = %q，期望 %q", in, got, want)
		}
	}
}

// TestEscapeRoundTrip 验证转义/解码可逆。
func TestEscapeRoundTrip(t *testing.T) {
	orig := `a&b<c>d"e'f`
	if got := unescapeXML(safeText(orig)); got != orig {
		t.Errorf("往返失败: escape->unescape = %q，期望 %q", got, orig)
	}
}

// ---------- 回归 #8：跨文件 CopySheet 保留目标原有工作表 ----------

func TestCopySheetToPreservesDstSheets(t *testing.T) {
	dir := t.TempDir()

	// 源工作簿：Create 自带 Sheet1，重命名为 Data，再加一张 ToDrop
	srcPath := filepath.Join(dir, "src.xlsx")
	src, _ := Create()
	if err := src.RenameSheet("Sheet1", "Data"); err != nil {
		t.Fatal(err)
	}
	sws, err := src.Sheet("Data")
	if err != nil {
		t.Fatal(err)
	}
	sws.SetCellStr("A1", "copied-data")
	src.AddSheet("ToDrop")
	src.SaveAs(srcPath)

	// 目标工作簿：已有 2 张表（Create 自带 Sheet1，再加 Keep2）
	dstPath := filepath.Join(dir, "dst.xlsx")
	dst, _ := Create()
	dws, err := dst.Sheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	dws.SetCellStr("A1", "keep-1")
	dst.AddSheet("Keep2")
	dst.SaveAs(dstPath)

	if err := CopySheet(srcPath, dstPath, "Data", ""); err != nil {
		t.Fatalf("CopySheet: %v", err)
	}

	list, err := GetSheetList(dstPath)
	if err != nil {
		t.Fatal(err)
	}
	hasKeep1, hasKeep2, hasData := false, false, false
	for _, s := range list {
		switch s {
		case "Sheet1":
			hasKeep1 = true
		case "Keep2":
			hasKeep2 = true
		case "Data", "Data_copy", "Data_1":
			hasData = true
		}
	}
	if !hasKeep1 || !hasKeep2 {
		t.Errorf("目标原有工作表被破坏: %v", list)
	}
	if !hasData {
		t.Errorf("源工作表未复制进目标: %v", list)
	}

	// 原表内容仍在，被复制表内容也对
	if v, _ := GetCell(dstPath, "Sheet1", "A1"); v != "keep-1" {
		t.Errorf("Sheet1!A1 = %q，期望 keep-1", v)
	}
	var copied string
	for _, s := range list {
		if s == "Data" || strings.HasPrefix(s, "Data_") {
			if v, _ := GetCell(dstPath, s, "A1"); v == "copied-data" {
				copied = s
			}
		}
	}
	if copied == "" {
		t.Errorf("复制过来的表内容不正确: %v", list)
	}

	// 源工作簿不应被修改
	srcList, _ := GetSheetList(srcPath)
	if len(srcList) != 2 {
		t.Errorf("源工作簿被修改: %v", srcList)
	}
}

// TestCopySheetToCreatesDstWhenMissing 验证目标文件不存在时仍能创建。
func TestCopySheetToCreatesDstWhenMissing(t *testing.T) {
	dir := t.TempDir()
	srcPath := filepath.Join(dir, "src2.xlsx")
	dstPath := filepath.Join(dir, "new.xlsx")

	src, _ := Create()
	if err := src.RenameSheet("Sheet1", "Data"); err != nil {
		t.Fatal(err)
	}
	ws, err := src.Sheet("Data")
	if err != nil {
		t.Fatal(err)
	}
	ws.SetCellStr("A1", "hello")
	src.SaveAs(srcPath)

	if err := CopySheet(srcPath, dstPath, "Data", "Data"); err != nil {
		t.Fatalf("CopySheet 到不存在的目标失败: %v", err)
	}
	if _, err := os.Stat(dstPath); err != nil {
		t.Fatalf("目标文件未创建: %v", err)
	}
	// 目标为新建空簿（自带 Sheet1），newName 指定的表名应原样落地
	if v, _ := GetCell(dstPath, "Data", "A1"); v != "hello" {
		t.Errorf("Data!A1 = %q，期望 hello", v)
	}
	list, _ := GetSheetList(dstPath)
	if len(list) != 2 {
		t.Errorf("目标工作表数 = %d（%v），期望 2（空簿 Sheet1 + 复制表）", len(list), list)
	}
}

// ---------- 回归 #9：GetTime 日期格式识别更稳健 ----------

func TestGetTimeDetectsCustomDateFormat(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "date.xlsx")
	b, _ := Create()
	ws, _ := b.Sheet("Sheet1")

	// 46000 表示 2025-11-04 左右
	if err := ws.SetCellNumeric("A1", 46000); err != nil {
		t.Fatal(err)
	}
	// 自定义中文日期格式：字面量包裹的 y/m/d，旧实现会误判为非日期
	if _, err := ws.SetStyle("A1", Style{NumFmt: `yyyy"年"m"月"d"日"`}); err != nil {
		t.Fatal(err)
	}
	if err := ws.SetCellNumeric("A2", 1); err != nil {
		t.Fatal(err)
	}
	if _, err := ws.SetStyle("A2", Style{NumFmt: "0.00"}); err != nil {
		t.Fatal(err)
	}
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}

	b2, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	ws2, _ := b2.Sheet("Sheet1")

	// 自定义中文日期格式必须被识别为日期
	tm, ok := ws2.Cell("A1").GetTime()
	if !ok {
		t.Errorf("自定义日期格式 %q 未被识别为日期", `yyyy"年"m"月"d"日"`)
	} else {
		if tm.Year() < 2020 || tm.Year() > 2035 {
			t.Errorf("GetTime 年份 = %d，序列值 46000 应落在 2020~2035", tm.Year())
		}
	}
	// 纯数字格式不得被误判为日期
	if _, ok := ws2.Cell("A2").GetTime(); ok {
		t.Error("数字格式 0.00 不应被识别为日期")
	}
}

// TestIsDateFormat 验证日期格式判定：内置 numFmtId + 自定义格式代码文本。
func TestIsDateFormat(t *testing.T) {
	cases := []struct {
		id   int
		code string
		want bool
	}{
		{14, "", true},  // 内置 m/d/yy
		{22, "", true},  // 内置 m/d/yy h:mm
		{45, "", true},  // 内置 mm:ss
		{0, "", false},  // General
		{10, "", false}, // 0.00%
		{0, "0.00", false},
		{0, `yyyy"年"m"月"d"日"`, true}, // 自定义中文日期
		{0, "hh:mm:ss", true},
		{0, `[$-804]年;m"月"d"日"`, true},
		{0, `0.0"m"`, false}, // 字面量 m（米）不得误判
		{0, `#,##0.00 "元"`, false},
	}
	for _, c := range cases {
		if got := isDateFormat(c.id, c.code); got != c.want {
			t.Errorf("isDateFormat(%d, %q) = %v，期望 %v", c.id, c.code, got, c.want)
		}
	}
}

// ---------- 回归 #10：行/列插入删除平移浮动图片锚点 ----------

// buildDrawingWithAnchor 构造一个含单个 twoCellAnchor 的 drawing part（0 基锚点）。
const anchorTemplate = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">` +
	`<xdr:twoCellAnchor>` +
	`<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
	`<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
	`<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Picture 1"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill/></xdr:pic>` +
	`<xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`

func TestShiftDrawingAnchorsOnRowInsert(t *testing.T) {
	// from=(col1,row1) to=(col3,row4) 全部 0 基
	xml := fmt.Sprintf(anchorTemplate, 1, 1, 3, 4)
	// 在第 2 行前插入 3 行 → 0 基 row >= 1 的 +3
	out := shiftDrawingAnchors(xml, 2, 3, nil, true)
	// from row: 1 -> 4 ; to row: 4 -> 7 ; col 不变
	if !strings.Contains(out, "<xdr:row>4</xdr:row>") {
		t.Errorf("from.row 未平移（应为 4）:\n%s", out)
	}
	if !strings.Contains(out, "<xdr:row>7</xdr:row>") {
		t.Errorf("to.row 未平移（应为 7）:\n%s", out)
	}
	if !strings.Contains(out, "<xdr:col>1</xdr:col>") || !strings.Contains(out, "<xdr:col>3</xdr:col>") {
		t.Errorf("行操作不应影响列锚点:\n%s", out)
	}
}

func TestShiftDrawingAnchorsOnColInsert(t *testing.T) {
	xml := fmt.Sprintf(anchorTemplate, 1, 1, 3, 4)
	out := shiftDrawingAnchors(xml, 2, 2, nil, false)
	if !strings.Contains(out, "<xdr:col>3</xdr:col>") {
		t.Errorf("from.col 未平移（应为 3）:\n%s", out)
	}
	if !strings.Contains(out, "<xdr:col>5</xdr:col>") {
		t.Errorf("to.col 未平移（应为 5）:\n%s", out)
	}
	if !strings.Contains(out, "<xdr:row>1</xdr:row>") || !strings.Contains(out, "<xdr:row>4</xdr:row>") {
		t.Errorf("列操作不应影响行锚点:\n%s", out)
	}
}

func TestShiftDrawingAnchorsRemovesDeletedBand(t *testing.T) {
	xml := fmt.Sprintf(anchorTemplate, 1, 1, 3, 4)
	// 删除第 2~3 行（1 基）→ 0 基 row 1~2；from.row=1 落入区间 → 整个锚点移除
	out := shiftDrawingAnchors(xml, 4, -2, &[2]int{2, 3}, true)
	if strings.Contains(out, "twoCellAnchor") {
		t.Errorf("锚点落入被删行区间应被移除:\n%s", out)
	}

	xml2 := fmt.Sprintf(anchorTemplate, 1, 5, 3, 8)
	out2 := shiftDrawingAnchors(xml2, 4, -2, &[2]int{2, 3}, true)
	if !strings.Contains(out2, "<xdr:row>3</xdr:row>") {
		t.Errorf("被删区间之后的锚点应前移 2 行（from.row 5→3）:\n%s", out2)
	}
	if !strings.Contains(out2, "<xdr:row>6</xdr:row>") {
		t.Errorf("to.row 应前移 2 行（8→6）:\n%s", out2)
	}
}

// TestInsertRowsShiftsDrawingInWorkbook 端到端：在真实工作簿上插入行后，
// drawing 锚点应同步平移。
func TestInsertRowsShiftsDrawingInWorkbook(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "draw.xlsx")
	b, _ := Create()
	ws, _ := b.Sheet("Sheet1")
	ws.SetCellStr("A1", "x")
	b.SaveAs(p)

	// 手工植入 drawing part + 关联（sheet1.xml.rels + <drawing> 引用）
	bb, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	const drawingPart = "xl/drawings/drawing1.xml"
	bb.fileMap[drawingPart] = []byte(fmt.Sprintf(anchorTemplate, 1, 1, 3, 4))
	// sheet rels
	sheetRels := `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rIdDrw1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`
	bb.fileMap["xl/worksheets/_rels/sheet1.xml.rels"] = []byte(sheetRels)
	// Content_Types
	bb.fileMap["[Content_Types].xml"] = ensureDrawingContentType(bb.fileMap["[Content_Types].xml"], drawingPart)
	// sheet1.xml 增加 <drawing r:id="rIdDrw1"/>
	bb.fileMap["xl/worksheets/sheet1.xml"] = injectDrawingRef(bb.fileMap["xl/worksheets/sheet1.xml"], "rIdDrw1")
	if err := bb.Save(); err != nil {
		t.Fatal(err)
	}

	if err := InsertRows(p, "Sheet1", 2, 3); err != nil {
		t.Fatalf("InsertRows: %v", err)
	}
	dw := readZipPart(t, p, drawingPart)
	if !strings.Contains(dw, "<xdr:row>4</xdr:row>") {
		t.Errorf("插入 3 行后 from.row 应为 4:\n%s", dw)
	}
	if !strings.Contains(dw, "<xdr:row>7</xdr:row>") {
		t.Errorf("插入 3 行后 to.row 应为 7:\n%s", dw)
	}
}

// ---------- 小工具 ----------

// ensureDrawingContentType 为 drawing 部件补齐 [Content_Types].xml 中的 Override。
func ensureDrawingContentType(ct []byte, part string) []byte {
	s := string(ct)
	if strings.Contains(s, "/"+part+`"`) {
		return ct
	}
	override := `<Override PartName="/` + part + `" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`
	if i := strings.Index(s, "</Types>"); i != -1 {
		return []byte(s[:i] + override + s[i:])
	}
	return ct
}

// injectDrawingRef 在 worksheet XML 的 </worksheet> 前插入 <drawing r:id="..."/>。
func injectDrawingRef(wsXML []byte, rid string) []byte {
	s := string(wsXML)
	tag := `<drawing r:id="` + rid + `"/>`
	if i := strings.Index(s, "</worksheet>"); i != -1 {
		return []byte(s[:i] + tag + s[i:])
	}
	return wsXML
}

// ---------- 跨文件 CopySheet 语义：只搬「这一张表」 ----------

// TestCopySheetToMovesOnlyRequestedSheet 验证跨文件复制只搬运被指定的那一张表，
// 源工作簿的其它工作表绝不跟着进入目标工作簿（CopySheet 不是整簿合并）。
func TestCopySheetToMovesOnlyRequestedSheet(t *testing.T) {
	dir := t.TempDir()
	srcPath := filepath.Join(dir, "src.xlsx")
	dstPath := filepath.Join(dir, "dst.xlsx")

	// 源：3 张表
	src, _ := Create()
	if err := src.RenameSheet("Sheet1", "Wanted"); err != nil {
		t.Fatal(err)
	}
	w, _ := src.Sheet("Wanted")
	w.SetCellStr("A1", "want")
	src.AddSheet("Other1")
	src.AddSheet("Other2")
	src.SaveAs(srcPath)

	// 目标：2 张表
	dst, _ := Create()
	d, _ := dst.Sheet("Sheet1")
	d.SetCellStr("A1", "dst-keep")
	dst.AddSheet("DstTwo")
	dst.SaveAs(dstPath)

	if err := CopySheet(srcPath, dstPath, "Wanted", ""); err != nil {
		t.Fatalf("CopySheet: %v", err)
	}

	list, err := GetSheetList(dstPath)
	if err != nil {
		t.Fatal(err)
	}
	// 目标原有 2 张表必须都在
	if !containsStr(list, "Sheet1") || !containsStr(list, "DstTwo") {
		t.Errorf("目标原有工作表丢失: %v", list)
	}
	// 只多出 Wanted 一张；Other1/Other2 绝不能出现
	added := 0
	for _, s := range list {
		if s == "Other1" || s == "Other2" {
			t.Errorf("源工作簿的其它表 %q 被错误复制过来: %v", s, list)
		}
	}
	for _, s := range list {
		if s == "Wanted" || s == "Wanted_copy" {
			added++
		}
	}
	if added != 1 {
		t.Errorf("应恰好新增 1 张表，实际新增 %d 张: %v", added, list)
	}
	if len(list) != 3 {
		t.Errorf("目标工作表总数 = %d，期望 3（原有 2 + 新增 1）: %v", len(list), list)
	}

	// 源工作簿不得被改动
	srcList, _ := GetSheetList(srcPath)
	if len(srcList) != 3 {
		t.Errorf("源工作簿被改动: %v", srcList)
	}
}

// TestCopySheetToSuffixControlsName 验证复制后的表名由位置参数 newName / WithSuffix 决定，
// 即 CopySheet(src, dst, sheet, "表名") 可直接指定，不被 "_copy" 后缀锁死。
func TestCopySheetToSuffixControlsName(t *testing.T) {
	dir := t.TempDir()
	srcPath := filepath.Join(dir, "s.xlsx")
	dstPath := filepath.Join(dir, "d.xlsx")

	src, _ := Create()
	if err := src.RenameSheet("Sheet1", "Data"); err != nil {
		t.Fatal(err)
	}
	src.SaveAs(srcPath)

	dst, _ := Create()
	dst.SaveAs(dstPath)

	// 自定义后缀
	if err := CopySheet(srcPath, dstPath, "Data", "", WithSuffix("_副本")); err != nil {
		t.Fatal(err)
	}
	list, _ := GetSheetList(dstPath)
	if !containsStr(list, "Data_副本") {
		t.Errorf("WithSuffix 未生效，表名列表: %v", list)
	}

	// 直接指定目标表名
	if err := CopySheet(srcPath, dstPath, "Data", "指定名"); err != nil {
		t.Fatalf("WithRename 复制失败: %v", err)
	}
	list2, _ := GetSheetList(dstPath)
	if !containsStr(list2, "指定名") {
		t.Errorf("newName 未生效，表名列表: %v", list2)
	}
}

// containsStr 判断字符串切片是否包含指定元素。
func containsStr(list []string, s string) bool {
	for _, v := range list {
		if v == s {
			return true
		}
	}
	return false
}

// TestCopySheetPositionalNewName 验证 ergonomic 写法：
//
//	excelgo.CopySheet(src, dst, "Sheet1", "汇总表")
//
// 即可直接指定新表名，无需 option；空串则沿用默认后缀。
func TestCopySheetPositionalNewName(t *testing.T) {
	dir := t.TempDir()
	srcPath := filepath.Join(dir, "s.xlsx")
	dstPath := filepath.Join(dir, "d.xlsx")

	src, _ := Create()
	src.SaveAs(srcPath) // 默认表名 Sheet1
	dst, _ := Create()
	dst.SaveAs(dstPath)

	// 你要的那种写法：第 4 个位置参数直接给新表名
	if err := CopySheet(srcPath, dstPath, "Sheet1", "汇总表"); err != nil {
		t.Fatalf("CopySheet 指定新表名失败: %v", err)
	}
	list, _ := GetSheetList(dstPath)
	if !containsStr(list, "汇总表") {
		t.Errorf("未按位置参数命名，表名列表: %v", list)
	}

	// 空串 → 沿用默认后缀
	if err := CopySheet(srcPath, dstPath, "Sheet1", ""); err != nil {
		t.Fatalf("CopySheet 空名失败: %v", err)
	}
	list2, _ := GetSheetList(dstPath)
	if !containsStr(list2, "Sheet1_copy") {
		t.Errorf("空串未沿用默认后缀，表名列表: %v", list2)
	}

	// 同工作簿内的方法版同样支持位置参数
	f, err := Open(srcPath)
	if err != nil {
		t.Fatal(err)
	}
	got, err := f.CopySheet("Sheet1", "簿内副本")
	if err != nil {
		t.Fatalf("CopySheet 方法指定新表名失败: %v", err)
	}
	if got != "簿内副本" {
		t.Errorf("返回值 = %q，期望 簿内副本", got)
	}
}

// TestCopySheetToRefusesToOverwriteUnreadableTarget 验证目标文件存在但无法解析时
// （损坏 / 加密 / 非 xlsx / 路径写错），CopySheet 报错且**不覆盖原文件**。
// 否则用户指向一个错误路径就会静默丢掉原数据。
func TestCopySheetToRefusesToOverwriteUnreadableTarget(t *testing.T) {
	dir := t.TempDir()
	srcPath := filepath.Join(dir, "src.xlsx")

	src, _ := Create()
	if err := src.RenameSheet("Sheet1", "S"); err != nil {
		t.Fatal(err)
	}
	ws, _ := src.Sheet("S")
	ws.SetCellStr("A1", "orig")
	src.SaveAs(srcPath)

	// 目标存在但内容不是合法 xlsx
	badPath := filepath.Join(dir, "corrupt.xlsx")
	garbage := []byte("this is definitely not a zip archive")
	if err := os.WriteFile(badPath, garbage, 0o644); err != nil {
		t.Fatal(err)
	}

	err := CopySheet(srcPath, badPath, "S", "x")
	if err == nil {
		t.Error("目标文件存在但无法解析时应返回错误，不得静默覆盖")
	} else {
		t.Logf("如期报错: %v", err)
	}

	// 原文件必须字节不变
	after, rerr := os.ReadFile(badPath)
	if rerr != nil {
		t.Fatalf("读取目标文件失败: %v", rerr)
	}
	if string(after) != string(garbage) {
		t.Errorf("目标文件被改写！期望保持 %q，实际 %q", string(garbage), string(after))
	}
}

// TestPrintTitlesCarriedOnCopy 验证重复打印行（_xlnm.Print_Titles）与打印区域
// 一起被复制到新表，且表名与 localSheetId 正确重映射。
// 覆盖三条路径：同工作簿 CopySheet、跨文件 CopySheetTo、MergeWorkbook。
func TestPrintTitlesCarriedOnCopy(t *testing.T) {
	dir := t.TempDir()

	newSrc := func(name string) string {
		p := filepath.Join(dir, name)
		b, _ := Create()
		if err := b.RenameSheet("Sheet1", "Src"); err != nil {
			t.Fatal(err)
		}
		ws, _ := b.Sheet("Src")
		ws.SetCellStr("A1", "x")
		if err := b.SaveAs(p); err != nil {
			t.Fatal(err)
		}
		if err := SetSheetProps(p, "Src", SheetProps{PrintArea: "A1:D20", PrintTitleRows: "1:1"}); err != nil {
			t.Fatal(err)
		}
		return p
	}
	// 断言目标簿中某个表同时保有打印区域与重复打印行
	assertTitles := func(tag, p, sheet string) {
		t.Helper()
		props, err := GetProps(p, sheet)
		if err != nil {
			t.Fatalf("[%s] GetProps(%s) 失败: %v", tag, sheet, err)
		}
		if props.PrintArea != "$A$1:$D$20" {
			t.Errorf("[%s] %s 的 PrintArea = %q，期望 $A$1:$D$20", tag, sheet, props.PrintArea)
		}
		if props.PrintTitleRows != "$1:$1" {
			t.Errorf("[%s] %s 的 PrintTitleRows = %q，期望 $1:$1（重复打印行未随复制搬运）", tag, sheet, props.PrintTitleRows)
		}
	}

	t.Run("同工作簿 CopySheet", func(t *testing.T) {
		p := newSrc("in.xlsx")
		b, err := Open(p)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := b.CopySheet("Src", "副本"); err != nil {
			t.Fatal(err)
		}
		if err := b.Save(); err != nil {
			t.Fatal(err)
		}
		assertTitles("同簿", p, "Src")
		assertTitles("同簿", p, "副本")
	})

	t.Run("跨文件 CopySheetTo", func(t *testing.T) {
		src := newSrc("x1.xlsx")
		dst := filepath.Join(dir, "x2.xlsx")
		d, _ := Create()
		if err := d.SaveAs(dst); err != nil {
			t.Fatal(err)
		}
		if err := CopySheet(src, dst, "Src", "搬入"); err != nil {
			t.Fatal(err)
		}
		assertTitles("跨文件", dst, "搬入")
	})

	t.Run("MergeWorkbook", func(t *testing.T) {
		src := newSrc("m1.xlsx")
		dst := filepath.Join(dir, "m2.xlsx")
		d, _ := Create()
		if err := d.SaveAs(dst); err != nil {
			t.Fatal(err)
		}
		if err := MergeWorkbook(dst, []SourceRef{{Workbook: src, Sheet: "Src"}}); err != nil {
			t.Fatal(err)
		}
		list, _ := GetSheetList(dst)
		assertTitles("合并", dst, list[len(list)-1])
	})
}

// TestRowColBoundsValidation 验证行列插入/删除的表尾边界校验：
//   - 合法操作（含在表尾+1 处追加，符合 Excel 行为）必须通过；
//   - 超出表尾的操作必须返回错误，而不是静默生成大片空行/空列。
func TestRowColBoundsValidation(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "b.xlsx")

	// 造 3 行 x 3 列的表，并设置列宽（产生 <col min/max> 定义）
	b, _ := Create()
	ws, _ := b.Sheet("Sheet1")
	for r := 1; r <= 3; r++ {
		for c := 1; c <= 3; c++ {
			cell := string(rune('A'+c-1)) + strconv.Itoa(r)
			ws.SetCellStr(cell, "v")
		}
	}
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}
	if err := SetColWidth(p, "Sheet1", 1, 15); err != nil {
		t.Fatal(err)
	}

	t.Run("合法插入", func(t *testing.T) {
		for _, row := range []int{1, 2, 4} { // 4 == maxRow+1，合法追加
			if err := InsertRows(p, "Sheet1", row, 1); err != nil {
				t.Errorf("InsertRows(row=%d) 应合法，却报错: %v", row, err)
			}
		}
	})
	t.Run("越界插入被拒", func(t *testing.T) {
		if err := InsertRows(p, "Sheet1", 99, 1); err == nil {
			t.Error("InsertRows(row=99) 超出表尾应报错")
		}
		if err := InsertCols(p, "Sheet1", 99, 1); err == nil {
			t.Error("InsertCols(col=99) 超出表尾应报错")
		}
	})
	t.Run("合法删除", func(t *testing.T) {
		if err := RemoveRows(p, "Sheet1", 1, 1); err != nil {
			t.Errorf("RemoveRows(1,1) 应合法，却报错: %v", err)
		}
		if err := RemoveCols(p, "Sheet1", 1, 1); err != nil {
			t.Errorf("RemoveCols(1,1) 应合法，却报错: %v", err)
		}
	})
	t.Run("越界删除被拒", func(t *testing.T) {
		if err := RemoveRows(p, "Sheet1", 10, 5); err == nil {
			t.Error("RemoveRows(10,5) 超出表尾应报错")
		}
		if err := RemoveCols(p, "Sheet1", 90, 5); err == nil {
			t.Error("RemoveCols(90,5) 超出表尾应报错")
		}
	})
	t.Run("对象式 API 同样校验", func(t *testing.T) {
		f, err := Open(p)
		if err != nil {
			t.Fatal(err)
		}
		s, err := f.Sheet("Sheet1")
		if err != nil {
			t.Fatal(err)
		}
		if err := s.InsertRows(99, 1); err == nil {
			t.Error("(*WorkSheet).InsertRows(99,1) 超出表尾应报错")
		}
		if err := s.InsertCols(99, 1); err == nil {
			t.Error("(*WorkSheet).InsertCols(99,1) 超出表尾应报错")
		}
		if err := s.RemoveRows(500, 2); err == nil {
			t.Error("(*WorkSheet).RemoveRows(500,2) 超出表尾应报错")
		}
		if err := s.RemoveCols(500, 2); err == nil {
			t.Error("(*WorkSheet).RemoveCols(500,2) 超出表尾应报错")
		}
	})
}

// TestSpecialCharSheetNames 验证表名含 XML 特殊字符（& < > "）时，
// 复制、重命名、合并均正确转义/解码，不破坏 workbook.xml 结构。
//
// 回归：未转义时裸 & 会打断 XML，GetSheetList 返回空 —— 所有工作表“凭空消失”。
func TestSpecialCharSheetNames(t *testing.T) {
	dir := t.TempDir()
	const srcName = `A&B<C>"D"`

	p := filepath.Join(dir, "sp.xlsx")
	b, _ := Create()
	if err := b.RenameSheet("Sheet1", srcName); err != nil {
		t.Fatal(err)
	}
	ws, _ := b.Sheet(srcName)
	ws.SetCellStr("A1", "x")
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}
	if err := SetSheetProps(p, srcName, SheetProps{PrintArea: "A1:D20", PrintTitleRows: "1:1"}); err != nil {
		t.Fatal(err)
	}

	t.Run("同簿复制", func(t *testing.T) {
		const newName = `新&表<1>`
		if err := CopySheet(p, p, srcName, newName); err != nil {
			t.Fatal(err)
		}
		list, err := GetSheetList(p)
		if err != nil {
			t.Fatalf("GetSheetList 失败: %v", err)
		}
		if len(list) != 2 {
			t.Fatalf("复制后应仍有 2 张表，实际 %d 张: %q（XML 很可能已损坏）", len(list), list)
		}
		if !containsStr(list, srcName) || !containsStr(list, newName) {
			t.Errorf("表名丢失或被破坏: %q", list)
		}
		// 打印属性须为正常值（不能是转义态，也不能双重转义）
		props, err := GetProps(p, newName)
		if err != nil {
			t.Fatalf("GetProps 失败: %v", err)
		}
		if props.PrintArea != "$A$1:$D$20" {
			t.Errorf("PrintArea = %q，期望 $A$1:$D$20（不应含转义实体）", props.PrintArea)
		}
		if props.PrintTitleRows != "$1:$1" {
			t.Errorf("PrintTitleRows = %q，期望 $1:$1", props.PrintTitleRows)
		}
	})

	t.Run("按含特殊字符的原名重命名", func(t *testing.T) {
		p2 := filepath.Join(dir, "rn.xlsx")
		c, _ := Create()
		if err := c.RenameSheet("Sheet1", srcName); err != nil {
			t.Fatal(err)
		}
		if err := c.SaveAs(p2); err != nil {
			t.Fatal(err)
		}
		// 回归：正则拿未转义的 srcName 匹配转义后的 A&amp;B，会静默不生效
		if err := RenameSheet(p2, srcName, "改名后"); err != nil {
			t.Fatalf("RenameSheet 失败: %v", err)
		}
		list, _ := GetSheetList(p2)
		if !containsStr(list, "改名后") {
			t.Errorf("重命名未生效: %q", list)
		}
		if containsStr(list, srcName) {
			t.Errorf("旧表名仍存在: %q", list)
		}
	})

	t.Run("跨文件复制与合并不破坏结构", func(t *testing.T) {
		p3 := filepath.Join(dir, "dst.xlsx")
		d, _ := Create()
		if err := d.SaveAs(p3); err != nil {
			t.Fatal(err)
		}
		if err := CopySheet(p, p3, srcName, "跨&文件"); err != nil {
			t.Fatal(err)
		}
		if err := MergeWorkbook(p3, []SourceRef{{Workbook: p, Sheet: srcName}}); err != nil {
			t.Fatal(err)
		}
		list, err := GetSheetList(p3)
		if err != nil {
			t.Fatalf("GetSheetList 失败: %v", err)
		}
		if len(list) != 3 { // Sheet1 + 跨&文件 + 合并进来的 A&B<C>"D"
			t.Errorf("目标工作簿应有 3 张表，实际 %d: %q", len(list), list)
		}
		for _, n := range list {
			props, err := GetProps(p3, n)
			if err != nil {
				t.Errorf("GetProps(%q) 失败: %v", n, err)
				continue
			}
			if n == "Sheet1" {
				continue
			}
			if strings.Contains(props.PrintArea, "&amp;") || strings.Contains(props.PrintArea, "&lt;") {
				t.Errorf("%s 的 PrintArea 含转义实体（双重转义）: %q", n, props.PrintArea)
			}
		}
	})
}

// TestSheetNameValidation 验证工作表名校验：Excel 限制表名最长 31 字符且不允许
// \ / ? * [ ] : ，超长/非法名会产出 Excel 拒绝打开的文件，须提前拦截。
func TestSheetNameValidation(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "n.xlsx")
	b, _ := Create()
	b.SaveAs(p)

	tooLong := strings.Repeat("长", 32)
	exactly31 := strings.Repeat("长", 31)

	cases := []struct {
		desc    string
		name    string
		wantErr bool
	}{
		{"31 字符（边界，应通过）", exactly31, false},
		{"32 字符（超长）", tooLong, true},
		{"含冒号", "a:b", true},
		{"含反斜杠", `a\b`, true},
		{"含斜杠", "a/b", true},
		{"含问号", "a?b", true},
		{"含星号", "a*b", true},
		{"含方括号", "a[1]", true},
		{"空名", "", true},
		{"正常中文名", "汇总表①", false},
		{"含 & 等需转义字符（合法）", `A&B<C>"D"`, false},
	}
	for _, c := range cases {
		err := validateSheetName(c.name)
		if c.wantErr && err == nil {
			t.Errorf("%s：%q 应报错，实际通过", c.desc, c.name)
		}
		if !c.wantErr && err != nil {
			t.Errorf("%s：%q 不应报错，实际 %v", c.desc, c.name, err)
		}
	}

	// 实际入口：AddSheet / NewSheet / RenameSheet / CopySheet 都应拦截
	if _, err := Create(); err == nil {
		bb, _ := Create()
		if _, err := bb.AddSheet(tooLong); err == nil {
			t.Error("AddSheet 超长名应报错")
		}
		if _, err := bb.AddSheet("a:b"); err == nil {
			t.Error("AddSheet 含冒号名应报错")
		}
	}
	if err := RenameSheet(p, "Sheet1", tooLong); err == nil {
		t.Error("RenameSheet 超长名应报错")
	}
	if err := CopySheet(p, p, "Sheet1", tooLong); err == nil {
		t.Error("CopySheet 超长 newName 应报错")
	}
	// 原表名超长时，追加默认后缀后仍超长也应报错
	if _, err := NewSheet(p, tooLong); err == nil {
		t.Error("NewSheet 超长名应报错")
	}
}

// TestRenameSheetClashNoInfiniteLoop 验证重命名遇到同名时按序号避让（_1/_2/...），
// 而不是固定加 "_ren" 后缀 —— 后者在「原名与新名互相占用」时会死循环。
func TestRenameSheetClashNoInfiniteLoop(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "r.xlsx")
	b, _ := Create()
	b.AddSheet("A")
	b.AddSheet("B")
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}

	// A 改名为 B：B 已存在，应避让且不得死循环
	done := make(chan error, 1)
	go func() { done <- RenameSheet(p, "A", "B") }()
	select {
	case err := <-done:
		if err != nil {
			t.Fatalf("RenameSheet 返回错误: %v", err)
		}
	case <-time.After(5 * time.Second):
		t.Fatal("RenameSheet 死循环！")
	}
	list, _ := GetSheetList(p)
	if !containsStr(list, "B_1") {
		t.Errorf("同名避让结果应含 B_1，实际 %q", list)
	}
}

// TestXMLInjectionResistance 验证「XML 注入」防护。
//
// 本库全程以字符串/正则编辑 XML（为保留 WPS 内嵌图片与版式），不做 XML 解析器校验，
// 因此每个写入点的转义就是唯一防线。测试用多种载荷写入各类字段，
// 断言：产出的所有 XML 部件均良构（可用严格解析器打开），且未注入多余节点。
func TestXMLInjectionResistance(t *testing.T) {
	payloads := map[string]string{
		"闭合标签":     `X"></comment><evil data="1"/><comment ref="A1`,
		"闭合加实体":    `X</t></text></comment><evil/>&amp;&lt;`,
		"属性引号逃逸":   `X" onLoad="alert(1)" data="`,
		"CDATA 伪装": `<![CDATA[</evil>]]>`,
		"注释伪装":     `X<!-- </evil> -->`,
		"控制字符":     "A\x00B\x08C\x0BD",
	}

	for name, payload := range payloads {
		t.Run(name, func(t *testing.T) {
			dir := t.TempDir()
			p := filepath.Join(dir, "inj.xlsx")

			// 把同一个载荷灌入所有可写入用户数据的字段
			f, _ := Create()
			ws, _ := f.Sheet("Sheet1")
			ws.SetCellStr("A1", "base")
			_ = ws.AddComment("B1", payload, "auth")         // 批注文本 + 作者
			_ = ws.AddComment("B2", "text", payload)         // 批注作者
			_ = ws.AddHyperlink("C1", payload, "disp")       // 超链接 URL
			_ = ws.AddHyperlink("C2", "https://ok", payload) // 超链接显示文本
			_ = ws.SetCellStr("D1", payload)                 // 共享字符串
			_ = ws.SetCellFormula("D2", payload, "0")        // 公式
			_ = f.SetDefinedName(payload, "Sheet1!$A$1")     // 命名区域名
			_ = f.SaveAs(p)

			// 走包级 API 的字段
			_ = SetDocProps(p, map[string]string{"title": payload, "author": payload})
			_ = AddTable(p, "Sheet1", "A5:C7", payload)
			_ = AddDataValidation(p, "Sheet1", "E1", "list", "equal", payload, "", true)
			_ = SetConditionalFormat(p, "Sheet1", "F1:F5", "cell", payload, 1, Style{NumFmt: payload})
			_ = SetSheetProps(p, "Sheet1", SheetProps{
				TabColor:     payload,
				HeaderFooter: &HeaderFooter{OddHeader: payload, OddFooter: payload},
			})
			_ = AddComment(p, "Sheet1", "G1", payload, payload)

			// 断言 1：所有 XML 部件良构（严格解析，非法 XML 会直接报错）
			assertAllXMLPartsWellFormed(t, p)

			// 断言 2：未注入 <evil 之类的多余节点
			assertNoInjectedNode(t, p, "evil")
		})
	}
}

// assertAllXMLPartsWellFormed 逐个 XML/rels 部件做严格解析。
func assertAllXMLPartsWellFormed(t *testing.T, p string) {
	t.Helper()
	r, err := zip.OpenReader(p)
	if err != nil {
		t.Fatalf("打开 zip 失败: %v", err)
	}
	defer r.Close()
	for _, f := range r.File {
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		data := readZipPart(t, p, f.Name)
		if err := xmlWellFormed(data); err != nil {
			t.Errorf("部件 %s 不是良构 XML: %v", f.Name, err)
		}
	}
}

// assertNoInjectedNode 检查所有 XML 部件中是否出现了注入的节点名。
func assertNoInjectedNode(t *testing.T, p, marker string) {
	t.Helper()
	r, err := zip.OpenReader(p)
	if err != nil {
		t.Fatalf("打开 zip 失败: %v", err)
	}
	defer r.Close()
	for _, f := range r.File {
		if !strings.HasSuffix(f.Name, ".xml") && !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		data := readZipPart(t, p, f.Name)
		if strings.Contains(data, "<"+marker) {
			t.Errorf("部件 %s 出现注入节点 <%s ...>", f.Name, marker)
		}
	}
}

// xmlWellFormed 用 encoding/xml 的严格解码器校验 XML 良构性。
// 该解码器遵循 XML 规范：非法控制字符、未闭合标签、属性未转义等都会直接返回错误。
func xmlWellFormed(data string) error {
	dec := xml.NewDecoder(strings.NewReader(data))
	for {
		_, err := dec.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}
	}
}

// TestColWidthRangeInterop 验证 SetColWidthRange 写出的列宽能被其它工具
// 完整读出（互操作性）。
//
// 回归：本库原先写 `<col min="4" max="6" width="12"/>`。这在 OOXML 规范上
// 合法，但消费方常把 column_dimensions 当稀疏字典处理 —— 实测 openpyxl
// 只为显式出现的列建条目，于是同一份文件 excelgo 自己读出 3 列都是 12，
// openpyxl 只读出 D 列，E/F 丢失。
// openpyxl 自身即使宽度相同也逐列展开（min=max），说明这是被广泛验证的
// 实现所选择的保守形式，故本库改为逐列写出。
func TestColWidthRangeInterop(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "cw.xlsx")
	b, _ := Create()
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}
	if err := SetColWidthRange(p, "Sheet1", 4, 6, 12); err != nil {
		t.Fatal(err)
	}

	// 1) 本库自己读：三列都应有宽度
	for _, c := range []int{4, 5, 6} {
		w, err := GetColWidth(p, "Sheet1", c)
		if err != nil {
			t.Fatalf("GetColWidth(%d) 失败: %v", c, err)
		}
		if w != 12 {
			t.Errorf("列 %d 宽度 = %v，期望 12", c, w)
		}
	}
	// 区间外不应受影响
	for _, c := range []int{3, 7} {
		if w, _ := GetColWidth(p, "Sheet1", c); w != 0 {
			t.Errorf("区间外列 %d 宽度 = %v，期望 0", c, w)
		}
	}

	// 2) XML 层面：必须是逐列 min=max，不能是合并区间
	//    （合并区间是互操作问题的根源，需锁死）
	data := readZipPart(t, p, "xl/worksheets/sheet1.xml")
	for _, want := range []string{
		`<col min="4" max="4"`, `<col min="5" max="5"`, `<col min="6" max="6"`,
	} {
		if !strings.Contains(data, want) {
			t.Errorf("sheet1.xml 缺少 %s —— 列宽未逐列写出，跨工具会丢列（got: %s）",
				want, extractCols(data))
		}
	}
	if strings.Contains(data, `min="4" max="6"`) {
		t.Errorf("仍存在合并区间 <col min=\"4\" max=\"6\">，openpyxl 等工具会读不到 E/F 列")
	}
}

// extractCols 取出 sheet XML 里的 <cols> 块，供失败信息展示。
func extractCols(sheetXML string) string {
	i := strings.Index(sheetXML, "<cols>")
	j := strings.Index(sheetXML, "</cols>")
	if i == -1 || j == -1 || j < i {
		return "(no <cols>)"
	}
	return sheetXML[i : j+len("</cols>")]
}

// TestTableReidentifyOnCopy 验证复制/合并工作表时表格部件被重新标识。
//
// 回归：部件搬运是字节级的（原样复制），源表与副本的 table 部件会带着
// 相同的 id 与 name/displayName。而这两个属性都是**工作簿级唯一**的：
//   - id 重复：消费方按 id 索引表，冲突即错乱；
//   - name 重复：openpyxl 直接抛 "Table with name X already exists"，
//     Excel 同样判定文件损坏。
//
// 症状是"复制一个带表格的工作表后，整个文件打不开"。
func TestTableReidentifyOnCopy(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "tbl.xlsx")
	b, _ := Create()
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}
	mustNoErr(t, RenameSheet(p, "Sheet1", "Src"))
	mustNoErr(t, SetCellStr(p, "Src", "E1", "列1"))
	mustNoErr(t, SetCellStr(p, "Src", "F1", "列2"))
	mustNoErr(t, SetCellStr(p, "Src", "E2", "a"))
	mustNoErr(t, SetCellStr(p, "Src", "F2", "b"))
	mustNoErr(t, AddTable(p, "Src", "E1:F2", "表A"))

	g, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := g.CopySheet("Src", "副本"); err != nil {
		t.Fatal(err)
	}
	if err := g.Save(); err != nil {
		t.Fatal(err)
	}

	// 收集所有 table 部件的 id 与 displayName，断言无重复
	type tblInfo struct{ part, id, name string }
	var infos []tblInfo
	for _, name := range zipPartNamesOf(p) {
		if !strings.HasPrefix(name, "xl/tables/") || !strings.HasSuffix(name, ".xml") {
			continue
		}
		data := readZipPart(t, p, name)
		infos = append(infos, tblInfo{
			part: name,
			id:   xmlAttrOf(data, "id"),
			name: xmlAttrOf(data, "displayName"),
		})
	}
	if len(infos) != 2 {
		t.Fatalf("应有 2 个表格部件，实际 %d: %+v", len(infos), infos)
	}
	seenID := map[string]string{}
	seenName := map[string]string{}
	for _, it := range infos {
		if prev, dup := seenID[it.id]; dup {
			t.Errorf("表格 id 重复：%s 与 %s 都是 id=%q（id 必须工作簿级唯一）",
				prev, it.part, it.id)
		}
		seenID[it.id] = it.part
		if prev, dup := seenName[it.name]; dup {
			t.Errorf("表格 displayName 重复：%s 与 %s 都是 %q"+
				"（openpyxl 会抛 \"Table with name X already exists\"，Excel 判文件损坏）",
				prev, it.part, it.name)
		}
		seenName[it.name] = it.part
	}
	// name 属性也必须唯一（openpyxl 判定重名用的是 name）
	seenPlain := map[string]bool{}
	for _, name := range zipPartNamesOf(p) {
		if !strings.HasPrefix(name, "xl/tables/") {
			continue
		}
		plain := xmlAttrOf(readZipPart(t, p, name), "name")
		if seenPlain[plain] {
			t.Errorf("表格 name 重复：%q（%s）", plain, name)
		}
		seenPlain[plain] = true
	}
}

// mustNoErr 是回归测试里的错误断言辅助。
func mustNoErr(t *testing.T, err error) {
	t.Helper()
	if err != nil {
		t.Fatalf("操作失败: %v", err)
	}
}

// zipPartNamesOf 列出 zip 全部部件名。
func zipPartNamesOf(p string) []string {
	r, err := zip.OpenReader(p)
	if err != nil {
		return nil
	}
	defer r.Close()
	var out []string
	for _, f := range r.File {
		out = append(out, f.Name)
	}
	return out
}

// xmlAttrOf 取出 XML 里第一个出现的指定属性值。
func xmlAttrOf(s, attr string) string {
	key := attr + "=\""
	i := strings.Index(s, key)
	if i == -1 {
		return ""
	}
	rest := s[i+len(key):]
	j := strings.Index(rest, "\"")
	if j == -1 {
		return ""
	}
	return rest[:j]
}

// TestMultiLevelPartChainPreserved 验证**多级关系链**的部件搬运。
//
// 回归：部件搬运只处理一层 rels。关系链实际是多级的 ——
//
//	sheet -> drawing -> chart
//
// 早期实现把 drawing 的 rels 原样搬完就不管了，chartN.xml 从未被复制，
// 而 rels 仍指向它。结果：悬空关系，Excel/WPS 判定文件损坏，
// openpyxl 报 "There is no item named 'xl/charts/chart1.xml'"。
//
// 这个缺陷只在真实含图表的文件上才暴露 —— 库自身不创建图表，
// 所以必须有外部实现（openpyxl）生成含图表文件来验证。
func TestMultiLevelPartChainPreserved(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "charts.xlsx")
	makeChartWorkbook(t, p)

	// 前置：源文件必须确实有多级链（Data 表 -> drawing -> chart）
	before := chartCount(t, p)
	if before < 3 {
		t.Fatalf("源文件图表数 = %d，测试前提不成立", before)
	}
	wantData := chartCountOfSheet(t, p, "Data")
	if wantData == 0 {
		t.Fatal("源文件 Data 表没有图表，测试前提不成立")
	}

	g, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := g.CopySheet("Data", "副本"); err != nil {
		t.Fatal(err)
	}
	if err := g.Save(); err != nil {
		t.Fatal(err)
	}

	// 1) 被复制表的图表应各拿到独立副本
	if n := chartCount(t, p); n != before+wantData {
		t.Errorf("图表部件数 = %d，期望 %d（源 %d + 副本 %d）",
			n, before+wantData, before, wantData)
	}

	// 2) 关系链必须闭合：drawing 的 rels 指向的 chart 真实存在
	for _, name := range zipPartNamesOf(p) {
		if !strings.HasPrefix(name, "xl/drawings/_rels/") {
			continue
		}
		baseDir := name[:strings.LastIndex(name, "/_rels/")]
		for _, e := range parseRels(readZipPart(t, p, name)) {
			if e.external || !strings.Contains(e.target, "chart") {
				continue
			}
			resolved := resolveRelTarget(baseDir, e.target)
			if !partExistsIn(t, p, resolved) {
				t.Errorf("图表链断裂：%s -> %s（解析为 %s，不存在）—— "+
					"这会让 Excel 判定文件损坏", name, e.target, resolved)
			}
		}
	}

	// 3) 下游部件必须在规范目录下：chart 不能被塞进 xl/drawings/
	for _, name := range zipPartNamesOf(p) {
		if strings.HasPrefix(name, "xl/drawings/chart") {
			t.Errorf("图表部件被放错目录：%s（应在 xl/charts/ 下）", name)
		}
	}

	// 4) 图表的数据引用表达式必须保留
	found := false
	for _, name := range zipPartNamesOf(p) {
		if strings.HasPrefix(name, "xl/charts/chart") && strings.HasSuffix(name, ".xml") {
			if strings.Contains(readZipPart(t, p, name), "<f>") {
				found = true
			}
		}
	}
	if !found {
		t.Error("所有图表都丢失了数据引用 <f>，Excel 打开后图表取不到数据")
	}

	// 5) 关系不得悬空
	assertNoDanglingRelsIn(t, p)
}

// partExistsIn 判断部件是否存在（回归测试内的精简版）。
func partExistsIn(t *testing.T, p, name string) bool {
	t.Helper()
	for _, n := range zipPartNamesOf(p) {
		if n == name {
			return true
		}
	}
	return false
}

// assertNoDanglingRelsIn 断言所有内部关系都能解析到真实部件。
func assertNoDanglingRelsIn(t *testing.T, p string) {
	t.Helper()
	exists := map[string]bool{}
	for _, n := range zipPartNamesOf(p) {
		exists[n] = true
	}
	for _, n := range zipPartNamesOf(p) {
		if !strings.HasSuffix(n, ".rels") {
			continue
		}
		baseDir := relsBaseDir(n)
		for _, e := range parseRels(readZipPart(t, p, n)) {
			if e.external {
				continue
			}
			resolved := resolveRelTarget(baseDir, e.target)
			if resolved != "" && !exists[resolved] {
				t.Errorf("悬空关系：%s -> %s（解析为 %s，不存在）", n, e.target, resolved)
			}
		}
	}
}

// TestConditionalFormatRuleTypes 验证 10 种条件格式规则类型都能正确写入，
// 且被 openpyxl 严格解析。
//
// 回归背景：早期只支持单条 expression/cellIs，且会给**无公式类型**
// （aboveAverage/top10/duplicateValues/uniqueValues）附一个空 <formula></formula>，
// 消费方读出 formula=[”] 并可能误判规则。
func TestConditionalFormatRuleTypes(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "cf.xlsx")
	b, _ := Create()
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}
	for i := 1; i <= 10; i++ {
		mustNoErr(t, SetCellNumeric(p, "Sheet1", "B"+strconv.Itoa(i), float64(i*10)))
	}

	above := true
	rules := []ConditionalFormatRule{
		{Type: "aboveAverage", AboveAverage: &above,
			Style: &Style{Font: &FontStyle{Color: "FFFF0000"}}},
		{Type: "cellIs", Operator: "lessThan", Formula: "50",
			Style: &Style{Fill: &FillStyle{Color: "FFFFFF00", PatternType: "solid"}}},
		{Type: "colorScale", ColorScale: []ColorScalePoint{
			{Type: "min", Color: "FF63BE7B"},
			{Type: "percentile", Value: "50", Color: "FFFFEB84"},
			{Type: "max", Color: "FFF8696B"},
		}},
		{Type: "dataBar", Color: "FF638EC6"},
		{Type: "iconSet", IconSet: "3TrafficLights1"},
		{Type: "top10", Rank: 3, Percent: true, Bottom: true},
		{Type: "expression", Formula: "$B1>50", StopIfTrue: true},
		{Type: "containsText", Text: "5", Operator: "containsText",
			Formula: `NOT(ISERROR(SEARCH("5",B1)))`},
		{Type: "duplicateValues", Style: &Style{Font: &FontStyle{Italic: true}}},
		{Type: "uniqueValues"},
	}
	g, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	sh, _ := g.Sheet("Sheet1")
	if err := sh.SetConditionalFormatRules("B1:B10", rules); err != nil {
		t.Fatalf("SetConditionalFormatRules 失败: %v", err)
	}
	if err := g.Save(); err != nil {
		t.Fatal(err)
	}

	// 1) openpyxl 必须能严格解析（任何非法结构都会在这里暴露）
	snap := runSnapshot(t, p)
	sheets := snap["sheets"].(map[string]interface{})
	s1 := sheets["Sheet1"].(map[string]interface{})
	cfs, _ := s1["conditional_formats"].([]interface{})
	if len(cfs) != 10 {
		t.Fatalf("openpyxl 读回 %d 条规则，期望 10", len(cfs))
	}
	seen := map[string]bool{}
	for _, c := range cfs {
		m := c.(map[string]interface{})
		seen[m["type"].(string)] = true
	}
	for _, r := range rules {
		if !seen[r.Type] {
			t.Errorf("规则类型 %q 未被 openpyxl 读回", r.Type)
		}
	}

	// 2) 无公式类型不得带空 <formula>
	data := readZipPart(t, p, "xl/worksheets/sheet1.xml")
	for _, noFormula := range []string{"aboveAverage", "top10", "duplicateValues", "uniqueValues"} {
		// 这四种规则靠属性与区域求值，不带子元素 —— 必须是自闭合标签。
		// （若附一个空 <formula></formula>，消费方读出 formula=[''] 会误判规则）
		openTag := regexp.MustCompile(`<cfRule type="` + noFormula + `"[^>]*?/?>`).FindString(data)
		if openTag == "" {
			t.Errorf("未找到 %s 规则", noFormula)
			continue
		}
		if !strings.HasSuffix(openTag, "/>") {
			t.Errorf("%s 规则不应有子元素（含 <formula>），实际: %s", noFormula, openTag)
		}
	}

	// 3) 文本类规则必须带 text 属性
	if !strings.Contains(data, `text="5"`) {
		t.Error("containsText 规则缺少 text 属性（Excel 靠它知道匹配什么）")
	}

	// 4) 非法规则必须被提前拒绝
	for _, bad := range []ConditionalFormatRule{
		{Type: "colorScale", ColorScale: []ColorScalePoint{{Type: "min", Color: "FF000000"}}},
		{Type: "colorScale", ColorScale: []ColorScalePoint{
			{Type: "min"}, {Type: "max", Color: "FF000000"}}},
		{Type: "iconSet", IconSet: "NotAnIconSet"},
		{Type: "dataBar"},
		{Type: "expression"},
		{Type: "cellIs", Formula: "1"},
		{Type: "containsText", Text: "x"},
		{Type: "不存在的类型"},
		{Type: "cellIs", Operator: "不存在的运算符", Formula: "1"},
	} {
		if err := validateCFRule(bad); err == nil {
			t.Errorf("非法规则 %+v 应报错却通过了（会产出消费方读不出来的文件）", bad)
		}
	}
}

// TestArrayFormulaAndMoveRange 验证数组公式与区域搬移。
//
// 两者都涉及"批量改写单元格"，历史上在这里出过两个坐标系 bug：
//  1. 内部 0 基与坐标字符串 1 基混用 → 生成 "A0" 这类非法坐标；
//  2. 逐格字符串拼接在"先清空一批再写另一批"时把行号算错（出现 <row r="0">）。
//
// 修复后统一为"内部 0 基、输出 +1"，并改为按行整体重建。
func TestArrayFormulaAndMoveRange(t *testing.T) {
	dir := t.TempDir()

	t.Run("refOf 坐标转换", func(t *testing.T) {
		// 覆盖 26 进制进位边界
		cases := []struct {
			row, col int
			want     string
		}{
			{0, 0, "A1"}, {0, 25, "Z1"}, {0, 26, "AA1"},
			{0, 27, "AB1"}, {0, 51, "AZ1"}, {0, 52, "BA1"},
			{2, 3, "D3"}, {9, 4, "E10"}, {0, 701, "ZZ1"}, {0, 702, "AAA1"},
		}
		for _, c := range cases {
			if got := refOf(c.row, c.col); got != c.want {
				t.Errorf("refOf(%d,%d) = %q，期望 %q", c.row, c.col, got, c.want)
			}
		}
	})

	t.Run("数组公式", func(t *testing.T) {
		p := filepath.Join(dir, "arr.xlsx")
		b, _ := Create()
		mustNoErr(t, b.SaveAs(p))
		ws, _ := b.Sheet("Sheet1")
		mustNoErr(t, ws.SetCellArrayFormula("A1", "A1:C3", "A1:A3*2", ""))
		mustNoErr(t, b.Save())

		data := readZipPart(t, p, "xl/worksheets/sheet1.xml")
		if !strings.Contains(data, `<f t="array"`) {
			t.Errorf("数组公式缺少 t=\"array\"：%s", data)
		}
		if !strings.Contains(data, `ref="A1:C3"`) {
			t.Errorf("数组公式缺少 ref 属性：%s", data)
		}
		// 非法输入必须被拒
		if err := ws.SetCellArrayFormula("A1", "bad ref!!", "X", ""); err == nil {
			t.Error("非法 ref 应报错")
		}
		if err := ws.SetCellArrayFormula("0A", "A1", "X", ""); err == nil {
			t.Error("非法 cell 应报错")
		}
	})

	t.Run("MoveRange 搬移", func(t *testing.T) {
		p := filepath.Join(dir, "mv.xlsx")
		b, _ := Create()
		mustNoErr(t, b.SaveAs(p))
		ws, _ := b.Sheet("Sheet1")
		for i := 1; i <= 3; i++ {
			ws.SetCellStr("A"+strconv.Itoa(i), "文本"+strconv.Itoa(i))
			ws.SetCellNumeric("B"+strconv.Itoa(i), float64(i*10))
		}
		mustNoErr(t, ws.MoveRange("A1:B3", "D1"))
		mustNoErr(t, b.Save())

		// 目标位置拿到全部值（文本 + 数值）
		for i := 1; i <= 3; i++ {
			gotD, _ := GetCell(p, "Sheet1", "D"+strconv.Itoa(i))
			if gotD != "文本"+strconv.Itoa(i) {
				t.Errorf("D%d = %q，期望 文本%d", i, gotD, i)
			}
			gotE, _ := GetCell(p, "Sheet1", "E"+strconv.Itoa(i))
			if gotE != strconv.Itoa(i*10) {
				t.Errorf("E%d = %q，期望 %d", i, gotE, i*10)
			}
			// 源区域已清空
			if v, _ := GetCellValue(p, "Sheet1", "A"+strconv.Itoa(i)); v != nil {
				t.Errorf("A%d 应被清空，实际 = %#v", i, v)
			}
			if v, _ := GetCellValue(p, "Sheet1", "B"+strconv.Itoa(i)); v != nil {
				t.Errorf("B%d 应被清空，实际 = %#v", i, v)
			}
		}
		// 行号必须 1 基（曾出现 <row r="0"> 被 Excel 判损坏）
		data := readZipPart(t, p, "xl/worksheets/sheet1.xml")
		for _, m := range regexp.MustCompile(`<row r="(\d+)"`).FindAllStringSubmatch(data, -1) {
			n, _ := strconv.Atoi(m[1])
			if n < 1 {
				t.Errorf("行号 %d 非法：OOXML 要求 1 基，写 0 会让 Excel 判文件损坏", n)
			}
		}
		// 形状不匹配且目标非单格时应报错
		if err := ws.MoveRange("A1:B3", "D1:E5"); err == nil {
			t.Error("形状不匹配应报错")
		}
	})
}

// TestNamedStyleAndOutline 验证命名样式与大纲折叠。
//
// 两个易错点被本测试锁住：
//  1. summaryBelow 属于 <sheetPr> 里的 <outlinePr> 子元素，**不是 sheetPr 的属性**
//     —— 写错位置会让 openpyxl 抛 "unexpected keyword argument"。
//  2. 折叠要同时设 hidden（隐藏数据行）与 collapsed（分组起点），
//     只做前者 Excel 里看不到折叠控件。
func TestNamedStyleAndOutline(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "nso.xlsx")
	b, _ := Create()
	mustNoErr(t, b.SaveAs(p))
	ws, _ := b.Sheet("Sheet1")
	for i := 1; i <= 10; i++ {
		ws.SetCellStr("A"+strconv.Itoa(i), "行"+strconv.Itoa(i))
	}

	// --- 命名样式 ---
	name, err := ws.NewNamedStyle("标题样式", Style{
		Font:      &FontStyle{Bold: true, Size: 14, Color: "FFFF0000"},
		Fill:      &FillStyle{Color: "FFFFFF00", PatternType: "solid"},
		Alignment: &AlignmentStyle{Horizontal: "center"},
	})
	mustNoErr(t, err)
	if name != "标题样式" {
		t.Errorf("NewNamedStyle 返回 %q", name)
	}
	mustNoErr(t, ws.SetNamedStyle("A1:A2", name))
	list, err := ws.GetNamedStyles()
	mustNoErr(t, err)
	if len(list) != 1 || list[0].Name != "标题样式" {
		t.Errorf("GetNamedStyles = %+v，期望含 标题样式", list)
	}
	// 样式确实落到单元���上
	st, err := ws.GetStyle("A1")
	mustNoErr(t, err)
	if st.Font == nil || !st.Font.Bold {
		t.Errorf("命名样式未应用到 A1: %+v", st.Font)
	}
	// 不存在的样式必须报错
	if err := ws.SetNamedStyle("A1", "查无此样式"); err == nil {
		t.Error("应用不存在的命名样式应报错")
	}
	// 非法样式仍应被枚举校验拦下
	if _, err := ws.NewNamedStyle("坏样式", Style{
		Border: &BorderStyle{Left: BorderSide{Style: "不存在的线型"}}}); err == nil {
		t.Error("非法边框线型应被拒绝")
	}

	// --- 大纲折叠 ---
	mustNoErr(t, ws.GroupRows(3, 8, 1))
	mustNoErr(t, ws.CollapseRows(3, 8))
	if !ws.IsRowHidden(3) {
		t.Error("折叠后第 3 行应隐藏")
	}
	if ws.IsRowHidden(9) {
		t.Error("第 9 行不在折叠区间，不应隐藏")
	}
	mustNoErr(t, ws.ExpandRows(3, 8))
	if ws.IsRowHidden(3) {
		t.Error("展开后第 3 行不应隐藏")
	}
	mustNoErr(t, ws.SetOutlineSummary(false))
	mustNoErr(t, b.Save())

	// 1) openpyxl 必须能严格解析
	snap := runSnapshot(t, p)
	sheets := snap["sheets"].(map[string]interface{})
	if _, ok := sheets["Sheet1"]; !ok {
		t.Fatalf("快照缺少 Sheet1: %v", snap)
	}

	// 2) outlinePr 必须在 sheetPr 内且带 summaryBelow
	data := readZipPart(t, p, "xl/worksheets/sheet1.xml")
	sheetPrRe := regexp.MustCompile(`(?s)<sheetPr\b.*?</sheetPr>|<sheetPr\b[^>]*/>`)
	sheetPr := sheetPrRe.FindString(data)
	if sheetPr == "" {
		t.Fatalf("未写入 sheetPr: %s", data[:minInt(len(data), 400)])
	}
	if !strings.Contains(sheetPr, "<outlinePr") {
		t.Errorf("outlinePr 应在 sheetPr 内，实际 sheetPr = %s", sheetPr)
	}
	if !strings.Contains(sheetPr, `summaryBelow="0"`) {
		t.Errorf("summaryBelow 未写入（应在上方）: %s", sheetPr)
	}
	// 反例：summaryBelow 不能挂在 sheetPr 属性上
	if strings.Contains(sheetPr, `<sheetPr summaryBelow=`) {
		t.Errorf("summaryBelow 挂在了 sheetPr 属性上 —— " +
			"它属于 <outlinePr> 子元素，写错位置会让消费方解析失败")
	}

	// 3) 分组行必须有 outlineLevel
	n := len(regexp.MustCompile(`<row\b[^>]*\boutlineLevel="1"`).FindAllString(data, -1))
	if n != 6 {
		t.Errorf("带 outlineLevel=\"1\" 的行数 = %d，期望 6（第 3~8 行）", n)
	}

	// 4) styles.xml 的 cellStyles 与 cellStyleXfs 必须成对存在
	stylesXML := readZipPart(t, p, "xl/styles.xml")
	if !strings.Contains(stylesXML, `<cellStyles`) {
		t.Errorf("styles.xml 缺少 <cellStyles>: %s", stylesXML[:minInt(len(stylesXML), 300)])
	}
	if !strings.Contains(stylesXML, "标题样式") {
		t.Error("cellStyles 里缺少命名样式 标题样式")
	}
	if !strings.Contains(stylesXML, `<cellStyleXfs`) {
		t.Error("styles.xml 缺少 <cellStyleXfs>（命名样式需底层 xf）")
	}
	// OOXML 要求 cellStyles 排在 dxfs 之前
	if iCS, iDx := strings.Index(stylesXML, "<cellStyles"),
		strings.Index(stylesXML, "<dxfs"); iCS > 0 && iDx > 0 && iCS > iDx {
		t.Error("cellStyles 必须排在 dxfs 之前（OOXML 元素顺序要求），否则 Excel 判损坏")
	}
}

// TestNamedStyleSharedByWorkbook 验证命名样式在多个工作表间可见。
func TestNamedStyleSharedByWorkbook(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "ns2.xlsx")
	b, _ := Create()
	mustNoErr(t, b.SaveAs(p))
	ws, _ := b.Sheet("Sheet1")
	_, err := b.NewSheet("第二表")
	mustNoErr(t, err)

	name, err := ws.NewNamedStyle("共享样式", Style{Font: &FontStyle{Bold: true}})
	mustNoErr(t, err)
	ws2, _ := b.Sheet("第二表")
	mustNoErr(t, ws2.SetNamedStyle("A1", name))
	mustNoErr(t, b.Save())

	// 第二表也应能列出该样式（命名样式是工作簿级的）
	list, err := GetNamedStyles(p, "第二表")
	mustNoErr(t, err)
	if len(list) != 1 || list[0].Name != "共享样式" {
		t.Errorf("第二表看到的命名样式 = %+v", list)
	}
	st, err := ws2.GetStyle("A1")
	mustNoErr(t, err)
	if st.Font == nil || !st.Font.Bold {
		t.Errorf("第二表 A1 未应用共享样式: %+v", st.Font)
	}
}

// TestStreamReader 验证只读流式模式。
//
// 价值：Open() 一次性把整个 zip 读进内存（map[string][]byte），大文件会吃掉
// 数百 MB。OpenReader 只解析 workbook.xml 与 sharedStrings，sheet XML 按需读，
// 且不载入 styles/theme/drawings 等与读单元格无关的部件。
//
// 一致性要求：流式读到的值必须与 Open 完全相同 —— 否则流式就是个"差不多对"的
// 第二套实现，不能用。
func TestStreamReader(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "stream.xlsx")

	// 造一份含各类单元格的数据
	b, _ := Create()
	mustNoErr(t, b.SaveAs(p))
	ws, _ := b.Sheet("Sheet1")
	rows := 500
	for i := 1; i <= rows; i++ {
		r := strconv.Itoa(i)
		ws.SetCellStr("A"+r, "名称-"+r)
		ws.SetCellStr("B"+r, "分类"+r)
		ws.SetCellNumeric("C"+r, float64(i))
		ws.SetCellNumeric("D"+r, float64(i)*1.5)
		if i%10 == 0 {
			ws.SetCellFormula("E"+r, "C"+r+"*2", strconv.Itoa(i*2))
		}
		if i%20 == 0 {
			ws.SetCellBool("F"+r, true)
		}
	}
	mustNoErr(t, b.SaveAs(p))

	sb, err := OpenReader(p)
	mustNoErr(t, err)
	defer sb.Close()

	t.Run("表清单", func(t *testing.T) {
		names := sb.GetSheetList()
		if len(names) != 1 || names[0] != "Sheet1" {
			t.Errorf("GetSheetList = %v", names)
		}
	})

	t.Run("单行读取与 Open 一致", func(t *testing.T) {
		got, err := sb.StreamGetRow("Sheet1", 100)
		mustNoErr(t, err)
		want, err := GetRange(p, "Sheet1", "A100:F100")
		mustNoErr(t, err)
		if len(want) != 1 {
			t.Fatalf("GetRange 返回 %d 行", len(want))
		}
		if len(got) == 0 {
			t.Fatal("StreamGetRow 返回空")
		}
		// 流式返回的是"到该行为止的所有列"，GetRange 也一样
		for i := range got {
			g := normCellForCompare(got[i])
			w := normCellForCompare(want[0][i])
			// 公式单元格两侧都返回结果值或公式文本，容许其一
			if g != w && w != "" {
				t.Errorf("第100行第%d列不一致：流式=%q Open=%q", i+1, g, w)
			}
		}
	})

	t.Run("单格读取与 Open 一致", func(t *testing.T) {
		for _, cell := range []string{"A1", "C250", "D499", "E20"} {
			got, err := sb.StreamGetCell("Sheet1", cell)
			mustNoErr(t, err)
			want, err := GetCellValue(p, "Sheet1", cell)
			mustNoErr(t, err)
			w := normCellForCompare(want)
			if got != w && w != "" {
				t.Errorf("%s 不一致：流式=%q Open=%q", cell, got, w)
			}
		}
	})

	t.Run("全表读取与 GetRows 一致", func(t *testing.T) {
		all, err := sb.StreamGetRows("Sheet1")
		mustNoErr(t, err)
		ref, err := GetRows(p, "Sheet1")
		mustNoErr(t, err)
		if len(all) != len(ref) {
			t.Fatalf("行数不一致：流式=%d Open=%d", len(all), len(ref))
		}
		bad := 0
		for i := range ref {
			for j := range ref[i] {
				if j >= len(all[i]) {
					continue
				}
				if all[i][j] != ref[i][j] {
					bad++
					if bad <= 3 {
						t.Errorf("第%d行第%d列不一致：流式=%q Open=%q",
							i+1, j+1, all[i][j], ref[i][j])
					}
				}
			}
		}
	})

	t.Run("迭代器", func(t *testing.T) {
		it, err := sb.StreamRows("Sheet1")
		mustNoErr(t, err)
		defer it.Close()
		n := 0
		for it.Next() {
			n++
			if n == 100 {
				if rn := it.RowNum(); rn != 100 {
					t.Errorf("第 100 次迭代的 RowNum = %d，期望 100", rn)
				}
				// 第 100 行是 10 与 20 的倍数，故有 6 列（A~F）
				if len(it.Row()) != 6 {
					t.Errorf("第100行应有 6 列，实际 %d: %v", len(it.Row()), it.Row())
				}
			}
		}
		mustNoErr(t, it.Error())
		if n != rows {
			t.Errorf("迭代得到 %d 行，期望 %d", n, rows)
		}
	})

	t.Run("非法表名报错", func(t *testing.T) {
		if _, err := sb.StreamGetRow("查无此表", 1); err == nil {
			t.Error("不存在的表应报错")
		}
	})
}

// TestChartCreation 验证图表创建：部件链完整、openpyxl 能读回、非法参数被拒。
//
// 两个易错点：
//  1. drawing 根元素若写成自闭合（<xdr:wsDr/>），后续追加的锚点会落在根元素
//     之外，产出 "junk after document element" 的坏 XML。
//  2. grouping 的合法值**因图表类型而异**：条形图有 clustered，折线/面积图只有
//     standard。给折线图写 clustered 会让消费方直接拒绝加载整个工作簿。
func TestChartCreation(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "chart.xlsx")
	b, _ := Create()
	mustNoErr(t, b.SaveAs(p))
	ws, _ := b.Sheet("Sheet1")
	ws.SetCellStr("A1", "月份")
	ws.SetCellStr("B1", "产值")
	ws.SetCellStr("C1", "成本")
	for i := 1; i <= 6; i++ {
		r := strconv.Itoa(i)
		ws.SetCellStr("A"+r, "2026-0"+r)
		ws.SetCellNumeric("B"+r, float64(i*100))
		ws.SetCellNumeric("C"+r, float64(i*80))
	}

	// 折线图：两个系列
	mustNoErr(t, ws.AddChart(ChartOptions{
		Type:   ChartLine,
		Title:  "月度趋势",
		Anchor: "E2",
		Series: []ChartSeries{
			{Name: "产值", Categories: "Sheet1!$A$2:$A$7", Values: "Sheet1!$B$2:$B$7"},
			{Name: "成本", Categories: "Sheet1!$A$2:$A$7", Values: "Sheet1!$C$2:$C$7"},
		},
	}))
	// 条形图：验证同一张表可有多个图表
	mustNoErr(t, ws.AddChart(ChartOptions{
		Type:     ChartBar,
		Title:    "对比",
		Anchor:   "E20",
		BarDir:   "col",
		Grouping: "clustered",
		Series: []ChartSeries{
			{NameRef: "Sheet1!$B$1", Values: "Sheet1!$B$2:$B$7"},
		},
	}))
	mustNoErr(t, b.Save())

	// 1) openpyxl 必须能严格解析（图表结构非法会在这里暴露）
	snap := runSnapshot(t, p)
	sheets := snap["sheets"].(map[string]interface{})
	if _, ok := sheets["Sheet1"]; !ok {
		t.Fatalf("快照缺少 Sheet1: %v", snap)
	}

	// 2) 部件链：sheet -> drawing -> chart
	requirePart(t, p, "xl/drawings/drawing1.xml", "图表应建立 drawing 部件")
	requirePart(t, p, "xl/worksheets/_rels/sheet1.xml.rels", "sheet 应有 drawing 关系")
	if countParts(t, p, "xl/charts/chart") < 2 {
		t.Errorf("chart 部件数 = %d，期望 ≥2", countParts(t, p, "xl/charts/chart"))
	}
	// drawing 的 rels 指向 chart
	dRels := "xl/drawings/_rels/drawing1.xml.rels"
	if !partExists(t, p, dRels) {
		t.Fatalf("缺少 %s", dRels)
	}
	nChartRels := 0
	for _, e := range parseRels(readPartStr(t, p, dRels)) {
		if strings.Contains(e.target, "charts/chart") {
			nChartRels++
			if !partExists(t, p, resolveRelTarget("xl/drawings", e.target)) {
				t.Errorf("drawing 指向的图表不存在: %s", e.target)
			}
		}
	}
	if nChartRels < 2 {
		t.Errorf("drawing 的 chart 关系数 = %d，期望 ≥2", nChartRels)
	}

	// 3) drawing 根元素必须非自闭合（否则锚点落在根元素外）
	dw := readPartStr(t, p, "xl/drawings/drawing1.xml")
	if !strings.Contains(dw, "</xdr:wsDr>") {
		t.Errorf("drawing 根元素未正常闭合，锚点会落在根元素之外: %s", dw[:minInt(len(dw), 200)])
	}
	if !strings.Contains(dw, "<xdr:twoCellAnchor") {
		t.Error("drawing 缺少图表锚点")
	}

	// 4) Content_Types 必须为 chart/drawing 声明 override
	ct := readPartStr(t, p, "[Content_Types].xml")
	for _, want := range []string{"/xl/charts/chart1.xml", "/xl/drawings/drawing1.xml"} {
		if !strings.Contains(ct, want) {
			t.Errorf("[Content_Types].xml 缺少 %s 的 Override", want)
		}
	}

	// 5) 关系不得悬空
	assertNoDanglingRels(t, p)

	// 6) 非法参数必须被提前拒绝
	for _, bad := range []struct {
		desc string
		opts ChartOptions
	}{
		{"不存在的类型", ChartOptions{Type: "不存在的图",
			Series: []ChartSeries{{Name: "x", Values: "Sheet1!$B$2"}}}},
		{"无系列", ChartOptions{Type: ChartBar}},
		{"系列缺 Values", ChartOptions{Type: ChartBar,
			Series: []ChartSeries{{Name: "x"}}}},
		{"系列缺名称", ChartOptions{Type: ChartBar,
			Series: []ChartSeries{{Values: "Sheet1!$B$2"}}}},
		{"非法锚点", ChartOptions{Anchor: "!!bad",
			Series: []ChartSeries{{Name: "x", Values: "Sheet1!$B$2"}}}},
		{"非法条形方向", ChartOptions{BarDir: "横着",
			Series: []ChartSeries{{Name: "x", Values: "Sheet1!$B$2"}}}},
		{"折线图用 clustered 分组", ChartOptions{Type: ChartLine, Grouping: "clustered",
			Series: []ChartSeries{{Name: "x", Values: "Sheet1!$B$2"}}}},
		{"非法图例位置", ChartOptions{LegendPos: "左上",
			Series: []ChartSeries{{Name: "x", Values: "Sheet1!$B$2"}}}},
	} {
		if err := validateChartOptions(&bad.opts); err == nil {
			t.Errorf("%s：应报错却通过了（会产出消费方读不出来的文件）", bad.desc)
		}
	}

	// 7) 各图表类型的默认分组必须合法
	for _, ct := range []ChartType{ChartBar, ChartLine, ChartArea, ChartPie,
		ChartScatter, ChartRadar} {
		o := ChartOptions{Type: ct, Series: []ChartSeries{{Name: "x", Values: "S!$A$1"}}}
		if err := validateChartOptions(&o); err != nil {
			t.Errorf("%s 的默认配置应合法: %v", ct, err)
		}
	}
}

// makePrinterSettingsWorkbook 生成一份含打印机设置二进制的 xlsx。
// 打印机设置无法由本库构造（它编码的是具体打印机驱动的能力集），
// 故用等价文件作为测试输入。
func makePrinterSettingsWorkbook(t *testing.T, out string) {
	t.Helper()
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	base := filepath.Join(dir, "base.xlsx")
	runOpenpyxlOps(t, base, `{"out": `+quoteJSON(base)+
		`, "ops": [{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"x"}]}`)
	script := filepath.Join(difftestDir(t), "make_printer.py")
	if o, err := execCommand(pythonExe(), script, base, out).CombinedOutput(); err != nil {
		t.Fatalf("生成打印机设置文件失败: %v\n%s", err, o)
	} else if !strings.Contains(string(o), "BUILD_OK") {
		t.Fatalf("未返回 BUILD_OK: %s", o)
	}
	if !partExists(t, out, "xl/printerSettings/printerSettings1.bin") {
		t.Fatalf("生成的文件缺打印机部件: %v", partNames(t, out))
	}
}

// TestPrinterSettingsPreserved 验证打印机设置二进制的搬运与不破坏。
//
// 这类部件是真实文件里才有、本库无法生成的东西，坑也在这里：
//  1. 改页面设置（纸张/方向）会把 <pageSetup> 整体替换掉，连带抹掉 r:id ——
//     部件与关系都还在（不报错），但 worksheet 不再引用，打印机设置**静默丢失**。
//  2. Content_Types 里必须是 <Default Extension="bin">，写成 Override 会被判非法。
//  3. 二次写入不应堆出多个 .bin 部件。
func TestPrinterSettingsPreserved(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	src := filepath.Join(dir, "src.xlsx")
	makePrinterSettingsWorkbook(t, src)

	// --- 1) 读取与往返 ---
	data, err := GetPrinterSettings(src, "Sheet1")
	mustNoErr(t, err)
	if len(data) == 0 {
		t.Fatal("读不到打印机设置二进制")
	}

	dst := filepath.Join(dir, "dst.xlsx")
	b, _ := Create()
	mustNoErr(t, b.SaveAs(dst))
	mustNoErr(t, SetPrinterSettings(dst, "Sheet1", data))
	back, err := GetPrinterSettings(dst, "Sheet1")
	mustNoErr(t, err)
	if !bytes.Equal(data, back) {
		t.Errorf("打印机设置往返字节不一致：原 %d 字节，回读 %d 字节",
			len(data), len(back))
	}
	// openpyxl 必须能严格解析
	snap := runSnapshot(t, dst)
	if _, ok := snap["sheets"].(map[string]interface{}); !ok {
		t.Fatalf("openpyxl 无法解析产物: %v", snap)
	}
	// Content_Types 必须是 Default 而非 Override
	ct := readPartStr(t, dst, "[Content_Types].xml")
	if !strings.Contains(ct, `<Default Extension="bin"`) {
		t.Errorf("缺少 <Default Extension=\"bin\"> 声明:\n%s", ct)
	}
	if strings.Contains(ct, `<Override`) && strings.Contains(
		regexp.MustCompile(`<Override[^>]*\.bin`).FindString(ct), ".bin") {
		t.Error("bin 不应写成 Override —— 它是扩展名级声明，Excel 会判非法")
	}

	// --- 2) 二次写入复用同一部件，不堆垃圾 ---
	mustNoErr(t, SetPrinterSettings(dst, "Sheet1", data))
	if n := countParts(t, dst, "xl/printerSettings/printerSettings"); n != 1 {
		t.Errorf("打印机部件数 = %d，期望 1（二次写入应复用而非新增）", n)
	}
	if n := countDefaultBin(t, dst); n != 1 {
		t.Errorf("Default bin 声明数 = %d，期望 1（应幂等）", n)
	}

	// --- 3) 改页面设置不得抹掉 r:id ---
	g, err := Open(dst)
	mustNoErr(t, err)
	ws, err := g.Sheet("Sheet1")
	mustNoErr(t, err)
	mustNoErr(t, ws.SetProps(SheetProps{
		PageSetup:   &PageSetup{PaperSize: 9, Orientation: "landscape"},
		PageMargins: &PageMargins{Left: 1.5, Right: 1.5, Top: 1, Bottom: 1},
	}))
	mustNoErr(t, g.Save())

	wsXML := readPartStr(t, dst, "xl/worksheets/sheet1.xml")
	tag := regexp.MustCompile(`(?s)<pageSetup\b[^>]*?/?>`).FindString(wsXML)
	if attrOf(tag, "r:id") == "" {
		t.Errorf("改页面设置后 r:id 被抹掉，打印机设置静默丢失: %s", tag)
	}
	if !strings.Contains(wsXML, `orientation="landscape"`) {
		t.Errorf("orientation 未生效: %s", tag)
	}
	after, err := GetPrinterSettings(dst, "Sheet1")
	mustNoErr(t, err)
	if !bytes.Equal(data, after) {
		t.Error("改页面设置后打印机设置内容变了")
	}
	assertNoDanglingRels(t, dst)

	// --- 4) 复制工作表时二进制跟着搬 ---
	cp := filepath.Join(dir, "cp.xlsx")
	copyFileForTest(t, src, cp)
	g2, err := Open(cp)
	mustNoErr(t, err)
	if _, err := g2.CopySheet("Sheet1", "副本"); err != nil {
		t.Fatal(err)
	}
	mustNoErr(t, g2.Save())
	if n := countParts(t, cp, "xl/printerSettings/printerSettings"); n != 2 {
		t.Errorf("复制后打印机部件数 = %d，期望 2（副本应各有一份）", n)
	}
	// 副本的 r:id 必须指向自己那份，不能与源表共用
	dRID := pageSetupRelIDForSheet(t, cp, "副本")
	sRID := pageSetupRelIDForSheet(t, cp, "Sheet1")
	if dRID == "" {
		t.Error("副本的 pageSetup 没有 r:id")
	}
	if dRID != "" && dRID == sRID {
		t.Errorf("副本与源表共用同一个 r:id=%q —— 各自应指向自己的打印机部件", dRID)
	}
	assertNoDanglingRels(t, cp)

	// --- 5) 空数据必须被拒 ---
	if err := SetPrinterSettings(dst, "Sheet1", nil); err == nil {
		t.Error("空打印机设置应报错")
	}
}

// countDefaultBin 数 Content_Types 里 bin 的 Default 声明数。
func countDefaultBin(t *testing.T, p string) int {
	t.Helper()
	ct := readPartStr(t, p, "[Content_Types].xml")
	return len(regexp.MustCompile(`<Default Extension="bin"`).FindAllString(ct, -1))
}

// pageSetupRelIDForSheet 返回指定表的 pageSetup r:id。
func pageSetupRelIDForSheet(t *testing.T, p, sheetName string) string {
	t.Helper()
	f := findSheetFileFor(t, p, sheetName)
	return attrOf(regexp.MustCompile(`(?s)<pageSetup\b[^>]*?/?>`).
		FindString(readPartStr(t, p, f)), "r:id")
}

// TestStreamWriter 验证 write_only 流式写入。
//
// 动因：常规写入每写一格都重扫整份 sheet XML，复杂度 O(格数 × 文档长度)。
// 实测 3000 行 × 4 列耗时 4 分多钟，写大表不可接受。流式按行追加输出，
// 5000 行实测 30 毫秒。
//
// 一致性要求：流式产物必须能被 openpyxl 严格解析，且样式/列宽/冻结窗格
// 都要正确落地 —— 否则就是个"快但坏"的第二条实现。
func TestStreamWriter(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "stream.xlsx")

	sw, err := NewStreamWriter(p)
	mustNoErr(t, err)

	// 样式必须在写数据之前登记
	head, err := sw.AddStyle(Style{
		Font:      &FontStyle{Bold: true, Color: "FFFFFFFF"},
		Fill:      &FillStyle{Color: "FF4472C4", PatternType: "solid"},
		Alignment: &AlignmentStyle{Horizontal: "center"},
	})
	mustNoErr(t, err)
	if head <= 0 {
		t.Errorf("AddStyle 索引 = %d，期望 >0", head)
	}
	// 相同样式应复用同一索引
	again, err := sw.AddStyle(Style{
		Font:      &FontStyle{Bold: true, Color: "FFFFFFFF"},
		Fill:      &FillStyle{Color: "FF4472C4", PatternType: "solid"},
		Alignment: &AlignmentStyle{Horizontal: "center"},
	})
	mustNoErr(t, err)
	if again != head {
		t.Errorf("相同样式得到不同索引 %d / %d（应去重复用）", again, head)
	}
	mustNoErr(t, sw.SetColWidth(1, 3, 18))
	mustNoErr(t, sw.SetFreezePanes("A2"))

	ws, err := sw.NewSheet("数据")
	mustNoErr(t, err)
	mustNoErr(t, ws.Append([]interface{}{
		StreamCell{Value: "名称", Style: head},
		StreamCell{Value: "数量", Style: head},
		StreamCell{Value: "金额", Style: head},
	}))
	n := 300
	for i := 1; i <= n; i++ {
		r := strconv.Itoa(i)
		mustNoErr(t, ws.Append([]interface{}{"项目-" + r, i, float64(i) * 1.5}))
	}
	// nil 应跳过该格（与 openpyxl 的 None 一致）
	mustNoErr(t, ws.Append([]interface{}{"空值行", nil, nil}))
	if ws.RowCount() != n+2 {
		t.Errorf("RowCount = %d，期望 %d", ws.RowCount(), n+2)
	}

	// 第二张表：验证多表
	ws2, err := sw.NewSheet("第二表")
	mustNoErr(t, err)
	mustNoErr(t, ws2.Append([]interface{}{"单独一行"}))
	if names := sw.SheetNames(); len(names) != 2 || names[0] != "数据" {
		t.Errorf("SheetNames = %v", names)
	}
	mustNoErr(t, sw.Save())

	// 1) openpyxl 必须能严格解析
	snap := runSnapshot(t, p)
	sheets, _ := snap["sheets"].(map[string]interface{})
	s1, _ := sheets["数据"].(map[string]interface{})
	if s1 == nil {
		t.Fatalf("快照缺少 数据 表: %v", snap)
	}
	if s2, _ := sheets["第二表"].(map[string]interface{}); s2 == nil {
		t.Errorf("快照缺少 第二表: %v", sheets)
	}

	// 2) 值与类型必须与常规写入一致
	for _, tc := range []struct {
		cell, want string
	}{
		{"A1", "名称"}, {"A2", "项目-1"}, {"B2", "1"}, {"C2", "1.5"},
		{"A300", "项目-299"}, {"C300", "448.5"},
	} {
		got, err := GetCell(p, "数据", tc.cell)
		mustNoErr(t, err)
		if got != tc.want {
			t.Errorf("%s = %q，期望 %q", tc.cell, got, tc.want)
		}
	}
	// 数值类型不能退化成字符串
	if v, _ := GetCellValue(p, "数据", "B2"); kindOfCellValue(v) != "num" {
		t.Errorf("B2 类型 = %q（%v），期望 num", kindOfCellValue(v), v)
	}

	// 3) nil 格不得写成空字符串
	if v, _ := GetCellValue(p, "数据", "B"+strconv.Itoa(n+2)); v != nil {
		t.Errorf("nil 格应为空，实际 = %#v", v)
	}

	// 4) 样式必须落地
	st, err := GetCellStyle(p, "数据", "A1")
	mustNoErr(t, err)
	if st.Font == nil || !st.Font.Bold {
		t.Errorf("表头粗体未生效: %+v", st.Font)
	}
	if st.Alignment == nil || st.Alignment.Horizontal != "center" {
		t.Errorf("表头居中未生效: %+v", st.Alignment)
	}

	// 5) 列宽与冻结窗格（注意：wsFile 必须确实是"数据"表的部件 ——
	//    findSheetFileFor 按 workbook.xml 的表顺序解析，别张冠李戴）
	wsFile := findSheetFileFor(t, p, "数据")
	if !strings.Contains(readPartStr(t, p, wsFile), "项目-") {
		t.Fatalf("wsFile 不是数据表的部件: %s", wsFile)
	}
	s := readPartStr(t, p, wsFile)
	if !strings.Contains(s, `<col min="1" max="3" width="18"`) {
		t.Errorf("列宽未写入: %s", s[:minInt(len(s), 400)])
	}
	if !strings.Contains(s, `state="frozen"`) {
		t.Error("冻结窗格未写入")
	}
	// cols 必须排在 sheetData 之前（OOXML 元素顺序要求）
	if strings.Index(s, "<cols>") > strings.Index(s, "<sheetData>") {
		t.Error("<cols> 排在了 <sheetData> 之后，Excel 会判文件损坏")
	}

	// 6) 行号必须是 1 基连续（曾出现 <row r="0">）
	rows := regexp.MustCompile(`<row r="(\d+)"`).FindAllStringSubmatch(s, -1)
	if len(rows) != n+2 {
		t.Errorf("行数 = %d，期望 %d", len(rows), n+2)
	}
	for i, m := range rows {
		got, _ := strconv.Atoi(m[1])
		if got != i+1 {
			t.Fatalf("第 %d 个 <row> 的 r = %d，期望 %d", i, got, i+1)
		}
	}

	// 7) 非法输入必须被提前拒绝
	if _, err := NewSheetOf(t, ""); err == nil {
		t.Error("空表名应报错")
	}
	// 同一写入器内重名应报错
	{
		w, err := NewStreamWriter(filepath.Join(dir, "dup.xlsx"))
		mustNoErr(t, err)
		_, err = w.NewSheet("同名")
		mustNoErr(t, err)
		if _, err := w.NewSheet("同名"); err == nil {
			t.Error("同一写入器内重名表应报错")
		}
		w.Abort()
	}

	// 8) 非法样式/参数
	w3, err := NewStreamWriter(filepath.Join(dir, "bad.xlsx"))
	mustNoErr(t, err)
	if _, err := w3.AddStyle(Style{Border: &BorderStyle{
		Left: BorderSide{Style: "不存在的线型"}}}); err == nil {
		t.Error("非法边框线型应被拒（流式路径也要过枚举校验）")
	}
	if err := w3.SetColWidth(3, 1, 10); err == nil {
		t.Error("反向列范围应报错")
	}
	if err := w3.SetFreezePanes("!!bad"); err == nil {
		t.Error("非法冻结位置应报错")
	}
	if err := w3.MergeCell("!!bad"); err == nil {
		t.Error("非法合并区域应报错")
	}
	w3.Abort()

	// 9) Abort 后不应留下半成品文件
	if _, err := os.Stat(filepath.Join(dir, "bad.xlsx")); err == nil {
		t.Error("Abort 后仍留下文件（用户会拿到一个打不开的 xlsx）")
	}
}

// NewSheetOf 建一个临时写入器并在其中声明一张表（用于测非法参数）。
// 声明失败也算返回 nil —— 调用方只关心"有没有报错"。
func NewSheetOf(t *testing.T, name string) (*StreamSheet, error) {
	t.Helper()
	w, err := NewStreamWriter(filepath.Join(t.TempDir(), "t.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	defer w.Abort()
	return w.NewSheet(name)
}

// TestStreamWriterVsNormal 一致性：同一批数据，流式与常规写入产出相同的语义。
//
// 这是流式模式能否被信任的核心判据 —— 它是第二条写入路径，
// 若两边的语义有偏差，用户就必须知道自己"在用哪条路径"才有意义。
func TestStreamWriterVsNormal(t *testing.T) {
	dir := t.TempDir()
	data := [][]interface{}{
		{"名称", "数量", "金额", "日期"},
		{"甲", 10, 100.5, "2026-01-01"},
		{"乙", 20, 200.5, "2026-02-01"},
		{"丙", -30, 0.25, "2026-03-01"},
	}

	// 常规路径
	normal := filepath.Join(dir, "normal.xlsx")
	nb, _ := Create()
	mustNoErr(t, nb.SaveAs(normal))
	nws, _ := nb.Sheet("Sheet1")
	for _, row := range data {
		nws.AppendRow(toAny(row))
	}
	mustNoErr(t, nb.Save())

	// 流式路径
	streamed := filepath.Join(dir, "stream.xlsx")
	sw, err := NewStreamWriter(streamed)
	mustNoErr(t, err)
	sws, err := sw.NewSheet("Sheet1")
	mustNoErr(t, err)
	for _, row := range data {
		mustNoErr(t, sws.Append(toAny(row)))
	}
	mustNoErr(t, sw.Save())

	// 逐格比对
	for ri := range data {
		for ci := range data[ri] {
			ref := refOf(ri, ci)
			a, err1 := GetCellValue(normal, "Sheet1", ref)
			b, err2 := GetCellValue(streamed, "Sheet1", ref)
			if err1 != nil || err2 != nil {
				t.Fatalf("%s 读取失败: %v / %v", ref, err1, err2)
			}
			if normCellForCompare(a) != normCellForCompare(b) {
				t.Errorf("%s 不一致: 常规=%q(%v) 流式=%q(%v)",
					ref, a, a, b, b)
			}
		}
	}
}

// toAny 把强类型切片转成 []interface{}。
func toAny(row []interface{}) []interface{} { return row }

// TestCellValueTypeParity 验证包级与方法的 SetCellValue 支持**同一套类型**，
// 且都覆盖 time.Time。
//
// 这类不一致极难察觉：同一段数据，换成包级调用就报
// "不支持的值类型: time.Time"，而方法级正常 —— 用户只会以为自己写错了。
//
// 三条写入路径（包级 / 方法级 / 流式）必须对同一批类型给出一致结果。
func TestCellValueTypeParity(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "parity.xlsx")
	b, _ := Create()
	mustNoErr(t, b.SaveAs(p))
	ws, _ := b.Sheet("Sheet1")

	tm := time.Date(2026, 3, 15, 12, 30, 0, 0, time.UTC)
	values := []interface{}{
		nil, "文本", true, false,
		int(1), int8(2), int16(3), int32(4), int64(5),
		uint(6), uint8(7), uint16(8), uint32(9), uint64(10),
		float32(1.5), float64(2.5),
		tm,
	}

	// 方法级：先写完并保存
	for i, v := range values {
		ref := "A" + strconv.Itoa(i+1)
		if err := ws.SetCellValue(ref, v); err != nil {
			t.Errorf("方法级 SetCellValue(%#v) 失败: %v", v, err)
		}
	}
	mustNoErr(t, b.Save())

	// 包级：注意每次调用都是独立的 open→改→save，
	// 必须**先关掉上面的对象写入**否则会被后一次 Save 用旧内存副本覆盖。
	// 这里用另一个文件，避开"两个写入者争同一个文件"的语义。
	gp := filepath.Join(dir, "pkg.xlsx")
	gb, _ := Create()
	mustNoErr(t, gb.SaveAs(gp))
	for i, v := range values {
		ref := "B" + strconv.Itoa(i+1)
		if err := SetCellValue(gp, "Sheet1", ref, v); err != nil {
			t.Errorf("包级 SetCellValue(%#v) 失败: %v", v, err)
		}
	}

	// 流式：**每个值独占一行**（第 N 个值在A(N+1)），
	// 这样三条路径的坐标才对得齐 —— 值列表是纵向的，不是横向的一行。
	sp := filepath.Join(dir, "stream.xlsx")
	sw, err := NewStreamWriter(sp)
	mustNoErr(t, err)
	sws, err := sw.NewSheet("Sheet1")
	mustNoErr(t, err)
	for _, v := range values {
		mustNoErr(t, sws.Append([]interface{}{v}))
	}
	mustNoErr(t, sw.Save())

	// 三条路径的结果必须一致（分别读三个文件）
	for i := range values {
		ref := strconv.Itoa(i + 1)
		m, err1 := GetCellValue(p, "Sheet1", "A"+ref)
		g, err2 := GetCellValue(gp, "Sheet1", "B"+ref)
		s, err3 := GetCellValue(sp, "Sheet1", "A"+ref)
		if err1 != nil || err2 != nil || err3 != nil {
			t.Fatalf("第 %d 项读取失败: %v / %v / %v", i+1, err1, err2, err3)
		}
		if normCellForCompare(m) != normCellForCompare(g) {
			t.Errorf("第 %d 项包级与方法级不一致: %q vs %q",
				i+1, normCellForCompare(m), normCellForCompare(g))
		}
		// 流式与常规对 time.Time 用**不同但都合法**的表示：
		//   常规 = 序列号 + 日期格式（t 缺省，读回是 float）
		//   流式 = ISO8601 字符串（t="d"，读回是 datetime 字符串）
		// 所以日期项单独比对"是否指向同一时刻"，其余项比字符串。
		if _, isTime := values[i].(time.Time); isTime {
			assertSameInstant(t, i+1, gp, sp, "B", "A", tm)
			continue
		}
		if normCellForCompare(g) != normCellForCompare(s) {
			t.Errorf("第 %d 项常规与流式不一致: %q vs %q",
				i+1, normCellForCompare(g), normCellForCompare(s))
		}
	}

	// 不支持的类型必须被明确拒绝，且错误信息列出可用类型
	err = ws.SetCellValue("Z1", struct{}{})
	if err == nil {
		t.Error("不支持的类型应报错")
	} else if !strings.Contains(err.Error(), "time.Time") {
		t.Errorf("错误信息未列出可用类型（应含 time.Time）: %v", err)
	}
	if err := SetCellValue(gp, "Sheet1", "Z2", struct{}{}); err == nil {
		t.Error("包级也不支持的类型应报错")
	}
}

// TestSetCellTime 验证日期时间写入。
//
// 关键点：写time.Time **必须同时设日期格式**，否则 Excel 里显示的是一串
// 数字 —— 值对了但人看不懂，且 GetTime 因格式判定失败而返回 ok=false。
func TestSetCellTime(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "date.xlsx")
	b, _ := Create()
	mustNoErr(t, b.SaveAs(p))
	ws, _ := b.Sheet("Sheet1")

	cases := []struct {
		ref  string
		tm   time.Time
		want string // 期望的数字格式
	}{
		{"A1", time.Date(2026, 3, 15, 0, 0, 0, 0, time.UTC), "yyyy-mm-dd"},
		{"A2", time.Date(2026, 3, 15, 12, 30, 0, 0, time.UTC), "yyyy-mm-dd hh:mm:ss"},
		{"A3", time.Date(2026, 12, 31, 23, 59, 59, 0, time.UTC), "yyyy-mm-dd hh:mm:ss"},
	}
	for _, c := range cases {
		mustNoErr(t, ws.SetCellTime(c.ref, c.tm))
	}
	mustNoErr(t, b.Save())

	for _, c := range cases {
		got, ok := ws.Cell(c.ref).GetTime()
		if !ok {
			t.Errorf("%s: GetTime 返回 ok=false（日期格式未生效?）", c.ref)
			continue
		}
		if !got.Equal(c.tm) {
			t.Errorf("%s 往返不一致: 写入 %v，读回 %v",
				c.ref, c.tm.Format(time.RFC3339), got.Format(time.RFC3339))
		}
		st, err := ws.GetStyle(c.ref)
		mustNoErr(t, err)
		if st.NumFmt != c.want {
			t.Errorf("%s 数字格式 = %q，期望 %q", c.ref, st.NumFmt, c.want)
		}
	}

	// openpyxl 必须能读成真正的 datetime（这是外部实现的独立验证）
	if pythonExe() == "" {
		return
	}
	snap := runSnapshot(t, p)
	sheets := snap["sheets"].(map[string]interface{})
	s1, _ := sheets["Sheet1"].(map[string]interface{})
	cells, _ := s1["cells"].(map[string]interface{})
	a1, _ := cells["A1"].(map[string]interface{})
	val, _ := a1["value"].(map[string]interface{})
	if val != nil {
		if k, _ := val["k"].(string); k == "str" || k == "other" {
			if got, _ := val["v"].(string); strings.HasPrefix(got, "2026-03-15") {
				return // openpyxl 独立验证通过
			}
		}
	}
	// 退回直接检查：序列号与日期格式必须配套，否则 Excel 里显示为纯数字。
	// 样式（NumFmt）在 styles.xml 而非 sheet XML 里，且 sheet 部件编号取决于
	// 工作表顺序 —— 都按名解析，别写死路径。
	sheetXML := readPartStr(t, p, findSheetFileFor(t, p, "Sheet1"))
	if !strings.Contains(sheetXML, "46096") {
		t.Errorf("A1 未写出 46096 序列号（openpyxl 侧识别为日期的前提）")
	}
	stylesXML := readPartStr(t, p, "xl/styles.xml")
	if !strings.Contains(stylesXML, "yyyy-mm-dd") {
		t.Errorf("日期格式未写入 styles.xml —— 序列号与格式不配套时 Excel 里显示为纯数字")
	}

	// 时间部分必须真的进了序列号的小数位
	// （GetCellValue 对数值返回 float64/int，不是 string —— 断言要按类型取值）
	raw, err := GetCellValue(p, "Sheet1", "A1")
	mustNoErr(t, err)
	rawA2, err := GetCellValue(p, "Sheet1", "A2")
	mustNoErr(t, err)
	if kindOfCellValue(rawA2) != "num" {
		t.Errorf("A2 类型 = %q（%#v），期望 num", kindOfCellValue(rawA2), rawA2)
	}
	serialA2 := parseSerial(t, rawA2)
	if serialA2 == 0 {
		t.Errorf("A2 序列号 = %v，应为含小数的日期时间", serialA2)
	}
	// 1900 日期系统基准：2026-03-15 应为 46096
	if s := parseSerial(t, raw); s != 46096 {
		t.Errorf("2026-03-15 的序列号 = %v，期望 46096", s)
	}

	// 秒级往返必须精确（曾经差 476~512 纳秒）
	for _, c := range []struct {
		ref string
		tm  time.Time
	}{
		{"A4", time.Date(2026, 12, 31, 23, 59, 59, 0, time.UTC)},
		{"A5", time.Date(2026, 1, 1, 0, 0, 1, 0, time.UTC)},
		{"A6", time.Date(2026, 6, 15, 12, 30, 45, 0, time.UTC)},
	} {
		mustNoErr(t, ws.SetCellTime(c.ref, c.tm))
	}
	mustNoErr(t, b.Save())
	for _, c := range []struct {
		ref string
		tm  time.Time
	}{
		{"A4", time.Date(2026, 12, 31, 23, 59, 59, 0, time.UTC)},
		{"A5", time.Date(2026, 1, 1, 0, 0, 1, 0, time.UTC)},
		{"A6", time.Date(2026, 6, 15, 12, 30, 45, 0, time.UTC)},
	} {
		got, ok := ws.Cell(c.ref).GetTime()
		if !ok {
			t.Errorf("%s GetTime ok=false", c.ref)
			continue
		}
		if !got.Equal(c.tm) {
			t.Errorf("%s 秒级往返不精确: 写入 %v，读回 %v（差 %v）",
				c.ref, c.tm.Format(time.RFC3339Nano),
				got.Format(time.RFC3339Nano), got.Sub(c.tm))
		}
	}
}

// assertSameInstant 断言两个文件里第 idx 行的同一时刻能被识别为相同时间。
//
// idx 是**1 基行号**（与坐标拼接一致），函数内不再 +1。
//
// 常规与流式对 time.Time 用不同表示（序列号 vs ISO8601 字符串），
// 直接比字符串必然不等 —— 必须比对语义。
func assertSameInstant(t *testing.T, idx int, f1, f2, col1, col2 string, want time.Time) {
	t.Helper()
	ws1, err := Open(f1)
	if err != nil {
		t.Fatalf("第 %d 项：打开 %s 失败: %v", idx, f1, err)
	}
	w1, _ := ws1.Sheet("Sheet1")
	got1, ok1 := w1.Cell(col1 + strconv.Itoa(idx)).GetTime()

	ws2, err := Open(f2)
	if err != nil {
		t.Fatalf("第 %d 项：打开 %s 失败: %v", idx, f2, err)
	}
	w2, _ := ws2.Sheet("Sheet1")
	got2, ok2 := w2.Cell(col2 + strconv.Itoa(idx)).GetTime()

	if !ok1 || !ok2 {
		t.Errorf("第 %d 项：常规 ok=%v、流式 ok=%v（两者都应能识别为日期）", idx, ok1, ok2)
		return
	}
	if !got1.Equal(want) {
		t.Errorf("第 %d 项常规侧时刻不对: %v，期望 %v",
			idx, got1.Format(time.RFC3339), want.Format(time.RFC3339))
	}
	if !got2.Equal(want) {
		t.Errorf("第 %d 项流式侧时刻不对: %v，期望 %v",
			idx, got2.Format(time.RFC3339), want.Format(time.RFC3339))
	}
}

// parseSerial 把 GetCellValue 的返回值转成float64 序列号。
func parseSerial(t *testing.T, v interface{}) float64 {
	t.Helper()
	switch n := v.(type) {
	case float64:
		return n
	case int:
		return float64(n)
	case int64:
		return float64(n)
	case string:
		f, err := strconv.ParseFloat(n, 64)
		if err != nil {
			t.Errorf("序列号 %q 无法解析为数字", n)
		}
		return f
	}
	t.Errorf("序列号类型 %T 异常: %#v", v, v)
	return 0
}

// TestMergePreservesConditionalFormatDxfs 回归：合并带条件格式的跨簿文件时，
// **差分样式（dxf）必须一起搬，且 dxfId 必须重映射**。
//
// dxfId 是**工作簿级索引**（与 cellXf 同理）。早期实现只搬 cellXf：
// sheet 里的 dxfId 原样带过去，而目标簿 styles.xml 的 <dxfs> 里没有对应条目
// → 悬空引用。openpyxl 直接抛
//
//	IndexError: list index out of range
//	  at styles/differential.py: __getitem__ (self.styles[idx])
//
// Excel 则报"文件已损坏"。单簿测试永远测不到 —— 必须跨簿合并。
func TestMergePreservesConditionalFormatDxfs(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	src := filepath.Join(dir, "src.xlsx")

	// 用本库造一个带条件格式的工作簿，再合并到一个空簿 ——
	// 直接验证 dxfs 搬运与 dxfId 重映射这条链路。
	mustNoErr(t, Create2(src))
	mustNoErr(t, SetConditionalFormatRules(src, "Sheet1", "B1:B10",
		[]ConditionalFormatRule{
			{Type: "cellIs", Operator: "lessThan", Formula: "50",
				Style: &Style{Fill: &FillStyle{Color: "FFFFFF00", PatternType: "solid"}}},
			{Type: "cellIs", Operator: "lessThan", Formula: "10",
				Style: &Style{Font: &FontStyle{Color: "FF00FF00", Italic: true}}},
		}))
	mustNoErr(t, SetConditionalFormatRules(src, "Sheet1", "C1:C10",
		[]ConditionalFormatRule{
			{Type: "cellIs", Operator: "greaterThan", Formula: "20",
				Style: &Style{Font: &FontStyle{Color: "FF0000FF", Bold: true}}},
		}))

	dst := filepath.Join(dir, "dst.xlsx")
	b, _ := Create()
	mustNoErr(t, b.SaveAs(dst))
	mustNoErr(t, b.Merge([]SourceRef{{Workbook: src, Sheet: "Sheet1"}}))
	mustNoErr(t, b.Save())

	// 1) openpyxl 必须能加载（悬空 dxfId 会在这里抛 IndexError）
	runSnapshot(t, dst)

	// 2) 目标簿的 dxfs 必须存在，且数量与源一致
	srcStyles := readPartStr(t, src, "xl/styles.xml")
	dstStyles := readPartStr(t, dst, "xl/styles.xml")
	nSrc := countDxfIn(srcStyles)
	nDst := countDxfIn(dstStyles)
	if nSrc == 0 {
		t.Fatal("源文件没有 dxf，测试前提不成立")
	}
	if nDst < nSrc {
		t.Errorf("目标簿 dxf 数 = %d，少于源 %d —— 差分样式没搬过来"+
			"（条件格式的 dxfId 会变成悬空引用）", nDst, nSrc)
	}

	// 3) 所有 dxfId 必须落在 dxfs 的实际范围内
	for _, sheetName := range []string{"Sheet1"} {
		f := findSheetFileFor(t, dst, sheetName)
		for _, id := range dxfIDsIn(readPartStr(t, dst, f)) {
			if id < 0 || id >= nDst {
				t.Errorf("%s 的 dxfId=%d 越界（dxfs 只有 %d 条）—— "+
					"openpyxl 会抛 IndexError，Excel 判文件损坏", sheetName, id, nDst)
			}
		}
	}

	// 4) 条件格式规则必须落在**合并进来的那张表**上（不是留在空表里）。
	//    合并进来的表名是源表名 + Suffix，需按名定位而不是假设 sheet1.xml。
	var mergedSheet string
	g2, err := Open(dst)
	if err != nil {
		t.Fatal(err)
	}
	for _, n := range g2.SheetNames() {
		if n != "Sheet1" { // 目标原本的��表
			mergedSheet = n
			break
		}
	}
	if mergedSheet == "" {
		t.Fatalf("未找到合并进来的表: %v", g2.SheetNames())
	}
	merged := dxfIDsIn(readPartStr(t, dst, findSheetFileFor(t, dst, mergedSheet)))
	if len(merged) == 0 {
		t.Errorf("合并进来的表 %s 上没有条件格式规则（dxfId 都没了）", mergedSheet)
	}
}

// countDxfIn 数 styles.xml 里的 <dxf> 条目数。
func countDxfIn(stylesXML string) int {
	return len(regexp.MustCompile(`<dxf>`).FindAllString(stylesXML, -1))
}

// dxfIDsIn 取出 sheet XML 里所有 dxfId 值。
func dxfIDsIn(sheetXML string) []int {
	var out []int
	for _, m := range regexp.MustCompile(`dxfId="(\d+)"`).
		FindAllStringSubmatch(sheetXML, -1) {
		if n, err := strconv.Atoi(m[1]); err == nil {
			out = append(out, n)
		}
	}
	return out
}

// TestPartTransportSingleWritePoint 回归：关联部件搬运必须只有一个写入点。
//
// 症状（真实踩过）：copySheetPartsToTarget 内部重写了工作表的 rId 引用并
// 写回 dstMap，调用方随后又用自己的 newWS 覆盖一次 —— 内层的 rId 重写被抹掉。
// 结果 sheet XML 仍引用源文件的旧 rId（rId1），而 rels 里已换成新 ID（rId5起），
// openpyxl 报 "Unknown relationship: rId1"，Excel 报文件损坏。
//
// 守卫方式：跨簿操作产出的工作表里，所有 r:id 都必须能在其 rels 中找到。
func TestPartTransportSingleWritePoint(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	src := filepath.Join(dir, "src.xlsx")

	// 源簿：一个表 + 图表 + 表格（产生多条 sheet 级关系）
	b0, _ := Create()
	mustNoErr(t, b0.SaveAs(src))
	ws0, _ := b0.Sheet("Sheet1")
	for i := 1; i <= 5; i++ {
		ws0.SetCellStr("A"+strconv.Itoa(i), "类"+strconv.Itoa(i))
		ws0.SetCellNumeric("B"+strconv.Itoa(i), float64(i*10))
	}
	mustNoErr(t, ws0.SetCellStr("D1", "量"))
	mustNoErr(t, ws0.SetCellNumeric("E1", 1))
	mustNoErr(t, AddTable(src, "Sheet1", "D1:E5", "表A"))
	mustNoErr(t, ws0.AddChart(ChartOptions{
		Type: ChartBar, Title: "图", Anchor: "G2",
		Series: []ChartSeries{{NameRef: "Sheet1!$E$1", Values: "Sheet1!$E$2:$E$5"}},
	}))
	mustNoErr(t, b0.Save())

	// 场景 1：Merge 到另一个簿
	dst := filepath.Join(dir, "merged.xlsx")
	b1, _ := Create()
	mustNoErr(t, b1.SaveAs(dst))
	mustNoErr(t, b1.Merge([]SourceRef{{Workbook: src, Sheet: "Sheet1"}}))
	mustNoErr(t, b1.Save())
	assertNoDanglingSheetRefs(t, dst)
	runSnapshot(t, dst) // openpyxl 严格解析：Unknown relationship 会在这里暴露

	// 场景 2：CopySheetTo 到另一个文件
	dst2 := filepath.Join(dir, "copied.xlsx")
	{
		g2, err := Open(src)
		mustNoErr(t, err)
		mustNoErr(t, g2.CopySheetTo(dst2, "Sheet1", "搬过来"))
	}
	assertNoDanglingSheetRefs(t, dst2)
	runSnapshot(t, dst2)
}

// assertNoDanglingSheetRefs 断言工作表 XML 里的每个 r:id 都能在其 rels 中找到。
//
// 这是"rId 重写被覆盖"的直接守卫：症状是 sheet 引用 rId1 但 rels 里没有它。
func assertNoDanglingSheetRefs(t *testing.T, p string) {
	t.Helper()
	wbXML := readPartStr(t, p, "xl/workbook.xml")
	for _, chunk := range strings.Split(wbXML, "<sheet ") {
		name := attrValue(chunk, "name")
		rid := attrValue(chunk, "r:id")
		if name == "" || rid == "" {
			continue
		}
		f := findSheetFileFor(t, p, name)
		if f == "" {
			continue
		}
		relFile := "xl/worksheets/_rels/" + filepathBase(f) + ".rels"
		if partExists(t, p, relFile) {
			ids := map[string]bool{}
			for _, e := range parseRels(readPartStr(t, p, relFile)) {
				ids[e.id] = true
			}
			for _, used := range rIdsIn(readPartStr(t, p, f)) {
				if !ids[used] {
					t.Errorf("工作表 %s 引用了 %s，但它不在 %s 里 —— "+
						"openpyxl 会报 \"Unknown relationship: %s\"，Excel 判文件损坏"+
						"（根因通常是关联部件搬运与工作表写回不是同一个写入点）",
						name, used, relFile, used)
				}
			}
		}
	}
}

// rIdsIn 取出 XML 里所有 r:id / r:embed 的值。
func rIdsIn(s string) []string {
	var out []string
	for _, m := range regexp.MustCompile(`r:(?:id|embed)="([^"]+)"`).
		FindAllStringSubmatch(s, -1) {
		out = append(out, m[1])
	}
	return out
}

// TestNoEmptyIntegerAttrs 回归：Integer 类型的 XML 属性在无值时必须**整个省略**，
// 不能写成空串。
//
// 真实踩过：跨簿合并命名样式时，cellStyleXfs 里的源 xf 若没有 xfId 属性，
// attrOf 返回 ""，buildXF 直接拼成 xfId=""。而 openpyxl 的
// CellStyle.xfId 声明为 Int，_convert("") 失败 ->
//
//	TypeError: expected <class 'int'>
//
// 整个工作簿读不出来。同类 Integer 属性还有 cellStyle 的 builtinId 等。
func TestNoEmptyIntegerAttrs(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()

	// 造一个源簿：命名样式 + 条件格式（两者分别触发 cellStyleXfs 与 dxfs 两条链）
	src := filepath.Join(dir, "src.xlsx")
	b, err := Create()
	mustNoErr(t, err)
	mustNoErr(t, b.SaveAs(src))
	ws, err := b.Sheet("Sheet1")
	mustNoErr(t, err)
	name, err := ws.NewNamedStyle("测试样式", Style{
		Font: &FontStyle{Bold: true, Color: "FF112233"},
		Fill: &FillStyle{Color: "FF445566", PatternType: "solid"},
	})
	mustNoErr(t, err)
	mustNoErr(t, ws.SetNamedStyle("A1:B1", name))
	mustNoErr(t, ws.SetConditionalFormatRules("A1:A10", []ConditionalFormatRule{
		{Type: "cellIs", Operator: "lessThan", Formula: "5",
			Style: &Style{Font: &FontStyle{Color: "FF0000FF"}}},
	}))
	mustNoErr(t, b.Save())

	// 合并到一个空簿
	dst := filepath.Join(dir, "dst.xlsx")
	b2, err := Create()
	mustNoErr(t, err)
	mustNoErr(t, b2.SaveAs(dst))
	mustNoErr(t, b2.Merge([]SourceRef{{Workbook: src, Sheet: "Sheet1"}}))
	mustNoErr(t, b2.Save())

	// 1) openpyxl 必须能加载（xfId="" 会在这里抛 TypeError）
	runSnapshot(t, dst)

	// 2) 直接验证 buildXF：无 xfId 的源 xf 不得产出 xfId=""
	// （组合场景未必每次都触发这条路径，直接测函数更可靠）
	if got := buildXF("0", "1", "2", "3",
		`<xf numFmtId="0" fontId="1" fillId="2" borderId="3"/>`); strings.Contains(got, `xfId=""`) {
		t.Errorf("buildXF 对缺 xfId 的源 xf 产出了空值属性: %s —— "+
			"openpyxl 会抛 \"TypeError: expected <class 'int'>\"", got)
	}
	// 有 xfId 时必须保留
	if got := buildXF("0", "1", "2", "3",
		`<xf numFmtId="0" fontId="1" fillId="2" borderId="3" xfId="5"/>`); !strings.Contains(got, `xfId="5"`) {
		t.Errorf("buildXF 丢掉了源 xf 的 xfId: %s", got)
	}

	// 3) styles.xml 里不得出现空值的 Integer 属性
	styles := readPartStr(t, dst, "xl/styles.xml")
	for _, attr := range []string{"xfId", "builtinId", "numFmtId", "fontId", "fillId", "borderId"} {
		if strings.Contains(styles, attr+`=""`) {
			t.Errorf("styles.xml 里出现空的 Integer 属性 %s=\"\" —— "+
				"消费方按 Int 转换会抛 \"TypeError: expected <class 'int'>\"，"+
				"整份工作簿读不出来", attr)
		}
	}

	// 4) 命名样式必须还在，且 xfId 落在 cellStyleXfs 范围内
	if !strings.Contains(styles, "测试样式") {
		t.Error("命名样式 测试样式 丢失")
	}
	nStyleXfs := len(extractBlockItemsForTest(styles, "cellStyleXfs"))
	for _, id := range dxfIDsIn(readPartStr(t, dst, findSheetFileFor(t, dst, "Sheet1"))) {
		if id >= nStyleXfs && id >= countDxfIn(styles) {
			t.Errorf("dxfId=%d 越界（dxfs 只有 %d 条）", id, countDxfIn(styles))
		}
	}
}

// extractBlockItemsForTest 供测试用：数某区块的子元素。
func extractBlockItemsForTest(xml, block string) []string {
	re := regexp.MustCompile(`(?s)<` + block + `\b[^>]*>(.*?)</` + block + `>`)
	m := re.FindStringSubmatch(xml)
	if m == nil {
		return nil
	}
	return regexp.MustCompile(`(?s)<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>`).FindAllString(m[1], -1)
}

// TestSharedStringsRelationshipRegistered 验证往「原本没有共享字符串表」的已有
// 工作簿写字符串时，会在 workbook.xml.rels 中登记 sharedStrings 关系。
//
// 症状：单元格在 openpyxl 里读得出值，在 Excel/WPS 里整片显示为空。
// 根因：OOXML 里共享字符串表不是"按部件名约定自动加载"的，它必须由
// xl/_rels/workbook.xml.rels 中的一条关系显式指向。excelgo 早期只写了
// xl/sharedStrings.xml 部件并补了 [Content_Types].xml 的 Override，唯独漏了这条
// 关系，于是部件成为"孤儿"：
//   - openpyxl 容错高，能按部件名自行推断，读得出值；
//   - Excel/WPS 严格按关系解析，找不到关系就整片空。
//
// 为什么 Create() 路径上看不到这个 bug：blankWorkbookMap 预置的 workbook rels 里
// 已经含有 sharedStrings 关系（见 book.go），所以只有「Open 一份 Excel/openpyxl
// 生成的文件再写字符串」这条路径会暴露 —— 而这类文件恰恰是工程上最常见的输入。
func TestSharedStringsRelationshipRegistered(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	src := filepath.Join(dir, "src.xlsx")
	// openpyxl 生成的最小工作簿：字符串以内联串存储，不含 sharedStrings.xml，
	// rels 里自然也没有 sharedStrings 关系 —— 正是现实中常见的输入形态。
	runOpenpyxlOps(t, src, `{"out": `+quoteJSON(src)+
		`, "ops": [{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"x"}]}`)
	if partExists(t, src, "xl/sharedStrings.xml") {
		t.Fatal("前置条件不成立：openpyxl 产物本就含 sharedStrings.xml，测不到本bug")
	}

	out := filepath.Join(dir, "out.xlsx")
	b, err := Open(src)
	if err != nil {
		t.Fatalf("打开失败: %v", err)
	}
	ws, err := b.NewSheet("S")
	if err != nil {
		t.Fatalf("新建工作表失败: %v", err)
	}
	for _, ref := range []string{"A1", "A4", "B2", "B3", "G7"} {
		if err := ws.SetCellValue(ref, "val-"+ref); err != nil {
			t.Fatalf("写%s: %v", ref, err)
		}
	}
	if err := b.SaveAs(out); err != nil {
		t.Fatalf("保存失败: %v", err)
	}

	// 1) 部件存在
	requirePart(t, out, "xl/sharedStrings.xml", "写了字符串就该有共享字符串表")
	// 2) Content_Types 有 Override
	if ct := readZipPart(t, out, "[Content_Types].xml"); !strings.Contains(ct, "/xl/sharedStrings.xml") {
		t.Error("[Content_Types].xml 缺 sharedStrings 的 Override")
	}
	// 3) 关键断言：workbook.xml.rels 里有 sharedStrings 关系
	rels := readZipPart(t, out, "xl/_rels/workbook.xml.rels")
	if !strings.Contains(rels, sharedStringsRelType) {
		t.Fatalf("workbook.xml.rels 未登记 sharedStrings 关系，所有 t=\"s\" 单元格在 Excel 里会显示为空：\n%s", rels)
	}
	// 关系必须指向真实存在的部件，不能是悬空引用
	if !relsTargetExists(t, rels, "xl/sharedStrings.xml") {
		t.Errorf("sharedStrings 关系的 Target 未指向真实部件：\n%s", rels)
	}
	// 4) 关系 Id 不得与既有关系冲突
	assertRIdsUnique(t, rels)
	// 5) t="s" 单元格确实指向了合法索引（0 <= idx < uniqueCount）
	assertSharedStringIndicesInRange(t, out)
}

// TestSharedStringsRelationshipIdempotent 验证重复写字符串不会堆出重复的
// sharedStrings 关系 —— 同一个 Target 出现两条关系属于非法，Excel 会判损坏。
func TestSharedStringsRelationshipIdempotent(t *testing.T) {
	dir := t.TempDir()
	out := filepath.Join(dir, "out.xlsx")
	b, err := Create()
	if err != nil {
		t.Fatal(err)
	}
	ws, err := b.NewSheet("S")
	if err != nil {
		t.Fatal(err)
	}
	// 多轮写入，含重复值（重复值仍只应有一条 sharedStrings 关系）
	refs := []string{"A1", "B1", "C1", "D1", "E1", "F1", "G1", "H1"}
	for i, ref := range refs {
		for _, s := range []string{"alpha", "beta", "alpha"} {
			if err := ws.SetCellValue(ref, s+"-"+strconv.Itoa(i)); err != nil {
				t.Fatal(err)
			}
		}
	}
	if err := b.SaveAs(out); err != nil {
		t.Fatal(err)
	}
	rels := readZipPart(t, out, "xl/_rels/workbook.xml.rels")
	if n := strings.Count(rels, sharedStringsRelType); n != 1 {
		t.Fatalf("sharedStrings 关系应恰好 1 条，实得 %d 条：\n%s", n, rels)
	}
	assertRIdsUnique(t, rels)
}

// TestSharedStringsRelNotCreatedWithoutPart 验证部件不存在时不建关系 ——
// 关系指向缺失部件同样是悬空引用。
func TestSharedStringsRelNotCreatedWithoutPart(t *testing.T) {
	fm := map[string][]byte{
		"xl/_rels/workbook.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1" Type="` + worksheetRelType + `" Target="worksheets/sheet1.xml"/>` +
			`</Relationships>`),
	}
	ensureSharedStringsRel(fm)
	if strings.Contains(string(fm["xl/_rels/workbook.xml.rels"]), sharedStringsRelType) {
		t.Error("部件不存在时不应创建 sharedStrings 关系（悬空引用）")
	}
	// 部件存在时才建，且只建一条
	fm["xl/sharedStrings.xml"] = []byte(buildSharedStrings([]string{"a"}))
	ensureSharedStringsRel(fm)
	ensureSharedStringsRel(fm)
	rels := string(fm["xl/_rels/workbook.xml.rels"])
	if n := strings.Count(rels, sharedStringsRelType); n != 1 {
		t.Fatalf("部件存在时应恰好 1 条关系，实得 %d 条：\n%s", n, rels)
	}
}

// relsTargetExists 检查 rels 中是否存在指向 part 的关系。
//
// part 传zip 内的完整部件路径（如 "xl/sharedStrings.xml"），而 rels 里的 Target
// 是**相对所属部件目录**的写法：workbook.xml.rels 的基准是 xl/，故
// "sharedStrings.xml" 与 "/xl/sharedStrings.xml" 都指向同一部件
// （两者都被 Excel 接受），sheet rels 的基准则是 xl/worksheets/。
// 因此三种等价写法都要认。
func relsTargetExists(t *testing.T, relsXML, part string) bool {
	t.Helper()
	re := regexp.MustCompile(`Target="([^"]*)"`)
	base := path.Base(part)
	for _, m := range re.FindAllStringSubmatch(relsXML, -1) {
		tgt := strings.TrimPrefix(m[1], "/")
		if tgt == part || tgt == base || tgt == "xl/"+base {
			return true
		}
	}
	return false
}

// assertRIdsUnique 断言 rels 中没有重复的 Id。
func assertRIdsUnique(t *testing.T, relsXML string) {
	t.Helper()
	re := regexp.MustCompile(`Id="([^"]*)"`)
	seen := map[string]bool{}
	for _, m := range re.FindAllStringSubmatch(relsXML, -1) {
		if seen[m[1]] {
			t.Errorf("workbook.xml.rels 中 Id 重复：%s", m[1])
		}
		seen[m[1]] = true
	}
}

// assertSharedStringIndicesInRange 断言所有 t="s" 单元格的索引都落在
// sharedStrings 表的实际条数范围内（越界即悬空，Excel 显示空）。
func assertSharedStringIndicesInRange(t *testing.T, xlsx string) {
	t.Helper()
	ss := readZipPart(t, xlsx, "xl/sharedStrings.xml")
	n := strings.Count(ss, "<si>")
	re := regexp.MustCompile(`<c\b[^>]*\bt="s"[^>]*>\s*<v>(\d+)</v>`)
	z, err := zip.OpenReader(xlsx)
	if err != nil {
		t.Fatalf("打开 zip 失败: %v", err)
	}
	defer z.Close()
	for _, f := range z.File {
		if !strings.HasPrefix(f.Name, "xl/worksheets/") || !strings.HasSuffix(f.Name, ".xml") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			t.Fatalf("打开部件失败: %v", err)
		}
		body, _ := io.ReadAll(rc)
		rc.Close()
		for _, m := range re.FindAllStringSubmatch(string(body), -1) {
			idx, _ := strconv.Atoi(m[1])
			if idx < 0 || idx >= n {
				t.Errorf("%s 中 t=\"s\" 索引 %d 越界（表内共 %d 条）", f.Name, idx, n)
			}
		}
	}
}

// TestDeleteSheetRemovesDefinedNamesReferencingIt 验证删除工作表时，
// 按**表名**引用该表的工作簿级命名区域会被一并清理。
//
// 症状：删掉模板表后，Excel 打开提示「发现不可读取的内容，是否修复」，
// 修复后**静默丢弃**打印区域与重复打印标题 —— 用户视角是「打印设置莫名丢了」。
//
// 根因：reindexDefinedNamesAfterDelete 早期只按 localSheetId 清理，而工作簿级
// 命名区域（无 localSheetId）靠 refersTo 里的表名表达引用，代码从未校验表名是否还在。
// 于是 '月报模板'!$1:$6 成为悬空引用。
func TestDeleteSheetRemovesDefinedNamesReferencingIt(t *testing.T) {
	dir := t.TempDir()
	f := filepath.Join(dir, "book.xlsx")
	b, err := Create()
	if err != nil {
		t.Fatal(err)
	}
	// 前缀关系的表名：验证 "月报" 与 "月报模板" 共存时不会误删
	for _, n := range []string{"月报", "月报模板", "数据"} {
		ws, err := b.NewSheet(n)
		if err != nil {
			t.Fatal(err)
		}
		if err := ws.SetCellValue("A1", n); err != nil {
			t.Fatal(err)
		}
	}
	if err := b.SaveAs(f); err != nil {
		t.Fatal(err)
	}

	// 一组覆盖各种写法的定义区域
	mustSet := map[string]string{
		"引用被删表":    "'月报模板'!$1:$6",
		"引用保留表":    "'月报'!$1:$6",
		"前缀相同不误删":  "'月报模板X'!$1:$6",
		"多表逗号含被删表": "'月报'!$1:$6,'月报模板'!$1:$6",
		"无引号写法":    "数据!$A$1",
		"常量不误删":    "42",
		"函数不误删":    "SUM(Sheet1!A1:A9)",
	}
	for name, ref := range mustSet {
		if err := SetDefinedName(f, name, ref); err != nil {
			t.Fatalf("SetDefinedName(%s): %v", name, err)
		}
	}
	// 表级打印标题行（带 localSheetId，删表本就该清理）
	tpl, err := b.Sheet("月报模板")
	if err != nil {
		t.Fatal(err)
	}
	_ = tpl
	if err := SetDefinedName(f, "_xlnm.Print_Titles", "月报模板!$1:$6"); err != nil {
		t.Fatalf("设置打印标题: %v", err)
	}

	if err := DeleteSheet(f, "月报模板"); err != nil {
		t.Fatalf("DeleteSheet: %v", err)
	}

	wbXML := readZipPart(t, f, "xl/workbook.xml")
	// 被删表必须消失
	if strings.Contains(wbXML, "月报模板") && !strings.Contains(wbXML, "月报模板X") {
		t.Errorf("删除后仍残留对被删表的引用：\n%s", wbXML)
	}
	// 逐条核对：应删的删了，不该删的还在
	assertDefinedNameAbsent(t, wbXML, "引用被删表")
	assertDefinedNameAbsent(t, wbXML, "多表逗号含被删表")
	assertDefinedNameAbsent(t, wbXML, "_xlnm.Print_Titles")
	assertDefinedNamePresent(t, wbXML, "引用保留表")
	assertDefinedNamePresent(t, wbXML, "前缀相同不误删")
	assertDefinedNamePresent(t, wbXML, "无引号写法")
	assertDefinedNamePresent(t, wbXML, "常量不误删")
	assertDefinedNamePresent(t, wbXML, "函数不误删")
}

// TestDefinedNameRefersToSheet 单元测试：表名边界判定。
func TestDefinedNameRefersToSheet(t *testing.T) {
	cases := []struct {
		refersTo string
		sheet    string
		want     bool
	}{
		{"'月报模板'!$1:$6", "月报模板", true},
		{"月报模板!$1:$6", "月报模板", true},
		{"'月报'!$1:$6", "月报模板", false},    // 前缀相同，不算命中
		{"'月报模板X'!$1:$6", "月报模板", false}, // 前缀相同，不算命中
		{"'数据'!A1,'月报模板'!$1", "月报模板", true},
		{"'Sheet1'!A1,数据!A2", "月报模板", false},
		{"42", "月报模板", false},
		{"SUM(Sheet1!A1:A9)", "月报模板", false},
		{"", "月报模板", false},
		{"'月报模板'!$1:$6", "", false}, // 无表名，不判
	}
	for _, c := range cases {
		if got := definedNameRefersToSheet(c.refersTo, c.sheet); got != c.want {
			t.Errorf("definedNameRefersToSheet(%q, %q) = %v, 期望 %v",
				c.refersTo, c.sheet, got, c.want)
		}
	}
}

// assertDefinedNameAbsent 断言 workbook.xml 中不存在名为 name 的 definedName。
func assertDefinedNameAbsent(t *testing.T, wbXML, name string) {
	t.Helper()
	if strings.Contains(wbXML, `name="`+name+`"`) {
		t.Errorf("definedName %q 应已被清理，但仍存在", name)
	}
}

// assertDefinedNamePresent 断言 workbook.xml 中存在名为 name 的 definedName。
func assertDefinedNamePresent(t *testing.T, wbXML, name string) {
	t.Helper()
	if !strings.Contains(wbXML, `name="`+name+`"`) {
		t.Errorf("definedName %q 不应被误删，但已消失", name)
	}
}

// TestInsertSheetDeclaresRelNamespace 验证往 workbook.xml 插入 <sheet> 时，
// 根元素缺 xmlns:r 会被补齐。
//
// 背景：openpyxl 生成的工作簿把 xmlns:r 声明在**每个 <sheet> 元素自己身上**，
// 根元素上没有。这种文件本身合法，但极具脆弱性：
//
//	<workbook xmlns="...main">
//	  <sheet xmlns:r="...relationships" name="配置" sheetId="1" r:id="rId1"/>
//	  <sheet xmlns:r="...relationships" name="月报模板" sheetId="2" r:id="rId2"/>
//	</workbook>
//
// 只要流程把带声明的那些表删光、只剩本库新插入的 <sheet>，r: 前缀就失去声明：
//
//	<workbook xmlns="...main">
//	  <sheet name="2026年1月" sheetId="4" r:id="rId6"/>   ← 前缀无人声明
//	</workbook>
//
// 结果是文件损坏：openpyxl 抛 ParseError(unbound prefix)，
// Excel 提示「发现不可读取的内容」甚至直接打不开。
//
// 精确触发条件（实测）：只 CopySheet 合法；CopySheet + 删部分表仍合法；
// **CopySheet + 把带声明的表全删光才损坏**；源文件根上已有声明则永不触发。
func TestInsertSheetDeclaresRelNamespace(t *testing.T) {
	// openpyxl 风格：声明写在各<sheet> 上，根上没有
	src := xmlDecl +
		`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		`<sheets>` +
		`<sheet xmlns:r="` + relationshipsNSURI + `" name="配置" sheetId="1" r:id="rId1"/>` +
		`</sheets></workbook>`

	got := string(insertSheetInWorkbookXML([]byte(src), "新表", 2, "rId6"))
	root := workbookRootTag(got)

	if !strings.Contains(root, `xmlns:r="`+relationshipsNSURI+`"`) {
		t.Fatalf("根元素未补 xmlns:r 声明，插入的 <sheet r:id> 会成为 unbound prefix：\n%s", root)
	}
	if !strings.Contains(got, `<sheet name="新表" sheetId="2" r:id="rId6"/>`) {
		t.Errorf("新表未正确插入：\n%s", got)
	}
	// 原有的表级声明不应被破坏
	if !strings.Contains(got, `<sheet xmlns:r="`+relationshipsNSURI+`" name="配置"`) {
		t.Errorf("原有 <sheet> 的声明被破坏：\n%s", got)
	}
}

// TestEnsureWorkbookRelNamespace 单元测试该函数的幂等性与边界。
func TestEnsureWorkbookRelNamespace(t *testing.T) {
	decl := ` xmlns:r="` + relationshipsNSURI + `"`
	cases := []struct {
		name string
		in   string
		//wantDecl 为 true 表示根元素应含声明
		wantDecl bool
	}{
		{"根上已有声明-不应重复添加",
			xmlDecl + `<workbook xmlns="` + spreadsheetMainNS + `"` + decl + `><sheets/></workbook>`, true},
		{"根上无声明-应补",
			xmlDecl + `<workbook xmlns="` + spreadsheetMainNS + `"><sheets/></workbook>`, true},
		{"声明只在 sheet 上-根仍需补",
			xmlDecl + `<workbook xmlns="` + spreadsheetMainNS + `"><sheets>` +
				`<sheet xmlns:r="` + relationshipsNSURI + `" name="A" r:id="rId1"/></sheets></workbook>`, true},
		{"非workbook 根-原样返回",
			`<foo><bar/></foo>`, false},
		{"空字符串", "", false},
		{"仅 xml 声明无根",
			xmlDecl, false},
	}
	for _, c := range cases {
		got := ensureWorkbookRelNamespace(c.in)
		if c.wantDecl {
			root := workbookRootTag(got)
			if root == "" {
				t.Errorf("%s: 未找到根元素", c.name)
				continue
			}
			if !strings.Contains(root, `xmlns:r="`+relationshipsNSURI+`"`) {
				t.Errorf("%s: 根元素应有声明，实际 %s", c.name, root)
			}
		}
		// 幂等：再调一次不应变化
		if again := ensureWorkbookRelNamespace(got); again != got {
			t.Errorf("%s: 非幂等，二次调用结果变了", c.name)
		}
		// 绝不能把属性插进 xml 声明里
		if i := strings.Index(got, "<?xml"); i >= 0 {
			declEnd := strings.Index(got[i:], "?>")
			if declEnd > 0 && strings.Contains(got[i:i+declEnd], "xmlns:r") {
				t.Errorf("%s: 属性被插进了 XML 声明：%s", c.name, got[i:i+declEnd])
			}
		}
	}
}

// TestCopySheetIntoWorkbookWithoutRootDeclEndToEnd 端到端复现用户报告的场景：
// openpyxl 风格源文件（声明只在各 <sheet> 上）→ CopySheet 造新表 → 删光带声明的原表
// → 产物必须仍能被 openpyxl 打开。
func TestCopySheetIntoWorkbookWithoutRootDeclEndToEnd(t *testing.T) {
	dir := t.TempDir()
	src := filepath.Join(dir, "src.xlsx")
	// 用 openpyxl 造一个双表工作簿，再把声明从根移到各 <sheet> 上，模拟真实形态
	writeWorkbookWithSheetLevelRelDecl(t, src, []string{"配置", "月报模板"})

	b, err := Open(src)
	if err != nil {
		t.Fatal(err)
	}
	// 前置条件：源文件根上确实没有声明，且带声明的表存在
	srcWB := readZipPart(t, src, "xl/workbook.xml")
	if root := workbookRootTag(srcWB); strings.Contains(root, `xmlns:r=`) {
		t.Fatalf("前置条件不成立：源文件根上已有 xmlns:r，测不到本bug")
	}

	// 同簿复制出两张新表，随后把带声明的原表删光
	for _, name := range []string{"2026年1月", "2026年2月"} {
		if _, err := b.CopySheet("配置", name); err != nil {
			t.Fatalf("CopySheet: %v", err)
		}
	}
	for _, name := range []string{"配置", "月报模板"} {
		if err := b.RemoveSheet(name); err != nil {
			t.Fatalf("RemoveSheet(%s): %v", name, err)
		}
	}
	out := filepath.Join(dir, "out.xlsx")
	if err := b.SaveAs(out); err != nil {
		t.Fatal(err)
	}

	// 产物必须合法：根上有声明，且无声明的 sheet 其r:id 有前缀可依
	outWB := readZipPart(t, out, "xl/workbook.xml")
	root := workbookRootTag(outWB)
	if !strings.Contains(root, `xmlns:r="`+relationshipsNSURI+`"`) {
		t.Fatalf("产物根元素缺 xmlns:r：\n%s", root)
	}
	if strings.Contains(outWB, "月报模板") {
		t.Errorf("已删表名仍残留在 workbook.xml：\n%s", outWB)
	}

	// 用 openpyxl 独立校验：这是最能证明"文件没坏"的判据
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl，跳过加载校验")
	}
	if o, err := runOpenpyxlLoad(t, out); err != nil {
		t.Fatalf("openpyxl 无法加载产物（文件已损坏）：%v\n%s", err, o)
	}
}

// workbookRootTag 提取 workbook.xml 的根元素起始标签（含属性），找不到返回空串。
func workbookRootTag(wbXML string) string {
	start := strings.Index(wbXML, "<workbook")
	if start == -1 {
		return ""
	}
	end := strings.Index(wbXML[start:], ">")
	if end == -1 {
		return ""
	}
	return wbXML[start : start+end+1]
}

// writeWorkbookWithSheetLevelRelDecl 造一份openpyxl 风格的工作簿：
// xmlns:r 声明写在**每个 <sheet> 上**，根元素上没有 —— 这是本bug 的关键前提。
// 表名用中文以贴近真实场景（表名本身与 bug 无关，只是更易复现真实流程）。
func writeWorkbookWithSheetLevelRelDecl(t *testing.T, path string, sheetNames []string) {
	t.Helper()
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	// 先用 openpyxl 造正常工作簿（它把声明写在 <sheet> 上，根上只有 xmlns）
	ops := `{"out": ` + quoteJSON(path) + `, "ops": [`
	for i, n := range sheetNames {
		if i > 0 {
			ops += ","
		}
		ops += `{"op":"create_sheet","name":` + quoteJSON(n) + `}`
	}
	ops += `]}`
	runOpenpyxlOps(t, path, ops)

	// openpyxl 已经在 <sheet> 上写了声明；为保证前提成立，强制确保根上**没有**声明
	data := readZipPart(t, path, "xl/workbook.xml")
	root := workbookRootTag(data)
	if strings.Contains(root, `xmlns:r=`) {
		// 去掉根上的声明，保留各 <sheet> 上的
		stripped := strings.Replace(data, ` xmlns:r="`+relationshipsNSURI+`"`, "", 1)
		// 只替换第一处（根元素那处），确保它出现在 <workbook ...> 内
		if r2 := workbookRootTag(stripped); strings.Contains(r2, `xmlns:r=`) {
			t.Skip("无法构造出「根上无声明」的源文件，跳过本用例")
		}
		data = stripped
	}
	if err := rewriteZipPart(t, path, "xl/workbook.xml", []byte(data)); err != nil {
		t.Fatalf("改写 workbook.xml 失败: %v", err)
	}
}

// rewriteZipPart 重写 zip 中指定部件的内容（其余部件原样复制）。
func rewriteZipPart(t *testing.T, path, part string, content []byte) error {
	t.Helper()
	zr, err := zip.OpenReader(path)
	if err != nil {
		return err
	}
	defer zr.Close()
	buf := new(bytes.Buffer)
	zw := zip.NewWriter(buf)
	for _, f := range zr.File {
		w, err := zw.CreateHeader(&zip.FileHeader{Name: f.Name, Method: zip.Deflate})
		if err != nil {
			return err
		}
		if f.Name == part {
			if _, err := w.Write(content); err != nil {
				return err
			}
			continue
		}
		rc, err := f.Open()
		if err != nil {
			return err
		}
		_, err = io.Copy(w, rc)
		rc.Close()
		if err != nil {
			return err
		}
	}
	if err := zw.Close(); err != nil {
		return err
	}
	return os.WriteFile(path, buf.Bytes(), 0o644)
}

// runOpenpyxlLoad 让 openpyxl 加载指定 xlsx，返回组合输出与错误。
// 用于验证产物是"真能打开"的，而不只是 XML 看起来对。
func runOpenpyxlLoad(t *testing.T, xlsx string) (string, error) {
	t.Helper()
	py := pythonExe()
	if py == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	script := filepath.Join(difftestDir(t), "load_check.py")
	cmd := exec.Command(py, script, xlsx)
	o, err := cmd.CombinedOutput()
	return string(o), err
}
