package excelgo

import (
	"regexp"
	"strings"
	"testing"
)

// 准备一个用于特性测试的临时工作簿（基于现有 full.xlsx 复制，避免依赖创建新簿的 API）。
func setupFeatureWB(t *testing.T) string {
	t.Helper()
	src := "../testdata_tmp/full.xlsx"
	dst := t.TempDir() + "/feat.xlsx"
	bf, err := Open(src)
	if err != nil {
		t.Fatalf("准备测试簿失败: %v", err)
	}
	if err := bf.CopySheetTo(dst, "Sheet1", "", WithSuffix("_base")); err != nil {
		t.Fatalf("准备测试簿失败: %v", err)
	}
	// 新建一个干净的工作表用于写入
	if _, err := NewSheet(dst, "Feat"); err != nil {
		t.Fatalf("新建 Feat 失败: %v", err)
	}
	return dst
}

// readCellFormula 读取 worksheet 中指定单元格的 <f> 公式文本（用于断言引用平移）。
func readCellFormula(t *testing.T, file, sheet, cell string) string {
	t.Helper()
	fm, err := readZipToMap(file)
	if err != nil {
		t.Fatalf("读盘失败: %v", err)
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fm, "xl/workbook.xml", wb); err != nil {
		t.Fatalf("读 workbook 失败: %v", err)
	}
	f, _, _, _, err := locateSheetInMap(fm, wb, sheet)
	if err != nil {
		t.Fatalf("定位表失败: %v", err)
	}
	ws := string(fm[f])
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + regexp.QuoteMeta(cell) + `"[^>]*?>(.*?)</c>`)
	cm := re.FindStringSubmatch(ws)
	if cm == nil {
		return ""
	}
	fmRe := regexp.MustCompile(`(?s)<f[^>]*>(.*?)</f>`)
	if m := fmRe.FindStringSubmatch(cm[1]); m != nil {
		return m[1]
	}
	return ""
}

// readCellStyleIdx 读取单元格 s 属性（样式索引）。
func readCellStyleIdx(t *testing.T, file, sheet, cell string) string {
	t.Helper()
	fm, err := readZipToMap(file)
	if err != nil {
		t.Fatalf("读盘失败: %v", err)
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fm, "xl/workbook.xml", wb); err != nil {
		t.Fatalf("读 workbook 失败: %v", err)
	}
	f, _, _, _, err := locateSheetInMap(fm, wb, sheet)
	if err != nil {
		t.Fatalf("定位表失败: %v", err)
	}
	ws := string(fm[f])
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + regexp.QuoteMeta(cell) + `"[^>]*?>`)
	if m := re.FindString(ws); m != "" {
		return attrOf(m, "s")
	}
	return ""
}

func TestFormulaShiftOnInsertRows(t *testing.T) {
	file := setupFeatureWB(t)
	if err := SetCellValue(file, "Feat", "A1", 10); err != nil {
		t.Fatal(err)
	}
	if err := SetCellValue(file, "Feat", "A2", 20); err != nil {
		t.Fatal(err)
	}
	// B1 引用 A1+A2，并引用一个绝对列 $C1
	if err := SetCellFormula(file, "Feat", "B1", "A1+A2+$C1", "30"); err != nil {
		t.Fatal(err)
	}
	// 在第 1 行之前插入 1 行：B1 -> B2；引用 A1->A2, A2->A3, $C1 列绝对、行相对 -> $C2
	if err := InsertRows(file, "Feat", 1, 1); err != nil {
		t.Fatal(err)
	}
	got := readCellFormula(t, file, "Feat", "B2")
	want := "A2+A3+$C2"
	if got != want {
		t.Fatalf("插入行后公式应平移为 %q，实际 %q", want, got)
	}
}

func TestFormulaShiftOnRemoveRows(t *testing.T) {
	file := setupFeatureWB(t)
	if err := SetCellFormula(file, "Feat", "B5", "A2+A4", "0"); err != nil {
		t.Fatal(err)
	}
	// 删除第 2~3 行：A2 落入删除区间 -> #REF!；A4 行>=4 -> A2
	if err := RemoveRows(file, "Feat", 2, 2); err != nil {
		t.Fatal(err)
	}
	got := readCellFormula(t, file, "Feat", "B3") // 原 B5 上移到 B3
	if !strings.Contains(got, "#REF!") || !strings.Contains(got, "A2") {
		t.Fatalf("删除行后公式应含 #REF! 与 A2，实际 %q", got)
	}
}

func TestFormulaShiftOnInsertCols(t *testing.T) {
	file := setupFeatureWB(t)
	if err := SetCellFormula(file, "Feat", "C1", "A1+B1+$A$1", "0"); err != nil {
		t.Fatal(err)
	}
	// 在第 1 列前插入 1 列：C1 -> D1；A1->B1, B1->C1, $A$1 绝对列不变
	if err := InsertCols(file, "Feat", 1, 1); err != nil {
		t.Fatal(err)
	}
	got := readCellFormula(t, file, "Feat", "D1")
	want := "B1+C1+$A$1"
	if got != want {
		t.Fatalf("插入列后公式应平移为 %q，实际 %q", want, got)
	}
}

func TestFormulaShiftOnRemoveCols(t *testing.T) {
	file := setupFeatureWB(t)
	if err := SetCellFormula(file, "Feat", "D5", "A2+B2+C2", "0"); err != nil {
		t.Fatal(err)
	}
	// 删除第 2~3 列：B2 落入删除区间 -> #REF!；A2 列<2 不变；C2 列>=4 -> B2
	if err := RemoveCols(file, "Feat", 2, 2); err != nil {
		t.Fatal(err)
	}
	got := readCellFormula(t, file, "Feat", "B5") // 原 D5 左移到 B5
	if !strings.Contains(got, "#REF!") || !strings.Contains(got, "A2") {
		t.Fatalf("删除列后公式应含 #REF! 与 A2，实际 %q", got)
	}
}

func TestRangeReadWrite(t *testing.T) {
	file := setupFeatureWB(t)
	// 写入 2x3 区域
	vals := [][]interface{}{
		{"a", 1, 2.5},
		{"b", 3, 4.5},
	}
	if err := SetRange(file, "Feat", "A1", vals); err != nil {
		t.Fatal(err)
	}
	grid, err := GetRange(file, "Feat", "A1:C2")
	if err != nil {
		t.Fatal(err)
	}
	if len(grid) != 2 || len(grid[0]) != 3 {
		t.Fatalf("区域尺寸应为 2x3，实际 %dx%d", len(grid), len(grid[0]))
	}
	if grid[0][0] != "a" || grid[0][1] != "1" || grid[1][2] != "4.5" {
		t.Fatalf("区域读取内容不符: %v", grid)
	}
	// 公式单元格写入（A5:B5 为一行两列）
	if err := SetRange(file, "Feat", "A5:B5", [][]interface{}{
		{CellFormula{Formula: "A1+B1"}, CellFormula{Formula: "A2*B2", Result: "15"}},
	}); err != nil {
		t.Fatal(err)
	}
	if got := readCellFormula(t, file, "Feat", "A5"); got != "A1+B1" {
		t.Fatalf("A5 公式应为 A1+B1，实际 %q", got)
	}
}

func TestSetCellStyleAPI(t *testing.T) {
	file := setupFeatureWB(t)
	style := Style{
		Font:      &FontStyle{Bold: true, Color: "FFFF0000", Size: 14, Name: "微软雅黑"},
		Fill:      &FillStyle{Color: "FFFFFF00"},
		Alignment: &AlignmentStyle{Horizontal: "center", Vertical: "center", WrapText: true},
		NumFmt:    "0.00",
	}
	idx, err := SetCellStyleRange(file, "Feat", "A1:C3", style)
	if err != nil {
		t.Fatal(err)
	}
	if idx < 1 {
		t.Fatalf("样式索引应 >=1，实际 %d", idx)
	}
	for _, c := range []string{"A1", "B2", "C3"} {
		if got := readCellStyleIdx(t, file, "Feat", c); got != itoa(idx) {
			t.Fatalf("单元格 %s 的 s 应为 %d，实际 %q", c, idx, got)
		}
	}
	// 同样式再次应用应复用同一索引（去重）
	idx2, err := SetCellStyleRange(file, "Feat", "D4", style)
	if err != nil {
		t.Fatal(err)
	}
	if idx2 != idx {
		t.Fatalf("相同样式应复用索引 %d，实际 %d", idx, idx2)
	}
}
