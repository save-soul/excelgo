package excelgo

import (
	"strings"
	"testing"
)

// TestCoordsAPI 覆盖 excelize 六个坐标转换函数的等价行为。
func TestCoordsAPI(t *testing.T) {
	cases := []struct {
		name string
		got  string
		want string
	}{
		// SplitCellName
		{"SplitCellName(AK74)", must2(SplitCellName("AK74")), "AK,74"},
		{"SplitCellName(a1)", must2(SplitCellName("a1")), "a,1"},
		// JoinCellName
		{"JoinCellName(AK,74)", must1(JoinCellName("AK", 74)), "AK74"},
		{"JoinCellName(a,1)", must1(JoinCellName("a", 1)), "A1"},
		// ColumnNameToNumber
		{"ColumnNameToNumber(AK)", mustInt(ColumnNameToNumber("AK")), "37"},
		{"ColumnNameToNumber(a)", mustInt(ColumnNameToNumber("a")), "1"},
		{"ColumnNameToNumber(AA)", mustInt(ColumnNameToNumber("AA")), "27"},
		{"ColumnNameToNumber(XFD)", mustInt(ColumnNameToNumber("XFD")), "16384"},
		// ColumnNumberToName
		{"ColumnNumberToName(37)", must1(ColumnNumberToName(37)), "AK"},
		{"ColumnNumberToName(1)", must1(ColumnNumberToName(1)), "A"},
		{"ColumnNumberToName(27)", must1(ColumnNumberToName(27)), "AA"},
		{"ColumnNumberToName(16384)", must1(ColumnNumberToName(16384)), "XFD"},
		// CellNameToCoordinates
		{"CellNameToCoordinates(A1)", must2int(CellNameToCoordinates("A1")), "1,1"},
		{"CellNameToCoordinates(Z3)", must2int(CellNameToCoordinates("Z3")), "26,3"},
		// CoordinatesToCellName
		{"CoordinatesToCellName(1,1)", must1(CoordinatesToCellName(1, 1)), "A1"},
		{"CoordinatesToCellName(1,1,true)", must1(CoordinatesToCellName(1, 1, true)), "$A$1"},
		{"CoordinatesToCellName(37,74)", must1(CoordinatesToCellName(37, 74)), "AK74"},
	}
	for _, c := range cases {
		if c.got != c.want {
			t.Errorf("%s = %s, want %s", c.name, c.got, c.want)
		}
	}
}

// TestCoordsRoundTrip 往返一致性：坐标 → 索引 → 坐标。
func TestCoordsRoundTrip(t *testing.T) {
	for _, cell := range []string{"A1", "Z9", "AA10", "AK74", "XFD1048576"} {
		col, row, err := CellNameToCoordinates(cell)
		if err != nil {
			t.Fatalf("%s: %v", cell, err)
		}
		back, err := CoordinatesToCellName(col, row)
		if err != nil {
			t.Fatalf("%s: %v", cell, err)
		}
		if back != cell {
			t.Errorf("往返失败: %s -> (%d,%d) -> %s", cell, col, row, back)
		}
	}
}

// TestCoordsErrors 非法输入必须报错而非静默返回零值。
func TestCoordsErrors(t *testing.T) {
	if _, _, err := SplitCellName("1A"); err == nil {
		t.Error("SplitCellName(\"1A\") 应报错")
	}
	if _, _, err := SplitCellName("A"); err == nil {
		t.Error("SplitCellName(\"A\") 应报错")
	}
	if _, _, err := SplitCellName("A0"); err == nil {
		t.Error("SplitCellName(\"A0\") 行号 0 应报错")
	}
	if _, _, err := SplitCellName("A1048577"); err == nil {
		t.Error("SplitCellName 超出最大行号应报错")
	}
	if _, err := ColumnNameToNumber("A1"); err == nil {
		t.Error("ColumnNameToNumber(\"A1\") 应报错")
	}
	if _, err := ColumnNameToNumber("XFE"); err == nil {
		t.Error("ColumnNameToNumber(\"XFE\") 超出最大列应报错")
	}
	if _, err := JoinCellName("A", 0); err == nil {
		t.Error("JoinCellName 行号 0 应报错")
	}
	if _, err := ColumnNumberToName(0); err == nil {
		t.Error("ColumnNumberToName(0) 应报错")
	}
	if _, err := ColumnNumberToName(16385); err == nil {
		t.Error("ColumnNumberToName(16385) 应报错")
	}
	if _, err := CoordinatesToCellName(1, 0); err == nil {
		t.Error("CoordinatesToCellName 行号 0 应报错")
	}
}

func must1(s string, err error) string {
	if err != nil {
		return "ERR:" + err.Error()
	}
	return s
}

func mustInt(i int, err error) string {
	if err != nil {
		return "ERR:" + err.Error()
	}
	return itoa(i)
}

func must2int(a, b int, err error) string {
	if err != nil {
		return "ERR:" + err.Error()
	}
	return itoa(a) + "," + itoa(b)
}

func must2(s string, i int, err error) string {
	if err != nil {
		return "ERR:" + err.Error()
	}
	return s + "," + itoa(i)
}

// TestUpdateLinkedValue 验证公式缓存被清除、普通数值单元格不受影响。
func TestUpdateLinkedValue(t *testing.T) {
	p := genPath("07_linked_value.xlsx")
	b0, err := Create()
	if err != nil {
		t.Fatal(err)
	}
	if err := b0.SaveAs(p); err != nil {
		t.Fatal(err)
	}

	f, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	sh, err := f.Sheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	if err := sh.SetCellFormula("B19", "SUM(Sheet2!D2,Sheet2!D11)", "100"); err != nil {
		t.Fatal(err)
	}
	if err := sh.SetCellValue("C5", 42); err != nil {
		t.Fatal(err)
	}
	if err := f.UpdateLinkedValue(); err != nil {
		t.Fatalf("UpdateLinkedValue: %v", err)
	}
	if err := f.Save(); err != nil {
		t.Fatal(err)
	}

	fm := readMap(t, p)
	ws := string(fm["xl/worksheets/sheet1.xml"])
	t.Logf("sheet1.xml = %s", ws)

	// 公式单元格：<f> 保留，<v> 清除
	if !strings.Contains(ws, "<f>SUM(Sheet2!D2,Sheet2!D11)</f>") {
		t.Error("公式文本 <f> 应保留")
	}
	if strings.Contains(ws, "<v>100</v>") {
		t.Error("公式缓存 <v>100</v> 应被清除")
	}
	// 普通数值单元格：<v> 必须保留（这是数据，不是缓存）
	if !strings.Contains(ws, "<v>42</v>") {
		t.Errorf("非公式单元格的 <v>42</v> 不应被清除，实际 XML: %s", ws)
	}
}

// TestUpdateLinkedValueNotGreedy 关键回归：清除公式缓存不得跨单元格贪婪匹配，
// 把公式单元格之后的普通数值单元格的 <v> 一并吃掉。
func TestUpdateLinkedValueNotGreedy(t *testing.T) {
	src := `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>` +
		`<row r="19"><c r="B19"><f>SUM(A1:A2)</f><v>100</v></c><c r="C19"><v>7</v></c></row>` +
		`</sheetData></worksheet>`
	out, cleared := clearFormulaCache(src)
	if cleared != 1 {
		t.Errorf("应只清除 1 个缓存，实际 %d", cleared)
	}
	if !strings.Contains(out, "<c r=\"B19\"><f>SUM(A1:A2)</f></c>") {
		t.Errorf("公式单元格应变为无 <v>，实际: %s", out)
	}
	if !strings.Contains(out, `<c r="C19"><v>7</v></c>`) {
		t.Errorf("紧随其后的普通数值单元格必须完好，实际: %s", out)
	}
}

// TestUpdateLinkedValueMultiSheet 所有工作表都要处理。
func TestUpdateLinkedValueMultiSheet(t *testing.T) {
	p := genPath("07_linked_multi.xlsx")
	b0, _ := Create()
	if err := b0.SaveAs(p); err != nil {
		t.Fatal(err)
	}
	f, _ := Open(p)
	if _, err := f.AddSheet("S2"); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"Sheet1", "S2"} {
		sh, _ := f.Sheet(name)
		if err := sh.SetCellFormula("A1", "1+1", "2"); err != nil {
			t.Fatal(err)
		}
	}
	if err := f.UpdateLinkedValue(); err != nil {
		t.Fatal(err)
	}
	if err := f.Save(); err != nil {
		t.Fatal(err)
	}
	fm := readMap(t, p)
	for k, v := range fm {
		if isWorksheetPart(k) && strings.Contains(string(v), "<v>2</v>") {
			t.Errorf("%s 的公式缓存未被清除", k)
		}
	}
}

// TestUpdateLinkedValueIdempotent 重复调用应稳定，不产生副作用。
func TestUpdateLinkedValueIdempotent(t *testing.T) {
	p := genPath("07_linked_idem.xlsx")
	b0, _ := Create()
	_ = b0.SaveAs(p)
	f, _ := Open(p)
	sh, _ := f.Sheet("Sheet1")
	_ = sh.SetCellFormula("A1", "1+1", "2")
	_ = f.Save()

	f2, _ := Open(p)
	if err := f2.UpdateLinkedValue(); err != nil {
		t.Fatal(err)
	}
	_ = f2.Save()
	fm1 := readMap(t, p)

	f3, _ := Open(p)
	if err := f3.UpdateLinkedValue(); err != nil {
		t.Fatal(err)
	}
	_ = f3.Save()
	fm2 := readMap(t, p)

	if string(fm1["xl/worksheets/sheet1.xml"]) != string(fm2["xl/worksheets/sheet1.xml"]) {
		t.Error("重复调用 UpdateLinkedValue 结果不一致（非幂等）")
	}
}

// BenchmarkCoords 坐标转换不应成为批量操作的瓶颈。
func BenchmarkCoords(b *testing.B) {
	for i := 0; i < b.N; i++ {
		if _, _, err := CellNameToCoordinates("AK74"); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkSplitCellName(b *testing.B) {
	for i := 0; i < b.N; i++ {
		if _, _, err := SplitCellName("BC123"); err != nil {
			b.Fatal(err)
		}
	}
}
