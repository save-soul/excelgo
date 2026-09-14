package excelgo

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

const gridTestSrc = "../testdata_tmp/grid.xlsx"

func copyGrid(t *testing.T, name string) string {
	t.Helper()
	dst := filepath.Join(t.TempDir(), name)
	data, err := os.ReadFile(gridTestSrc)
	if err != nil {
		t.Fatalf("读测试源文件失败: %v", err)
	}
	if err := os.WriteFile(dst, data, 0644); err != nil {
		t.Fatalf("写临时文件失败: %v", err)
	}
	return dst
}

// mergeExists 用字符串检查 worksheet 中是否存在指定合并区域（不依赖 openpyxl）。
func mergeExists(file, sheetRef, ref string) bool {
	fileMap, err := readZipToMap(file)
	if err != nil {
		return false
	}
	wb := &Workbook{}
	getXMLFromMap(fileMap, "xl/workbook.xml", wb)
	f, _, _, _, err := locateSheetInMap(fileMap, wb, sheetRef)
	if err != nil {
		return false
	}
	ws := string(fileMap[f])
	return strings.Contains(ws, `ref="`+ref+`"`)
}

func TestGetSheetListAndRows(t *testing.T) {
	tmp := copyGrid(t, "list.xlsx")
	list, err := GetSheetList(tmp)
	if err != nil {
		t.Fatalf("GetSheetList: %v", err)
	}
	if len(list) != 1 || list[0] != "网格" {
		t.Fatalf("期望 [网格]，实际 %v", list)
	}
	rows, err := GetRows(tmp, "网格")
	if err != nil {
		t.Fatalf("GetRows: %v", err)
	}
	if len(rows) < 8 {
		t.Fatalf("行数不足，实际 %d", len(rows))
	}
	if rows[0][0] != "销售表" {
		t.Errorf("A1 应为 销售表，实际 %q", rows[0][0])
	}
	if rows[0][3] != "备注" {
		t.Errorf("D1 应为 备注，实际 %q", rows[0][3])
	}
	if rows[2][0] != "苹果" || rows[2][1] != "10" || rows[2][2] != "3.5" {
		t.Errorf("第3行数据异常: %v", rows[2])
	}
	// 合计行第8行存在（公式内容，可能因无缓存结果为空字符串，仅校验单元格存在）
	if rows[7] == nil {
		t.Errorf("第8行(合计)不应缺失")
	}
}

func TestInsertRowsKeepFormat(t *testing.T) {
	tmp := copyGrid(t, "insrow.xlsx")
	if err := InsertRows(tmp, "网格", 3, 2); err != nil {
		t.Fatalf("InsertRows: %v", err)
	}
	rows, _ := GetRows(tmp, "网格")
	// 空行插在 3,4（index 2,3），苹果（原row3）下移到 row5（index 4）
	if rows[2][0] != "" || rows[3][0] != "" {
		t.Errorf("第3,4行应为空行，实际 %q %q", rows[2][0], rows[3][0])
	}
	if rows[4][0] != "苹果" {
		t.Errorf("插入行后苹果应下移到第5行，实际 %q", rows[4][0])
	}
	if rows[1][0] != "名称" {
		t.Errorf("表头应保留，实际 %q", rows[1][0])
	}
	// 合并区域 A10:B10 -> A12:B12（随行平移）
	if !mergeExists(tmp, "网格", "A12:B12") {
		t.Errorf("合并区域应平移为 A12:B12")
	}
}

func TestRemoveRowsKeepFormat(t *testing.T) {
	tmp := copyGrid(t, "delrow.xlsx")
	if err := RemoveRows(tmp, "网格", 3, 1); err != nil {
		t.Fatalf("RemoveRows: %v", err)
	}
	rows, _ := GetRows(tmp, "网格")
	// 删除第3行（苹果），香蕉（原row4）上移到第3行（index 2）
	if rows[2][0] != "香蕉" {
		t.Errorf("删除后香蕉应上移到第3行，实际 %q", rows[2][0])
	}
	if rows[1][0] != "名称" {
		t.Errorf("表头应保留，实际 %q", rows[1][0])
	}
	// 合并区域 A10:B10 -> A9:B9
	if !mergeExists(tmp, "网格", "A9:B9") {
		t.Errorf("合并区域应平移为 A9:B9")
	}
}

func TestInsertColsKeepFormat(t *testing.T) {
	tmp := copyGrid(t, "inscol.xlsx")
	if err := InsertCols(tmp, "网格", 2, 1); err != nil {
		t.Fatalf("InsertCols: %v", err)
	}
	rows, _ := GetRows(tmp, "网格")
	// 原 B 列"数量"右移到 C 列（index 2）
	if rows[1][2] != "数量" {
		t.Errorf("插入列后数量应右移到 C 列，实际 %v", rows[1])
	}
	// 原 A 列"名称"不变（index 0）
	if rows[1][0] != "名称" {
		t.Errorf("A 列应不变，实际 %q", rows[1][0])
	}
	// 原 C 列"单价"右移到 D 列（index 3）
	if rows[1][3] != "单价" {
		t.Errorf("单价应右移到 D 列，实际 %v", rows[1])
	}
	// 合并区域 A1:C1 -> A1:D1（列平移），A10:B10 -> A10:C10
	if !mergeExists(tmp, "网格", "A1:D1") {
		t.Errorf("合并 A1:C1 应平移为 A1:D1")
	}
	if !mergeExists(tmp, "网格", "A10:C10") {
		t.Errorf("合并 A10:B10 应平移为 A10:C10")
	}
}

func TestRemoveColsKeepFormat(t *testing.T) {
	tmp := copyGrid(t, "delcol.xlsx")
	if err := RemoveCols(tmp, "网格", 2, 1); err != nil {
		t.Fatalf("RemoveCols: %v", err)
	}
	rows, _ := GetRows(tmp, "网格")
	// 删除 B 列（数量），原 C 列"单价"左移到 B 列（index 1）
	if rows[1][1] != "单价" {
		t.Errorf("删除列后单价应左移到 B 列，实际 %v", rows[1])
	}
	// 原 A 列"名称"不变
	if rows[1][0] != "名称" {
		t.Errorf("A 列应不变，实际 %q", rows[1][0])
	}
	// 合并区域 A1:C1 -> A1:B1，A10:B10 -> A10:A10（单格，仍含 A10）
	if !mergeExists(tmp, "网格", "A1:B1") {
		t.Errorf("合并 A1:C1 应平移为 A1:B1")
	}
}

// TestVerifyArtifacts 生成固定路径产物供 openpyxl 权威校验（不破坏原格式）。
func TestVerifyArtifacts(t *testing.T) {
	dir := "../testdata_tmp"
	cases := []struct {
		name string
		fn   func(string, string) error
	}{
		{"verify_insrow.xlsx", func(f, s string) error { return InsertRows(f, s, 3, 2) }},
		{"verify_delrow.xlsx", func(f, s string) error { return RemoveRows(f, s, 3, 1) }},
		{"verify_inscol.xlsx", func(f, s string) error { return InsertCols(f, s, 2, 1) }},
		{"verify_delcol.xlsx", func(f, s string) error { return RemoveCols(f, s, 2, 1) }},
	}
	for _, c := range cases {
		dst := filepath.Join(dir, c.name)
		data, err := os.ReadFile(gridTestSrc)
		if err != nil {
			t.Fatalf("read src: %v", err)
		}
		if err := os.WriteFile(dst, data, 0644); err != nil {
			t.Fatalf("write: %v", err)
		}
		if err := c.fn(dst, "网格"); err != nil {
			t.Fatalf("%s: %v", c.name, err)
		}
	}
}
