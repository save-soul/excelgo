package excelgo

import (
	"os"
	"path/filepath"
	"testing"
)

// TestObjectAPI 验证「文件 -> 工作表 -> 单元格」面向对象链式 API 可用且不破坏数据。
func TestObjectAPI(t *testing.T) {
	src := filepath.Join("..", "testdata_tmp", "full.xlsx")
	dst := filepath.Join("..", "testdata_tmp", "obj_api.xlsx")
	copyFileForTest(t, src, dst)

	b, err := Open(dst)
	if err != nil {
		t.Fatalf("Open: %v", err)
	}
	// 新建工作表
	if _, err := b.NewSheet("Obj"); err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	s, err := b.Sheet("Obj")
	if err != nil {
		t.Fatalf("Sheet: %v", err)
	}
	// 链式：单元格写入 + 公式 + 样式 + 行列操作
	if err := s.Cell("A1").SetStr("分包队"); err != nil {
		t.Fatalf("SetStr: %v", err)
	}
	if err := s.Cell("B1").SetStr("进度"); err != nil {
		t.Fatalf("SetStr: %v", err)
	}
	if err := s.Cell("A2").Set(75); err != nil {
		t.Fatalf("Set(75): %v", err)
	}
	if err := s.Cell("B2").SetFormula("A2*2", "150"); err != nil {
		t.Fatalf("SetFormula: %v", err)
	}
	st := Style{Font: &FontStyle{Bold: true}, Fill: &FillStyle{Color: "FFFFFF00", PatternType: "solid"}}
	if _, err := s.Cell("A1").SetStyle(st); err != nil {
		t.Fatalf("SetStyle: %v", err)
	}
	// 区域写入
	if err := s.Range("A4:C5").Set([][]interface{}{
		{"a", "b", "c"},
		{"1", "2", "3"},
	}); err != nil {
		t.Fatalf("Range.Set: %v", err)
	}
	// 行列插入（公式引用应自动平移）：在第 2 行前插入，A2/B2 下移到 A3/B3
	if err := s.InsertRows(2, 1); err != nil {
		t.Fatalf("InsertRows: %v", err)
	}
	// 合并
	if err := s.MergeCells("A1:B1"); err != nil {
		t.Fatalf("MergeCells: %v", err)
	}
	// 冻结 + 列宽
	if err := s.FreezePanes("A2"); err != nil {
		t.Fatalf("Freeze: %v", err)
	}
	if err := s.SetColWidth(1, 20); err != nil {
		t.Fatalf("SetColWidth: %v", err)
	}
	if err := b.Save(); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// 校验：重新打开，确认数据/样式/公式平移正确
	b2, err := Open(dst)
	if err != nil {
		t.Fatalf("re-Open: %v", err)
	}
	s2, err := b2.Sheet("Obj")
	if err != nil {
		t.Fatalf("re-Sheet: %v", err)
	}
	if v := s2.Cell("A1").GetStr(); v != "分包队" {
		t.Errorf("A1 = %q, want 分包队", v)
	}
	// 在第 2 行前插入后，原 A2/B2 下移到 A3/B3，且公式引用随平移 A2->A3
	if v := s2.Cell("A3").GetNum(); v != 75 {
		t.Errorf("A3 (原A2数值) = %v, want 75", v)
	}
	f := s2.Cell("B3").GetFormula()
	if f != "A3*2" {
		t.Errorf("B3 公式 = %q, want A3*2（引用随插入行平移）", f)
	}
	if _, err := s2.Cell("A1").GetStyle(); err != nil {
		t.Errorf("GetStyle A1: %v", err)
	}
}

// copyFileForTest 复制源文件到目标（测试用）。
func copyFileForTest(t *testing.T, src, dst string) {
	t.Helper()
	data, err := os.ReadFile(src)
	if err != nil {
		t.Fatalf("read %s: %v", src, err)
	}
	if err := os.WriteFile(dst, data, 0644); err != nil {
		t.Fatalf("write %s: %v", dst, err)
	}
}
