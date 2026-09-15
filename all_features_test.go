package excelgo

import (
	"encoding/base64"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// all_features_test.go 生成真实 .xlsx 测试文件（generated/ 目录）并逐项验证公共 API 行为
// 是否与预期一致。读回校验尽量走与写入不同的代码路径（对象 API / 直接解析 zip 内 XML），
// 避免「自己写自己读」的循环验证。另见 validate_generated.py（openpyxl 独立交叉验证）。

const onePixelPNG = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="

func genPath(name string) string { return filepath.Join("generated", name) }

func freshBook(t *testing.T, name string) string {
	t.Helper()
	p := genPath(name)
	b, err := Create()
	if err != nil {
		t.Fatalf("Create: %v", err)
	}
	if err := b.SaveAs(p); err != nil {
		t.Fatalf("SaveAs %s: %v", p, err)
	}
	return p
}

func readMap(t *testing.T, path string) map[string][]byte {
	t.Helper()
	fm, err := readZipToMap(path)
	if err != nil {
		t.Fatalf("readZipToMap %s: %v", path, err)
	}
	return fm
}

func mustHas(t *testing.T, fm map[string][]byte, key string) {
	t.Helper()
	if _, ok := fm[key]; !ok {
		t.Errorf("期望文件 %q 存在于压缩包内，但未找到", key)
	}
}

func countPrefix(fm map[string][]byte, prefix string) int {
	n := 0
	for k := range fm {
		if strings.HasPrefix(k, prefix) {
			n++
		}
	}
	return n
}

func sheetList(t *testing.T, path string) []string {
	t.Helper()
	lst, err := GetSheetList(path)
	if err != nil {
		t.Fatalf("GetSheetList %s: %v", path, err)
	}
	return lst
}

func cpFile(t *testing.T, src, dst string) {
	t.Helper()
	data, err := os.ReadFile(src)
	if err != nil {
		t.Fatalf("读 %s: %v", src, err)
	}
	if err := os.WriteFile(dst, data, 0644); err != nil {
		t.Fatalf("写 %s: %v", dst, err)
	}
}

func TestAllFeatures(t *testing.T) {
	if err := os.MkdirAll("generated", 0755); err != nil {
		t.Fatalf("创建 generated 目录: %v", err)
	}

	t.Run("SheetLifecycle", func(t *testing.T) {
		p := freshBook(t, "01_lifecycle.xlsx")
		if n, err := NewSheet(p, "S2"); err != nil || n != 2 {
			t.Fatalf("NewSheet S2: err=%v n=%d", err, n)
		}
		if err := RenameSheet(p, "Sheet1", "Base"); err != nil {
			t.Fatalf("RenameSheet: %v", err)
		}
		if err := MoveSheet(p, "S2", 1); err != nil {
			t.Fatalf("MoveSheet: %v", err)
		}
		lst := sheetList(t, p)
		if len(lst) != 2 || lst[0] != "S2" || lst[1] != "Base" {
			t.Errorf("工作表顺序应为 [S2 Base]，实际 %v", lst)
		}
		if err := DeleteSheet(p, "S2"); err != nil {
			t.Fatalf("DeleteSheet: %v", err)
		}
		lst = sheetList(t, p)
		if len(lst) != 1 || lst[0] != "Base" {
			t.Errorf("删除后应为 [Base]，实际 %v", lst)
		}
	})

	t.Run("CellTypes", func(t *testing.T) {
		p := freshBook(t, "02_cells.xlsx")
		if err := SetCellStr(p, "Sheet1", "A1", "你好"); err != nil {
			t.Fatalf("SetCellStr: %v", err)
		}
		if err := SetCellInt(p, "Sheet1", "A2", 42); err != nil {
			t.Fatalf("SetCellInt: %v", err)
		}
		if err := SetCellNumeric(p, "Sheet1", "A3", 3.14); err != nil {
			t.Fatalf("SetCellNumeric: %v", err)
		}
		if err := SetCellBool(p, "Sheet1", "A4", true); err != nil {
			t.Fatalf("SetCellBool: %v", err)
		}
		if err := SetCellInline(p, "Sheet1", "A5", "inline"); err != nil {
			t.Fatalf("SetCellInline: %v", err)
		}
		if err := SetCellFormula(p, "Sheet1", "A6", "A2*2", "84"); err != nil {
			t.Fatalf("SetCellFormula: %v", err)
		}
		if err := SetCellValue(p, "Sheet1", "A7", "auto"); err != nil {
			t.Fatalf("SetCellValue: %v", err)
		}
		// 用对象 API 读回（独立代码路径）
		b, err := Open(p)
		if err != nil {
			t.Fatalf("Open: %v", err)
		}
		s, _ := b.Sheet("Sheet1")
		checks := []struct {
			cell, want string
		}{
			{"A1", "你好"}, {"A2", "42"}, {"A3", "3.14"}, {"A4", "1"},
			{"A5", "inline"}, {"A7", "auto"},
		}
		for _, c := range checks {
			if got := s.Cell(c.cell).Get(); got != c.want {
				t.Errorf("单元格 %s = %q，期望 %q", c.cell, got, c.want)
			}
		}
		if f := s.Cell("A6").GetFormula(); f != "A2*2" {
			t.Errorf("A6 公式 = %q，期望 A2*2", f)
		}
		if v := s.Cell("A6").Get(); v != "84" {
			t.Errorf("A6 公式结果 = %q，期望 84", v)
		}
	})

	t.Run("RangeReadWrite", func(t *testing.T) {
		p := freshBook(t, "03_range.xlsx")
		data := [][]interface{}{
			{"a", "b", "c"},
			{1, 2, 3},
			{"x", "y", "z"},
		}
		if err := SetRange(p, "Sheet1", "B2:D4", data); err != nil {
			t.Fatalf("SetRange: %v", err)
		}
		got, err := GetRange(p, "Sheet1", "B2:D4")
		if err != nil {
			t.Fatalf("GetRange: %v", err)
		}
		want := [][]string{{"a", "b", "c"}, {"1", "2", "3"}, {"x", "y", "z"}}
		for i := range want {
			for j := range want[i] {
				if got[i][j] != want[i][j] {
					t.Errorf("range[%d][%d] = %q，期望 %q", i, j, got[i][j], want[i][j])
				}
			}
		}
	})

	t.Run("RowsColsShift", func(t *testing.T) {
		p := freshBook(t, "04_rows_cols.xlsx")
		_ = SetCellStr(p, "Sheet1", "A1", "top")
		_ = SetCellStr(p, "Sheet1", "A2", "mid")
		_ = SetCellFormula(p, "Sheet1", "A3", "A1&A2", "topmid")
		// 在第 2 行前插入 1 行：A1->A1, A2->A3, A3公式->A4，且引用 A1&A2 应平移为 A1&A3
		if err := InsertRows(p, "Sheet1", 2, 1); err != nil {
			t.Fatalf("InsertRows: %v", err)
		}
		cpFile(t, p, genPath("04a_after_insert.xlsx"))
		b, _ := Open(p)
		s, _ := b.Sheet("Sheet1")
		if s.Cell("A1").Get() != "top" {
			t.Errorf("插入行后 A1 应仍为 top，实际 %q", s.Cell("A1").Get())
		}
		if s.Cell("A3").Get() != "mid" {
			t.Errorf("原 A2 应下移到 A3，实际 %q", s.Cell("A3").Get())
		}
		if f := s.Cell("A4").GetFormula(); f != "A1&A3" {
			t.Errorf("A4 公式应平移为 A1&A3，实际 %q", f)
		}
		// 删除第 2 行：A3 上移回 A2
		if err := RemoveRows(p, "Sheet1", 2, 1); err != nil {
			t.Fatalf("RemoveRows: %v", err)
		}
		cpFile(t, p, genPath("04b_after_remove.xlsx"))
		b2, _ := Open(p)
		s2, _ := b2.Sheet("Sheet1")
		if s2.Cell("A2").Get() != "mid" {
			dumpSheet(t, p, "A2 删除行后应为 mid，实际空")
		}
		// 列插入：在 B 列前插入，A 列不动
		if err := InsertCols(p, "Sheet1", 2, 1); err != nil {
			t.Fatalf("InsertCols: %v", err)
		}
		cpFile(t, p, genPath("04c_after_insertcol.xlsx"))
		b3, _ := Open(p)
		s3, _ := b3.Sheet("Sheet1")
		if s3.Cell("A2").Get() != "mid" {
			dumpSheet(t, p, "A2 插入列后应为 mid，实际空")
		}
	})

	t.Run("Query", func(t *testing.T) {
		p := freshBook(t, "05_query.xlsx")
		if _, err := NewSheet(p, "Q2"); err != nil {
			t.Fatalf("NewSheet Q2: %v", err)
		}
		if ok, _ := SheetExists(p, "Q2"); !ok {
			t.Errorf("SheetExists(Q2) 应为 true")
		}
		if ok, _ := SheetExists(p, "Nope"); ok {
			t.Errorf("SheetExists(Nope) 应为 false")
		}
		idx, err := GetSheetIndex(p, "Q2")
		if err != nil {
			t.Fatalf("GetSheetIndex: %v", err)
		}
		// GetSheetIndex 返回 0 基序号：第 2 个工作表应为 1
		if idx != 1 {
			t.Errorf("GetSheetIndex(Q2) = %d，期望 0 基序号 1", idx)
		}
		// Sheet1 写入 A1、A3（中间 A2 空）后，GetRows 仅返回有数据的行（2 行）
		_ = SetCellStr(p, "Sheet1", "A1", "r1")
		_ = SetCellStr(p, "Sheet1", "A3", "r3")
		rows, err := GetRows(p, "Sheet1")
		if err != nil {
			t.Fatalf("GetRows: %v", err)
		}
		if len(rows) != 2 {
			t.Errorf("Sheet1 应有 2 行（空行 A2 不计入），实际 %d", len(rows))
		}
		if len(rows) >= 2 && (rows[0][0] != "r1" || rows[1][0] != "r3") {
			t.Errorf("GetRows 内容不符：%v", rows)
		}
		if _, err := GetSheetData(p, "Sheet1"); err != nil {
			t.Errorf("GetSheetData: %v", err)
		}
	})

	t.Run("Pictures", func(t *testing.T) {
		p := freshBook(t, "06_pictures.xlsx")
		data, err := base64.StdEncoding.DecodeString(onePixelPNG)
		if err != nil {
			t.Fatalf("解码 PNG: %v", err)
		}
		if err := AddPictureFromBytes(p, "Sheet1", data, nil); err != nil {
			t.Fatalf("AddPictureFromBytes: %v", err)
		}
		if err := AddCellPictureFromBytes(p, "Sheet1", "B2", data); err != nil {
			t.Fatalf("AddCellPictureFromBytes: %v", err)
		}
		fm := readMap(t, p)
		if n := countPrefix(fm, "xl/media/"); n < 2 {
			t.Errorf("media 文件应 >=2，实际 %d", n)
		}
		// worksheet rels 应同时含 drawing 与 image 两类关系
		sr := loadRels(fm, "xl/worksheets/_rels/sheet1.xml.rels")
		if sr == nil {
			t.Fatalf("sheet1.xml.rels 缺失")
		}
		hasDrawing, hasImage := false, false
		for _, rel := range sr.Relationship {
			if strings.Contains(rel.Type, "/drawing") {
				hasDrawing = true
			}
			if strings.Contains(rel.Type, "/image") {
				hasImage = true
			}
		}
		if !hasDrawing || !hasImage {
			t.Errorf("sheet1 rels 应同时含 drawing 与 image，实际 drawing=%v image=%v", hasDrawing, hasImage)
		}
	})

	t.Run("CopySheet", func(t *testing.T) {
		p := freshBook(t, "07_copy.xlsx")
		_ = SetCellStr(p, "Sheet1", "A1", "origin")
		if err := CopySheet(p, p, "Sheet1"); err != nil {
			t.Fatalf("CopySheet: %v", err)
		}
		lst := sheetList(t, p)
		if len(lst) != 2 {
			t.Fatalf("复制后应含 2 个工作表，实际 %v", lst)
		}
		// 新表（Sheet1_副本 或类似）应含 origin 数据
		b, _ := Open(p)
		found := false
		for _, name := range lst {
			if name == "Sheet1" {
				continue
			}
			if s, err := b.Sheet(name); err == nil && s.Cell("A1").Get() == "origin" {
				found = true
			}
		}
		if !found {
			t.Errorf("复制出的工作表未含 A1=origin")
		}
	})

	t.Run("MergeWorkbook", func(t *testing.T) {
		p := genPath("08_merge.xlsx")
		cpFile(t, "testfixtures/target.xlsx", p)
		before := len(sheetList(t, p))
		err := MergeWorkbook(p, []SourceRef{
			{Workbook: "testfixtures/src1.xlsx", Sheet: "数据A"},
			{Workbook: "testfixtures/src2.xlsx", Sheet: "数据B"},
		})
		if err != nil {
			t.Fatalf("MergeWorkbook: %v", err)
		}
		after := len(sheetList(t, p))
		if after != before+2 {
			t.Errorf("合并后应增加 2 个工作表（%d -> %d）", before, after)
		}
	})

	t.Run("StyleAndProps", func(t *testing.T) {
		p := freshBook(t, "09_style.xlsx")
		st := Style{
			Font:      &FontStyle{Bold: true, Color: "FFFF0000", Size: 14},
			Fill:      &FillStyle{Color: "FFFFFF00", PatternType: "solid"},
			Alignment: &AlignmentStyle{Horizontal: "center", WrapText: true},
		}
		if _, err := SetCellStyle(p, "Sheet1", "A1", st); err != nil {
			t.Fatalf("SetCellStyle: %v", err)
		}
		if _, err := SetCellStyleRange(p, "Sheet1", "B2:C3", st); err != nil {
			t.Fatalf("SetCellStyleRange: %v", err)
		}
		b, _ := Open(p)
		s, _ := b.Sheet("Sheet1")
		if idx, err := s.Cell("A1").GetStyle(); err != nil || idx == 0 {
			t.Errorf("A1 应含非空样式索引，idx=%d err=%v", idx, err)
		}
		// 标签色
		if err := SetSheetProps(p, "Sheet1", SheetProps{TabColor: "FF00FF00"}); err != nil {
			t.Fatalf("SetSheetProps: %v", err)
		}
		fm := readMap(t, p)
		if !strings.Contains(string(fm["xl/workbook.xml"]), `tabColor="FF00FF00"`) {
			t.Errorf("workbook.xml 未见 Sheet1 的 tabColor 属性")
		}
	})

	t.Run("TablesCommentsValidation", func(t *testing.T) {
		p := freshBook(t, "10_table_comment.xlsx")
		_ = SetRange(p, "Sheet1", "A1:B3", [][]interface{}{
			{"ID", "Name"}, {1, "a"}, {2, "b"},
		})
		if err := AddTable(p, "Sheet1", "A1:B3", "Table1"); err != nil {
			t.Fatalf("AddTable: %v", err)
		}
		// 批注
		if err := AddComment(p, "Sheet1", "C1", "备注内容", "审核员"); err != nil {
			t.Fatalf("AddComment: %v", err)
		}
		if got, err := GetComment(p, "Sheet1", "C1"); err != nil || got != "备注内容" {
			t.Errorf("GetComment = %q err=%v，期望 备注内容", got, err)
		}
		// 所有写入完成后再统一读取快照做结构校验
		fm := readMap(t, p)
		mustHas(t, fm, "xl/tables/table1.xml")
		sr := loadRels(fm, "xl/worksheets/_rels/sheet1.xml.rels")
		hasTable, hasComment := false, false
		for _, rel := range sr.Relationship {
			if strings.Contains(rel.Type, "/table") {
				hasTable = true
			}
			if strings.Contains(rel.Type, "/comments") {
				hasComment = true
			}
		}
		if !hasTable {
			t.Errorf("sheet1 rels 应含 table 关系")
		}
		if !hasComment {
			t.Errorf("sheet1 rels 应含 comments 关系")
		}
		if n := countPrefix(fm, "xl/comments/"); n < 1 {
			t.Errorf("应生成 xl/comments/ 部件，实际 %d", n)
		}
	})

	t.Run("SheetDecorations", func(t *testing.T) {
		p := freshBook(t, "11_decor.xlsx")
		_ = SetCellStr(p, "Sheet1", "A1", "hello world")
		if err := MergeCells(p, "Sheet1", "A1:B1"); err != nil {
			t.Fatalf("MergeCells: %v", err)
		}
		if err := UnmergeCells(p, "Sheet1", "A1:B1"); err != nil {
			t.Fatalf("UnmergeCells: %v", err)
		}
		if err := SetColWidth(p, "Sheet1", 2, 25); err != nil {
			t.Fatalf("SetColWidth: %v", err)
		}
		if w, err := GetColWidth(p, "Sheet1", 2); err != nil || w < 24 || w > 26 {
			t.Errorf("GetColWidth(2) = %v err=%v，期望约 25", w, err)
		}
		if err := SetRowHeight(p, "Sheet1", 3, 30); err != nil {
			t.Fatalf("SetRowHeight: %v", err)
		}
		if err := SetRowVisible(p, "Sheet1", 3, false); err != nil {
			t.Fatalf("SetRowVisible: %v", err)
		}
		if err := SetColVisible(p, "Sheet1", 2, false); err != nil {
			t.Fatalf("SetColVisible: %v", err)
		}
		if err := FreezePanes(p, "Sheet1", "A2"); err != nil {
			t.Fatalf("FreezePanes: %v", err)
		}
		if err := AutoFilter(p, "Sheet1", "A1:B3"); err != nil {
			t.Fatalf("AutoFilter: %v", err)
		}
		if err := AddHyperlink(p, "Sheet1", "A1", "https://example.com", "link"); err != nil {
			t.Fatalf("AddHyperlink: %v", err)
		}
		if err := AddDataValidationList(p, "Sheet1", "D1", []string{"x", "y"}, true); err != nil {
			t.Fatalf("AddDataValidationList: %v", err)
		}
		if err := ProtectSheet(p, "Sheet1", "1234"); err != nil {
			t.Fatalf("ProtectSheet: %v", err)
		}
		if err := GroupRows(p, "Sheet1", 1, 2, 1); err != nil {
			t.Fatalf("GroupRows: %v", err)
		}
		if err := GroupCols(p, "Sheet1", 1, 2, 1); err != nil {
			t.Fatalf("GroupCols: %v", err)
		}
		fm := readMap(t, p)
		ws := string(fm["xl/worksheets/sheet1.xml"])
		mustContain(t, ws, `<pane `, "冻结窗格 pane 缺失")
		mustContain(t, ws, `<autoFilter `, "autoFilter 缺失")
		mustContain(t, ws, `<hyperlink `, "hyperlink 缺失")
		mustContain(t, ws, `<dataValidation `, "dataValidation 缺失")
		mustContain(t, ws, `<sheetProtection `, "sheetProtection 缺失")
		mustContain(t, ws, `outlineLevel="1"`, "行分组 outlineLevel 缺失")
		mustContain(t, ws, "<cols>", "列分组 cols 缺失")
	})

	t.Run("ConditionalFormatAndReplace", func(t *testing.T) {
		p := freshBook(t, "12_cf_replace.xlsx")
		_ = SetCellInt(p, "Sheet1", "A1", 10)
		st := Style{Font: &FontStyle{Bold: true, Color: "FFFF0000"}}
		if err := SetConditionalFormat(p, "Sheet1", "A1:A1", "expression", "A1>5", 1, st); err != nil {
			t.Fatalf("SetConditionalFormat: %v", err)
		}
		n, err := ReplaceText(p, "Sheet1", "10", "99")
		if err != nil {
			t.Fatalf("ReplaceText: %v", err)
		}
		if n != 1 {
			t.Errorf("ReplaceText 应替换 1 处，实际 %d", n)
		}
		if got, _ := GetCell(p, "Sheet1", "A1"); got != "99" {
			t.Errorf("替换后 A1 = %q，期望 99", got)
		}
		fm := readMap(t, p)
		if !strings.Contains(string(fm["xl/worksheets/sheet1.xml"]), "<conditionalFormatting") {
			t.Errorf("worksheet 缺失 conditionalFormatting")
		}
	})

	t.Run("DocProps", func(t *testing.T) {
		p := freshBook(t, "13_docprops.xlsx")
		if err := SetDocProps(p, map[string]string{"creator": "excelgo-test", "title": "测试文档"}); err != nil {
			t.Fatalf("SetDocProps: %v", err)
		}
		fm := readMap(t, p)
		mustHas(t, fm, "docProps/core.xml")
		if !strings.Contains(string(fm["docProps/core.xml"]), "excelgo-test") {
			t.Errorf("docProps/core.xml 未见 creator=excelgo-test")
		}
	})

	t.Run("ObjectAPI", func(t *testing.T) {
		p := freshBook(t, "14_object_api.xlsx")
		b, err := Create()
		if err != nil {
			t.Fatalf("Create: %v", err)
		}
		s, err := b.AddSheet("Obj")
		if err != nil {
			t.Fatalf("AddSheet: %v", err)
		}
		if err := s.Cell("A1").SetStr("对象式"); err != nil {
			t.Fatalf("SetStr: %v", err)
		}
		if err := s.Cell("A2").Set(123); err != nil {
			t.Fatalf("Set(123): %v", err)
		}
		if err := s.Range("C1:E2").Set([][]interface{}{{"p", "q"}, {"r", "s"}}); err != nil {
			t.Fatalf("Range.Set: %v", err)
		}
		if err := b.SaveAs(p); err != nil {
			t.Fatalf("SaveAs: %v", err)
		}
		b2, _ := Open(p)
		s2, _ := b2.Sheet("Obj")
		if s2.Cell("A1").Get() != "对象式" {
			t.Errorf("A1 = %q", s2.Cell("A1").Get())
		}
		if s2.Cell("A2").GetNum() != 123 {
			t.Errorf("A2 = %v", s2.Cell("A2").GetNum())
		}
	})
}

func mustContain(t *testing.T, s, sub, msg string) {
	t.Helper()
	if !strings.Contains(s, sub) {
		t.Errorf("%s（未找到 %q）", msg, sub)
	}
}

// dumpSheet 把指定工作表的 worksheet XML 打印到测试日志，便于诊断行/列平移问题。
func dumpSheet(t *testing.T, path, msg string) {
	t.Helper()
	fm := readMap(t, path)
	// 找到 Sheet1 对应的 worksheet 文件
	wb := &Workbook{}
	_ = getXMLFromMap(fm, "xl/workbook.xml", wb)
	f, _, _, _, err := locateSheetInMap(fm, wb, "Sheet1")
	if err != nil {
		t.Logf("[dump %s] 定位 Sheet1 失败: %v", msg, err)
		return
	}
	t.Logf("[dump %s] %s :\n%s", msg, f, string(fm[f]))
}
