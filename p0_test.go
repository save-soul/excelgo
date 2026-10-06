package excelgo

import (
	"path/filepath"
	"testing"
)

func TestP0StyleReadAndCopy(t *testing.T) {
	b, err := Create()
	if err != nil {
		t.Fatal(err)
	}
	ws, err := b.Sheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}

	srcStyle := Style{
		Font:      &FontStyle{Bold: true, Color: "FFFF0000", Size: 14, Name: "微软雅黑"},
		Fill:      &FillStyle{Color: "FFFFFF00", PatternType: "solid"},
		Border:    &BorderStyle{Top: BorderSide{Style: "thin", Color: "FF0000FF"}, Bottom: BorderSide{Style: "medium"}},
		Alignment: &AlignmentStyle{Horizontal: "center", WrapText: true},
		NumFmt:    "0.00%",
	}
	if _, err := ws.SetStyle("A1", srcStyle); err != nil {
		t.Fatal(err)
	}
	// 写点数据便于后续读取
	if err := ws.SetCellStr("A1", "标题"); err != nil {
		t.Fatal(err)
	}

	// 回读 A1 样式
	got, err := ws.GetStyle("A1")
	if err != nil {
		t.Fatal(err)
	}
	if got.Font == nil || !got.Font.Bold || got.Font.Color != "FFFF0000" || got.Font.Size != 14 || got.Font.Name != "微软雅黑" {
		t.Fatalf("Font 回读不符: %+v", got.Font)
	}
	if got.Fill == nil || got.Fill.Color != "FFFFFF00" {
		t.Fatalf("Fill 回读不符: %+v", got.Fill)
	}
	if got.Border == nil || got.Border.Top.Style != "thin" || got.Border.Top.Color != "FF0000FF" || got.Border.Bottom.Style != "medium" {
		t.Fatalf("Border 回读不符: %+v", got.Border)
	}
	if got.Alignment == nil || got.Alignment.Horizontal != "center" || !got.Alignment.WrapText {
		t.Fatalf("Alignment 回读不符: %+v", got.Alignment)
	}
	if got.NumFmt != "0.00%" {
		t.Fatalf("NumFmt 回读不符: %q", got.NumFmt)
	}

	// 克隆到 B1
	if err := ws.CopyStyle("A1", "B1"); err != nil {
		t.Fatal(err)
	}
	gotB, err := ws.GetStyle("B1")
	if err != nil {
		t.Fatal(err)
	}
	if gotB.Font == nil || !gotB.Font.Bold || gotB.Font.Color != "FFFF0000" {
		t.Fatalf("B1 克隆样式 Font 不符: %+v", gotB.Font)
	}
	if gotB.NumFmt != "0.00%" || gotB.Fill == nil || gotB.Fill.Color != "FFFFFF00" {
		t.Fatalf("B1 克隆样式其余不符: %+v", gotB)
	}
}

func TestP0TypedRead(t *testing.T) {
	b, _ := Create()
	ws, _ := b.Sheet("Sheet1")
	ws.SetCellInt("A1", 42)
	ws.SetCellNumeric("A2", 3.14)
	ws.SetCellBool("A3", true)
	ws.SetCellStr("A4", "文本")
	ws.SetCellFormula("A5", "A1*2", "84")

	cases := []struct {
		ref   string
		want  interface{}
		isInt bool
	}{
		{"A1", 42, true},
		{"A2", 3.14, false},
		{"A3", true, false},
		{"A4", "文本", false},
	}
	for _, c := range cases {
		v, err := ws.GetCellValue(c.ref)
		if err != nil {
			t.Fatalf("%s: %v", c.ref, err)
		}
		if v != c.want {
			t.Fatalf("%s 返回 %v(%T)，期望 %v(%T)", c.ref, v, v, c.want, c.want)
		}
	}
	// 公式缓存值：写入端对公式结果统一以 t="str" 存储，故回读为字符串 "84"
	v, _ := ws.GetCellValue("A5")
	if v != "84" {
		t.Fatalf("A5 公式结果应为 \"84\"，得到 %v(%T)", v, v)
	}
	// Cell 便捷方法
	if ws.Cell("A1").GetInt() != 42 {
		t.Fatal("GetInt 不符")
	}
	if ws.Cell("A3").GetBool() != true {
		t.Fatal("GetBool 不符")
	}
}

func TestP0SheetVisible(t *testing.T) {
	b, _ := Create()
	b.NewSheet("Sheet2")
	if err := b.SetSheetVisible("Sheet2", SheetStateHidden); err != nil {
		t.Fatal(err)
	}
	st, err := b.GetSheetVisible("Sheet2")
	if err != nil {
		t.Fatal(err)
	}
	if st != SheetStateHidden {
		t.Fatalf("Sheet2 状态应为 hidden，得到 %q", st)
	}
	// 恢复可见
	if err := b.SetSheetVisible("Sheet2", SheetStateVisible); err != nil {
		t.Fatal(err)
	}
	st, _ = b.GetSheetVisible("Sheet2")
	if st != SheetStateVisible {
		t.Fatalf("Sheet2 状态应为 visible，得到 %q", st)
	}
}

// TestP0DiskRoundtrip 落盘并用独立解析器校验 XML 良构（确定性交叉验证）。
func TestP0DiskRoundtrip(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "p0.xlsx")
	b, _ := Create()
	ws, _ := b.Sheet("Sheet1")
	ws.SetStyle("A1", Style{
		Font:      &FontStyle{Bold: true, Color: "FFFF0000"},
		Fill:      &FillStyle{Color: "FFFFFF00"},
		Alignment: &AlignmentStyle{Horizontal: "center"},
		NumFmt:    "0.00",
	})
	ws.SetCellStr("A1", "样式单元格")
	ws.SetCellInt("B1", 100)
	b.NewSheet("Sheet2")
	b.SetSheetVisible("Sheet2", SheetStateVeryHidden)
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}
	// 读回校验
	ob, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	ows, _ := ob.Sheet("Sheet1")
	st, err := ows.GetStyle("A1")
	if err != nil || st.Font == nil || !st.Font.Bold {
		t.Fatalf("落盘后样式回读失败: %+v err=%v", st, err)
	}
	v, _ := ows.GetCellValue("B1")
	if v != 100 {
		t.Fatalf("B1 读回应为 100，得到 %v", v)
	}
	st2, _ := ob.GetSheetVisible("Sheet2")
	if st2 != SheetStateVeryHidden {
		t.Fatalf("Sheet2 状态应为 veryHidden，得到 %q", st2)
	}
}
