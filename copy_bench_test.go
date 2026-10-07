package excelgo

import (
	"fmt"
	"strings"
	"testing"
)

// TestCopySheetNotGrouped 验证复制体不再继承源表的 tabSelected="1"
// （两张表同时 tabSelected=1 会被 Excel/WPS 判定为「成组工作表」）。
func TestCopySheetNotGrouped(t *testing.T) {
	p := genPath("bench_copy.xlsx")
	b0, _ := Create()
	if err := b0.SaveAs(p); err != nil {
		t.Fatal(err)
	}

	// 把 Sheet1 设为活动表（tabSelected="1"），复现真实 Excel/WPS 保存态
	b1, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	ws := string(b1.fileMap["xl/worksheets/sheet1.xml"])
	ws = strings.Replace(ws, "<sheetData>",
		`<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData>`, 1)
	b1.fileMap["xl/worksheets/sheet1.xml"] = []byte(ws)
	if err := b1.Save(); err != nil {
		t.Fatal(err)
	}

	b2, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := b2.CopySheet("Sheet1", "Sheet1_copy"); err != nil {
		t.Fatal(err)
	}
	if err := b2.Save(); err != nil {
		t.Fatal(err)
	}

	fm := readMap(t, p)
	src := string(fm["xl/worksheets/sheet1.xml"])
	dst := string(fm["xl/worksheets/sheet2.xml"])
	if !strings.Contains(src, `tabSelected="1"`) {
		t.Fatal("前置条件失败：源表应有 tabSelected=\"1\"")
	}
	if strings.Contains(dst, "tabSelected") {
		t.Errorf("副本不应保留 tabSelected，实际: %s", sheetViewsSeg(dst))
	}
	t.Logf("源表   : %s", sheetViewsSeg(src))
	t.Logf("副本表 : %s", sheetViewsSeg(dst))
}

func sheetViewsSeg(ws string) string {
	i := strings.Index(ws, "<sheetViews")
	j := strings.Index(ws, "</sheetViews>")
	if i < 0 || j <= i {
		return "(no sheetViews)"
	}
	return ws[i : j+13]
}

func BenchmarkCopySheet(b *testing.B) {
	src := genPath("bench_copy_src.xlsx")
	b0, _ := Create()
	if err := b0.SaveAs(src); err != nil {
		b.Fatal(err)
	}
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		bb, err := Open(src)
		if err != nil {
			b.Fatal(err)
		}
		if _, err := bb.CopySheet("Sheet1", "Sheet1_copy"); err != nil {
			b.Fatal(err)
		}
	}
}

func BenchmarkCopySheetLarge(b *testing.B) {
	src := genPath("bench_copy_large.xlsx")

	// 用内存态 API 造数据：SetCellValue 挂在 *WorkSheet 上，纯内存累积不落盘
	b0, _ := Create()
	sh, err := b0.Sheet("Sheet1")
	if err != nil {
		b.Fatal(err)
	}
	pad := strings.Repeat("x", 20)
	for r := 1; r <= 5000; r++ {
		_ = sh.SetCellValue(fmt.Sprintf("A%d", r), pad)
		_ = sh.SetCellValue(fmt.Sprintf("B%d", r), "value")
	}
	// 标为活动表，复现真实 Excel/WPS 保存态
	ws := string(b0.fileMap["xl/worksheets/sheet1.xml"])
	ws = strings.Replace(ws, "<sheetData>",
		`<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews><sheetData>`, 1)
	b0.fileMap["xl/worksheets/sheet1.xml"] = []byte(ws)
	if err := b0.SaveAs(src); err != nil {
		b.Fatal(err)
	}

	b.ResetTimer()
	b.ReportAllocs()
	for i := 0; i < b.N; i++ {
		b2, err := Open(src)
		if err != nil {
			b.Fatal(err)
		}
		if _, err := b2.CopySheet("Sheet1", "Sheet1_copy"); err != nil {
			b.Fatal(err)
		}
	}
}
