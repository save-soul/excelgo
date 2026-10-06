package excelgo

import (
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

func TestP1PrintProperties(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "print.xlsx")
	b, _ := Create()
	ws, _ := b.Sheet("Sheet1")
	ws.SetCellStr("A1", "评定表标题")
	b.SaveAs(p)

	gridTrue := true
	err := SetSheetProps(p, "Sheet1", SheetProps{
		TabColor:      "FFFF0000",
		ShowGridLines: &gridTrue,
		Zoom:          120,
		PageMargins: &PageMargins{
			Left: 0.7, Right: 0.7, Top: 0.75, Bottom: 0.75, Header: 0.3, Footer: 0.3,
		},
		PageSetup: &PageSetup{
			PaperSize: 9, Orientation: "landscape", Scale: 100, FitToWidth: 1, FitToHeight: 0,
		},
		PrintArea:      "A1:D20",
		PrintTitleRows: "1:1",
		HeaderFooter: &HeaderFooter{
			OddHeader: "&C评定表", OddFooter: "&L第&P页",
		},
	})
	if err != nil {
		t.Fatalf("SetSheetProps 失败: %v", err)
	}

	props, err := GetProps(p, "Sheet1")
	if err != nil {
		t.Fatalf("GetProps 失败: %v", err)
	}
	if props.TabColor != "FFFF0000" {
		t.Errorf("TabColor 不符: %q", props.TabColor)
	}
	if props.ShowGridLines == nil || !*props.ShowGridLines {
		t.Errorf("ShowGridLines 不符: %v", props.ShowGridLines)
	}
	if props.Zoom != 120 {
		t.Errorf("Zoom 不符: %d", props.Zoom)
	}
	if props.PageMargins == nil || props.PageMargins.Left != 0.7 || props.PageMargins.Bottom != 0.75 {
		t.Errorf("PageMargins 不符: %+v", props.PageMargins)
	}
	if props.PageSetup == nil || props.PageSetup.PaperSize != 9 || props.PageSetup.Orientation != "landscape" || props.PageSetup.FitToWidth != 1 {
		t.Errorf("PageSetup 不符: %+v", props.PageSetup)
	}
	if props.PrintArea != "$A$1:$D$20" {
		t.Errorf("PrintArea 不符: %q", props.PrintArea)
	}
	if props.PrintTitleRows != "$1:$1" {
		t.Errorf("PrintTitleRows 不符: %q", props.PrintTitleRows)
	}
	if props.HeaderFooter == nil || props.HeaderFooter.OddHeader != "&C评定表" || props.HeaderFooter.OddFooter != "&L第&P页" {
		t.Errorf("HeaderFooter 不符: %+v", props.HeaderFooter)
	}

	// 独立交叉校验：openpyxl 严格解析（任何非法 XML 会直接报错），并核对打印属性
	crossCheckPrintWithOpenpyxl(t, p)
}

// crossCheckPrintWithOpenpyxl 用 openpyxl 严格加载文件，验证良构性并核对关键打印属性。
func crossCheckPrintWithOpenpyxl(t *testing.T, p string) {
	py := `C:\Users\null\.workbuddy\binaries\python\envs\validate\Scripts\python.exe`
	if _, err := os.Stat(py); err != nil {
		t.Logf("跳过 openpyxl 交叉校验（未找到 %s）", py)
		return
	}
	script := `import openpyxl, sys
f = sys.argv[1]
wb = openpyxl.load_workbook(f)  # 严格解析，非法 XML 直接抛错
ws = wb["Sheet1"]
assert ws.page_margins is not None, "page_margins 缺失"
assert abs(ws.page_margins.left - 0.7) < 1e-6, ws.page_margins.left
assert ws.page_setup.paperSize == 9, ws.page_setup.paperSize
assert ws.page_setup.orientation == "landscape", ws.page_setup.orientation
assert "A1:D20" in ws.print_area.replace("$", ""), ws.print_area
print("OPENPYXL_OK")
`
	out, err := exec.Command(py, "-c", script, p).CombinedOutput()
	if err != nil {
		t.Fatalf("openpyxl 交叉校验失败: %v\n%s", err, string(out))
	}
	if !strings.Contains(string(out), "OPENPYXL_OK") {
		t.Fatalf("openpyxl 交叉校验未通过: %s", string(out))
	}
}
