package excelgo

import (
	"path/filepath"
	"testing"
)

func TestP1ReadSymmetric(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "p1read.xlsx")

	b, err := Create()
	if err != nil {
		t.Fatal(err)
	}
	ws, err := b.Sheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	ws.SetCellStr("A1", "数据")

	// 写入各项特性
	if err := ws.MergeCells("A1:B1"); err != nil {
		t.Fatal(err)
	}
	if err := ws.AddComment("A1", "批注文本", "作者X"); err != nil {
		t.Fatal(err)
	}
	if err := ws.AddHyperlink("A1", "https://example.com", "示例"); err != nil {
		t.Fatal(err)
	}
	if err := ws.AddDataValidation("A2:A10", "list", "between", `"a,b,c"`, "", true); err != nil {
		t.Fatal(err)
	}
	if err := ws.SetConditionalFormat("B1:B10", "cellIs", ">5", 1, Style{}); err != nil {
		t.Fatal(err)
	}
	if err := b.SaveAs(p); err != nil {
		t.Fatal(err)
	}
	// 图片需落盘后由独立函数写入
	if err := AddPicture(p, "Sheet1", "testfixtures/ph_blue.png", &PictureOptions{}); err != nil {
		t.Fatal(err)
	}

	// 重新打开读取
	ob, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	ows, err := ob.Sheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}

	// 合并
	merges, _ := ows.GetMergeCells()
	if len(merges) != 1 || merges[0] != "A1:B1" {
		t.Fatalf("合并读回不符: %v", merges)
	}
	// 批注
	comments, _ := ows.GetComments()
	if len(comments) != 1 || comments[0].Cell != "A1" || comments[0].Text != "批注文本" || comments[0].Author != "作者X" {
		t.Fatalf("批注读回不符: %+v", comments)
	}
	// 超链接
	hls, _ := ows.GetHyperlinks()
	found := false
	for _, h := range hls {
		if h.Cell == "A1" && h.URL == "https://example.com" && h.External {
			found = true
		}
	}
	if !found {
		t.Fatalf("超链接读回不符: %+v", hls)
	}
	// 数据验证
	dvs, _ := ows.GetDataValidations()
	if len(dvs) != 1 || dvs[0].Ref != "A2:A10" || dvs[0].Type != "list" || dvs[0].Formula1 != `"a,b,c"` || !dvs[0].AllowBlank {
		t.Fatalf("数据验证读回不符: %+v", dvs)
	}
	// 条件格式
	cfs, _ := ows.GetConditionalFormats()
	if len(cfs) != 1 || cfs[0].Ref != "B1:B10" || cfs[0].Type != "cellIs" || cfs[0].Formula != ">5" || cfs[0].Priority != 1 {
		t.Fatalf("条件格式读回不符: %+v", cfs)
	}
	// 图片
	pics, _ := ows.GetPictures()
	if len(pics) != 1 {
		t.Fatalf("图片读回数量不符: %+v", pics)
	}
	if pics[0].CellRef != "A1" || pics[0].Name == "" || pics[0].FileInZip == "" {
		t.Fatalf("图片信息不符: %+v", pics[0])
	}
}

func TestP1ReadGlobals(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "p1g.xlsx")
	b, _ := Create()
	ws, _ := b.Sheet("Sheet1")
	ws.SetCellStr("C3", "x")
	ws.MergeCells("C3:D3")
	ws.AddComment("C3", "全局读回测试", "Tester")
	b.SaveAs(p)

	if merges, err := GetMergeCells(p, "Sheet1"); err != nil || len(merges) != 1 || merges[0] != "C3:D3" {
		t.Fatalf("GetMergeCells 不符: %v err=%v", merges, err)
	}
	cs, err := GetComments(p, "Sheet1")
	if err != nil || len(cs) != 1 || cs[0].Text != "全局读回测试" {
		t.Fatalf("GetComments 不符: %v err=%v", cs, err)
	}
}
