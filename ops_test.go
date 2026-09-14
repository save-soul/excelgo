package excelgo

import (
	"os"
	"path/filepath"
	"testing"
)

const opsTestSrc = "../testdata_tmp/full.xlsx"

// copyTestFile 把源复制到临时文件，避免污染样本。
func copyTestFile(t *testing.T, src, dst string) {
	t.Helper()
	data, err := os.ReadFile(src)
	if err != nil {
		t.Fatalf("读源文件失败: %v", err)
	}
	if err := os.WriteFile(dst, data, 0644); err != nil {
		t.Fatalf("写临时文件失败: %v", err)
	}
}

func TestSheetLifecycle(t *testing.T) {
	tmp := filepath.Join(t.TempDir(), "lifecycle.xlsx")
	copyTestFile(t, opsTestSrc, tmp)

	// 新建
	idx, err := NewSheet(tmp, "新建表")
	if err != nil {
		t.Fatalf("NewSheet: %v", err)
	}
	if idx != 4 {
		t.Errorf("NewSheet 返回序号应为 4，实际 %d", idx)
	}
	// 改名（避开同名自动加后缀）
	if err := RenameSheet(tmp, "新建表", "改名表"); err != nil {
		t.Fatalf("RenameSheet: %v", err)
	}
	// 移动：把 Sheet2 移到最前
	if err := MoveSheet(tmp, "Sheet2", 1); err != nil {
		t.Fatalf("MoveSheet: %v", err)
	}
	// 删除 Sheet3（带图片，应清理其专属 media）
	if err := DeleteSheet(tmp, "Sheet3"); err != nil {
		t.Fatalf("DeleteSheet: %v", err)
	}
	// 校验可再次打开且结构合法（交给 openpyxl；此处仅确认无panic）
	if _, err := os.Stat(tmp); err != nil {
		t.Fatalf("产物不存在: %v", err)
	}
}

func TestCellReadWrite(t *testing.T) {
	tmp := filepath.Join(t.TempDir(), "cell.xlsx")
	copyTestFile(t, opsTestSrc, tmp)

	// 写多种格式
	if err := SetCellStr(tmp, "Sheet1", "F1", "hello"); err != nil {
		t.Fatal(err)
	}
	if err := SetCellInt(tmp, "Sheet1", "F2", 42); err != nil {
		t.Fatal(err)
	}
	if err := SetCellNumeric(tmp, "Sheet1", "F3", 3.14); err != nil {
		t.Fatal(err)
	}
	if err := SetCellBool(tmp, "Sheet1", "F4", true); err != nil {
		t.Fatal(err)
	}
	if err := SetCellFormula(tmp, "Sheet1", "F5", "F2*2", "84"); err != nil {
		t.Fatal(err)
	}
	if err := SetCellValue(tmp, "Sheet1", "F6", "auto-string"); err != nil {
		t.Fatal(err)
	}

	// 读回
	checks := map[string]string{"F1": "hello", "F2": "42", "F3": "3.14", "F4": "1", "F6": "auto-string"}
	for c, want := range checks {
		got, err := GetCell(tmp, "Sheet1", c)
		if err != nil {
			t.Fatalf("GetCell %s: %v", c, err)
		}
		if got != want {
			t.Errorf("单元格 %s 读写不一致: 期望 %q，实际 %q", c, want, got)
		}
	}
	// 公式单元格读回结果
	if got, _ := GetCell(tmp, "Sheet1", "F5"); got != "84" {
		t.Errorf("公式结果应为 84，实际 %q", got)
	}
	// 原有数据不应被破坏
	if got, _ := GetCell(tmp, "Sheet1", "A1"); got != "标题" {
		t.Errorf("原有 A1 被破坏: %q", got)
	}
}

func TestAddPictures(t *testing.T) {
	tmp := filepath.Join(t.TempDir(), "pic.xlsx")
	copyTestFile(t, opsTestSrc, tmp)
	picRed := "../testdata_tmp/ph_red.png"
	picBlue := "../testdata_tmp/ph_blue.png"

	// 浮动图片（带位置与缩放）
	if err := AddPicture(tmp, "Sheet1", picRed, &PictureOptions{
		ScaleX: 0.5, ScaleY: 0.5,
		Position: &PicturePosition{Cell: "H1", ColOffset: 10, RowOffset: 5},
	}); err != nil {
		t.Fatalf("AddPicture 浮动: %v", err)
	}
	// WPS 单元格内嵌图片
	if err := AddCellPicture(tmp, "Sheet2", "B2", picBlue); err != nil {
		t.Fatalf("AddCellPicture: %v", err)
	}
	// 再次在 Sheet1 追加一张浮动图（验证复用已有 drawing）
	if err := AddPicture(tmp, "Sheet1", picBlue, nil); err != nil {
		t.Fatalf("AddPicture 复用 drawing: %v", err)
	}
}

// TestWriteVerifyArtifacts 把各功能产物写到 testdata_tmp 固定路径，供 openpyxl 离线校验。
// 该测试专注于生成产物，不在此断言；断言在外部 Python 校验脚本进行。
func TestWriteVerifyArtifacts(t *testing.T) {
	base := "../testdata_tmp"
	gen := func(name string) string {
		p := filepath.Join(base, name)
		copyTestFile(t, opsTestSrc, p)
		return p
	}

	// 生命周期
	lc := gen("verify_lifecycle.xlsx")
	NewSheet(lc, "新建表")
	RenameSheet(lc, "新建表", "改名表")
	MoveSheet(lc, "Sheet2", 1)
	DeleteSheet(lc, "Sheet3")

	// 单元格
	cl := gen("verify_cell.xlsx")
	SetCellStr(cl, "Sheet1", "F1", "hello")
	SetCellInt(cl, "Sheet1", "F2", 42)
	SetCellNumeric(cl, "Sheet1", "F3", 3.14)
	SetCellBool(cl, "Sheet1", "F4", true)
	SetCellFormula(cl, "Sheet1", "F5", "F2*2", "84")
	SetCellValue(cl, "Sheet1", "F6", "auto-string")

	// 图片
	pic := gen("verify_pic.xlsx")
	AddPicture(pic, "Sheet1", "../testdata_tmp/ph_red.png", &PictureOptions{
		ScaleX: 0.5, ScaleY: 0.5,
		Position: &PicturePosition{Cell: "H1", ColOffset: 10, RowOffset: 5},
	})
	AddCellPicture(pic, "Sheet2", "B2", "../testdata_tmp/ph_blue.png")
	AddPicture(pic, "Sheet1", "../testdata_tmp/ph_blue.png", nil)
}
