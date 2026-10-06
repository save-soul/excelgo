package excelgo

import (
	"archive/zip"
	"os"
	"path/filepath"
	"testing"
)

// TestCopySheetSharedVsIndependent 验证两种媒体策略都能产出可打开的文件，
// 且独立策略会复制媒体、共享策略不复制。需要仓库根目录存在测试样本。
func TestCopySheetSharedVsIndependent(t *testing.T) {
	// 在仓库内向上查找测试样本
	candidates := []string{
		filepath.Join("..", "新建 XLSX 工作表.xlsx"),
		"新建 XLSX 工作表.xlsx",
	}
	sample := ""
	for _, c := range candidates {
		if _, err := os.Stat(c); err == nil {
			sample = c
			break
		}
	}
	if sample == "" {
		t.Skip("跳过：未找到测试样本 新建 XLSX 工作表.xlsx")
	}

	sharedOut := filepath.Join(t.TempDir(), "shared.xlsx")
	bf, err := Open(sample)
	if err != nil {
		t.Fatalf("open sample: %v", err)
	}
	if err := bf.CopySheetTo(sharedOut, "Sheet1", "", WithMedia(MediaShared)); err != nil {
		t.Fatalf("shared copy failed: %v", err)
	}
	indepOut := filepath.Join(t.TempDir(), "indep.xlsx")
	bf2, err := Open(sample)
	if err != nil {
		t.Fatalf("open sample: %v", err)
	}
	if err := bf2.CopySheetTo(indepOut, "Sheet1", "", WithMedia(MediaIndependent)); err != nil {
		t.Fatalf("independent copy failed: %v", err)
	}

	// 直接以 zip 结构校验（不依赖外部库）
	for _, f := range []string{sharedOut, indepOut} {
		if _, err := os.Stat(f); err != nil {
			t.Fatalf("output not created: %v", err)
		}
	}

	// 共享应只含 1 个媒体；独立应含 2 个媒体。
	if cnt := countMedia(sharedOut); cnt != 1 {
		t.Errorf("shared media count = %d, want 1", cnt)
	}
	if cnt := countMedia(indepOut); cnt != 2 {
		t.Errorf("independent media count = %d, want 2", cnt)
	}
}

func countMedia(path string) int {
	r, err := zip.OpenReader(path)
	if err != nil {
		return -1
	}
	defer r.Close()
	cnt := 0
	for _, f := range r.File {
		if len(f.Name) >= 9 && f.Name[:9] == "xl/media/" {
			cnt++
		}
	}
	return cnt
}
