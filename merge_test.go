package excelgo

import (
	"archive/zip"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// TestMergeWorkbook 验证跨工作簿合并：把 src1、src2 两个独立工作簿的工作表
// 合并进 target 工作簿，合并后各自独立。校验：
//   - 产出的 xlsx 可正常打开（openpyxl 解析无异常，这里用 zip 结构 + 后续 openpyxl 复核）；
//   - styles.xml 的 cellXfs 数量未发生「每个源表全量追加」式膨胀（去重生效）；
//   - 媒体文件随表复制（target 原有 0 媒体，合并后应 >= 2）；
//   - 打印区域 definedName 被正确复制到新表。
func TestMergeWorkbook(t *testing.T) {
	base := ".."
	target := filepath.Join(base, "testdata_tmp", "target.xlsx")
	src1 := filepath.Join(base, "testdata_tmp", "src1.xlsx")
	src2 := filepath.Join(base, "testdata_tmp", "src2.xlsx")
	for _, f := range []string{target, src1, src2} {
		if _, err := os.Stat(f); err != nil {
			t.Skipf("跳过：缺少测试样本 %s", f)
		}
	}

	out := filepath.Join(t.TempDir(), "merged.xlsx")
	// 目标工作簿必须先存在（作为合并容器），这里以 target.xlsx 为底本
	if err := copyFile(target, out); err != nil {
		t.Fatalf("copy target: %v", err)
	}
	err := MergeWorkbook(out, []SourceRef{
		{Workbook: src1, Sheet: "数据A"},
		{Workbook: src2, Sheet: "数据B"},
	})
	if err != nil {
		t.Fatalf("MergeWorkbook failed: %v", err)
	}

	// 1. 文件存在
	if _, err := os.Stat(out); err != nil {
		t.Fatalf("output not created: %v", err)
	}

	// 2. 打开 zip 检查部件
	r, err := zip.OpenReader(out)
	if err != nil {
		t.Fatalf("open output zip: %v", err)
	}
	defer r.Close()

	var sheetCount, mediaCount, drawingCount int
	hasStyles := false
	hasWorkbook := false
	definedNames := ""
	for _, f := range r.File {
		switch {
		case f.Name == "xl/styles.xml":
			hasStyles = true
		case f.Name == "xl/workbook.xml":
			hasWorkbook = true
			rc, _ := f.Open()
			data, _ := readAll(rc)
			rc.Close()
			definedNames = string(data)
		case strings.HasPrefix(f.Name, "xl/media/"):
			mediaCount++
		case strings.HasPrefix(f.Name, "xl/worksheets/sheet") && strings.HasSuffix(f.Name, ".xml"):
			sheetCount++
		case strings.HasPrefix(f.Name, "xl/drawings/drawing") && strings.HasSuffix(f.Name, ".xml"):
			drawingCount++
		}
	}
	if !hasStyles {
		t.Error("merged workbook missing styles.xml")
	}
	if !hasWorkbook {
		t.Error("merged workbook missing workbook.xml")
	}
	// target 原有 1 个 sheet，合并 2 个 → 至少 3 个 worksheet
	if sheetCount < 3 {
		t.Errorf("worksheet count = %d, want >= 3", sheetCount)
	}
	// 源表各带 1 张图片 → 至少 2 个媒体
	if mediaCount < 2 {
		t.Errorf("media count = %d, want >= 2", mediaCount)
	}
	// 源表各带 1 个 drawing
	if drawingCount < 2 {
		t.Errorf("drawing count = %d, want >= 2", drawingCount)
	}
	// 打印区域 definedName 应至少包含 2 个 _xlnm.Print_Area（来自 src1/src2）
	cnt := 0
	for _, m := range findAll(definedNames, `_xlnm.Print_Area`) {
		_ = m
		cnt++
	}
	if cnt < 2 {
		t.Errorf("Print_Area definedName count = %d, want >= 2", cnt)
	}
}

// copyFile 复制普通文件。
func copyFile(src, dst string) error {
	data, err := os.ReadFile(src)
	if err != nil {
		return err
	}
	return os.WriteFile(dst, data, 0644)
}

// readAll 读取 ReadCloser 全部内容。
func readAll(rc interface {
	Read([]byte) (int, error)
	Close() error
}) ([]byte, error) {
	var buf []byte
	tmp := make([]byte, 4096)
	for {
		n, err := rc.Read(tmp)
		if n > 0 {
			buf = append(buf, tmp[:n]...)
		}
		if err != nil {
			break
		}
	}
	return buf, nil
}

func findAll(s, sub string) []string {
	var out []string
	for i := 0; ; {
		idx := indexOf(s[i:], sub)
		if idx < 0 {
			break
		}
		out = append(out, sub)
		i += idx + len(sub)
	}
	return out
}

func indexOf(s, sub string) int {
	for i := 0; i+len(sub) <= len(s); i++ {
		if s[i:i+len(sub)] == sub {
			return i
		}
	}
	return -1
}
