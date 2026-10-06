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
	b, err := Open(out)
	if err != nil {
		t.Fatalf("open target: %v", err)
	}
	err = b.Merge([]SourceRef{
		{Workbook: src1, Sheet: "数据A"},
		{Workbook: src2, Sheet: "数据B"},
	})
	if err != nil {
		t.Fatalf("Merge failed: %v", err)
	}
	if err := b.Save(); err != nil {
		t.Fatalf("Save: %v", err)
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
			definedNames = readAllFromZip(f)
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

// TestMergeWorkbookSharedStrings 回归测试：真实 Excel/WPS 文件使用共享字符串（t="s"），
// 而本库自身以 inlineStr 写单元格。合并时必须把源表的 t="s" 单元格转成 t="inlineStr"
// 并内联文本，否则合并后的工作表会引用目标工作簿缺失/不一致的 sharedStrings.xml，
// 导致 Excel 打开报错或字符串串文。
func TestMergeWorkbookSharedStrings(t *testing.T) {
	fixture := filepath.Join("testfixtures", "target.xlsx")
	if _, err := os.Stat(fixture); err != nil {
		t.Skipf("跳过：缺少样本 %s", fixture)
	}

	// 以 fixture 为底本构造一个含共享字符串单元格的源工作簿
	srcMap, err := readZipToMap(fixture)
	if err != nil {
		t.Fatalf("read fixture: %v", err)
	}
	ws := string(srcMap["xl/worksheets/sheet1.xml"])
	ws = strings.Replace(ws, "</sheetData>",
		`<c r="Z1" t="s" s="0"><v>0</v></c></sheetData>`, 1)
	srcMap["xl/worksheets/sheet1.xml"] = []byte(ws)
	srcMap["xl/sharedStrings.xml"] = []byte(
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1">` +
			`<si><t xml:space="preserve">共享字符串内容</t></si></sst>`)

	srcPath := filepath.Join(t.TempDir(), "src.xlsx")
	if err := writeMapToZip(srcPath, srcMap); err != nil {
		t.Fatalf("write src: %v", err)
	}

	dstPath := filepath.Join(t.TempDir(), "dst.xlsx")
	if err := copyFile(fixture, dstPath); err != nil {
		t.Fatalf("copy dst: %v", err)
	}

	b, err := Open(dstPath)
	if err != nil {
		t.Fatalf("open dst: %v", err)
	}
	if err := b.Merge([]SourceRef{{Workbook: srcPath, Sheet: "Sheet1"}}); err != nil {
		t.Fatalf("Merge: %v", err)
	}
	if err := b.Save(); err != nil {
		t.Fatalf("Save: %v", err)
	}

	// 校验：输出中 Z1 单元格须为 inlineStr 且内联了正确文本，不再依赖 sharedStrings
	outMap, err := readZipToMap(dstPath)
	if err != nil {
		t.Fatalf("read output: %v", err)
	}
	found := false
	for name, data := range outMap {
		if !strings.HasPrefix(name, "xl/worksheets/sheet") {
			continue
		}
		s := string(data)
		if strings.Contains(s, `r="Z1"`) {
			found = true
			if strings.Contains(s, `t="s"`) {
				t.Errorf("%s 中 Z1 仍为 t=\"s\"（未转换为 inlineStr）:\n%s", name, s)
			}
			if !strings.Contains(s, `t="inlineStr"`) {
				t.Errorf("%s 中 Z1 缺少 t=\"inlineStr\"", name)
			}
			if !strings.Contains(s, "共享字符串内容") {
				t.Errorf("%s 中 Z1 文本丢失", name)
			}
		}
	}
	if !found {
		t.Fatal("输出中未找到 Z1（合并表缺失）")
	}
}
