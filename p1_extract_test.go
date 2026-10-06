package excelgo

import (
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

func TestP1ExtractPicture(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "extract.xlsx")
	b, _ := Create()
	ws, _ := b.Sheet("Sheet1")
	ws.SetCellStr("A1", "x")
	b.SaveAs(p)

	src, err := os.ReadFile("testfixtures/ph_blue.png")
	if err != nil {
		t.Fatal(err)
	}
	if err := AddPictureFromBytes(p, "Sheet1", src, &PictureOptions{Position: &PicturePosition{Cell: "A1"}}); err != nil {
		t.Fatal(err)
	}

	pics, err := GetPictures(p, "Sheet1")
	if err != nil || len(pics) != 1 {
		t.Fatalf("GetPictures 不符: %v err=%v", pics, err)
	}
	if pics[0].FileInZip == "" || pics[0].Name == "" {
		t.Fatalf("图片信息不完整: %+v", pics[0])
	}
	// 导出字节
	data, err := ExtractPicture(p, "Sheet1", pics[0].FileInZip)
	if err != nil {
		t.Fatalf("ExtractPicture 失败: %v", err)
	}
	if len(data) != len(src) {
		t.Fatalf("导出字节长度不符: got %d want %d", len(data), len(src))
	}
	for i := range src {
		if data[i] != src[i] {
			t.Fatalf("导出字节内容不一致，位置 %d", i)
		}
	}

	// 独立交叉校验：openpyxl 严格加载文件，确认图片部件良构可用
	py := `C:\Users\null\.workbuddy\binaries\python\envs\validate\Scripts\python.exe`
	if _, statErr := os.Stat(py); statErr == nil {
		script := `import openpyxl, sys
f = sys.argv[1]
wb = openpyxl.load_workbook(f)  # 严格解析
ws = wb["Sheet1"]
print("OPENPYXL_OK", ws)
`
		out, e := exec.Command(py, "-c", script, p).CombinedOutput()
		if e != nil {
			t.Fatalf("openpyxl 交叉校验失败: %v\n%s", e, string(out))
		}
		if !strings.Contains(string(out), "OPENPYXL_OK") {
			t.Fatalf("openpyxl 交叉校验未通过: %s", string(out))
		}
	} else {
		t.Logf("跳过 openpyxl 交叉校验（未找到 %s）", py)
	}
}
