package excelgo

import (
	"fmt"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

// 定位 openpyxl 校验 venv（独立于业务环境，专用于严格解析交叉验证）。
func openpyxlValidatePy() string {
	candidates := []string{
		`C:\Users\null\.workbuddy\binaries\python\envs\validate\Scripts\python.exe`,
		`C:\Users\null\.workbuddy\binaries\python\envs\default\Scripts\python.exe`,
	}
	for _, p := range candidates {
		if fileExists(p) {
			return p
		}
	}
	return ""
}

// 真正做存在性检查（复用 exec，避免跨平台 os.Stat 路径差异）。
func fileExists(path string) bool {
	// #nosec G304 - 仅检查文件存在性
	cmd := exec.Command("test", "-f", path)
	return cmd.Run() == nil
}

// crossCheckWorkbook 用 openpyxl 严格解析检查生成的 xlsx 合法，
// 并断言工作簿结构保护锁与命名区域（refersTo）符合预期。
// lockStructure 期望锁状态；definedName 为 map[name]refersTo，空表示不校验。
func crossCheckWorkbook(t *testing.T, filename string, lockStructure bool, definedName map[string]string) {
	t.Helper()
	py := openpyxlValidatePy()
	if py == "" {
		t.Skip("未找到 openpyxl 校验 venv，跳过交叉验证")
	}
	if !fileExists(py) {
		t.Skip("openpyxl 校验解释器不存在，跳过交叉验证")
	}

	var namesPart strings.Builder
	for name, ref := range definedName {
		namesPart.WriteString(fmt.Sprintf("try:\n    val = wb.defined_names[%q].value\n    print('DN %s=' + str(val))\n    assert %q in val, 'defined name %s 引用不符: ' + str(val)\nexcept Exception as e:\n    raise AssertionError('读取命名区域 %s 失败: ' + str(e))\n", name, name, ref, name, name))
	}

	script := `
import openpyxl, sys
fn = r'''` + filename + `'''
wb = openpyxl.load_workbook(fn, read_only=False)
# 结构保护锁（未保护时 wb.security 为 None）
sec = wb.security
cur = bool(sec.lock_structure) if sec is not None else False
assert cur == ` + b2s(lockStructure) + `, 'lock_structure 不符: ' + str(cur)
` + namesPart.String() + `
print('OK')
`
	out, err := exec.Command(py, "-c", script).CombinedOutput()
	if err != nil {
		t.Fatalf("openpyxl 交叉验证失败: %v\n输出:\n%s", err, string(out))
	}
	if !strings.Contains(string(out), "OK") {
		t.Fatalf("openpyxl 交叉验证未输出 OK:\n%s", string(out))
	}
}

func b2s(b bool) string {
	if b {
		return "True"
	}
	return "False"
}

func TestP1WorkbookProtection(t *testing.T) {
	dir := t.TempDir()
	fn := filepath.Join(dir, "wb_protect.xlsx")

	b, err := Create()
	if err != nil {
		t.Fatal(err)
	}
	if err := b.SaveAs(fn); err != nil {
		t.Fatal(err)
	}

	// 初始未保护
	if wp, _ := b.GetWorkbookProtection(); wp.LockStructure {
		t.Fatalf("初始应为未保护，却得到 lockStructure=true")
	}

	// 启用结构保护（无密码）
	if err := b.ProtectWorkbook(""); err != nil {
		t.Fatal(err)
	}
	wp, err := b.GetWorkbookProtection()
	if err != nil {
		t.Fatal(err)
	}
	if !wp.LockStructure {
		t.Fatalf("ProtectWorkbook 后 lockStructure 应为 true")
	}
	if wp.PasswordHash != "" {
		t.Fatalf("无密码时 PasswordHash 应为空，却得到 %q", wp.PasswordHash)
	}
	if err := b.Save(); err != nil {
		t.Fatal(err)
	}
	crossCheckWorkbook(t, fn, true, nil)

	// 带密码保护
	if err := b.ProtectWorkbook("secret"); err != nil {
		t.Fatal(err)
	}
	wp, _ = b.GetWorkbookProtection()
	if wp.PasswordHash == "" {
		t.Fatalf("带密码时 PasswordHash 不应为空")
	}
	if !isValidHex4(wp.PasswordHash) {
		t.Fatalf("PasswordHash 应为 4 位十六进制，却得到 %q", wp.PasswordHash)
	}
	// 相同密码哈希应稳定
	h1 := hashExcelPassword("secret")
	h2 := hashExcelPassword("secret")
	if h1 != h2 {
		t.Fatalf("相同密码哈希应稳定: %q vs %q", h1, h2)
	}
	// 不同密码哈希应不同
	if hashExcelPassword("secret") == hashExcelPassword("other") {
		t.Fatalf("不同密码哈希应不同")
	}

	// 解除保护
	if err := b.UnprotectWorkbook(); err != nil {
		t.Fatal(err)
	}
	wp, _ = b.GetWorkbookProtection()
	if wp.LockStructure {
		t.Fatalf("UnprotectWorkbook 后 lockStructure 应为 false")
	}
	if err := b.Save(); err != nil {
		t.Fatal(err)
	}
	crossCheckWorkbook(t, fn, false, nil)
}

func TestP1DefinedName(t *testing.T) {
	dir := t.TempDir()
	fn := filepath.Join(dir, "dn.xlsx")

	b, err := Create()
	if err != nil {
		t.Fatal(err)
	}
	if err := b.SaveAs(fn); err != nil {
		t.Fatal(err)
	}
	if _, err := b.AddSheet("Data"); err != nil {
		t.Fatal(err)
	}
	// 写入一些数据以便引用
	ws, _ := b.Sheet("Sheet1")
	ws.SetCellValue("A1", 10)
	ws.SetCellValue("A2", 20)
	ws.SetCellValue("A3", 30)

	// 设置全局命名区域
	if err := b.SetDefinedName("MyRange", "Sheet1!$A$1:$A$3"); err != nil {
		t.Fatal(err)
	}
	if got := b.GetDefinedName("MyRange"); got != "Sheet1!$A$1:$A$3" {
		t.Fatalf("GetDefinedName 返回 %q，期望 %q", got, "Sheet1!$A$1:$A$3")
	}
	if err := b.Save(); err != nil {
		t.Fatal(err)
	}
	crossCheckWorkbook(t, fn, false, map[string]string{"MyRange": "Sheet1!$A$1:$A$3"})

	// 重新打开后仍能读到
	b2, err := Open(fn)
	if err != nil {
		t.Fatal(err)
	}
	if got := b2.GetDefinedName("MyRange"); got != "Sheet1!$A$1:$A$3" {
		t.Fatalf("重开后 GetDefinedName 返回 %q，期望 %q", got, "Sheet1!$A$1:$A$3")
	}

	// 替换同名
	if err := b2.SetDefinedName("MyRange", "Sheet1!$A$1:$A$1"); err != nil {
		t.Fatal(err)
	}
	if got := b2.GetDefinedName("MyRange"); got != "Sheet1!$A$1:$A$1" {
		t.Fatalf("替换后 GetDefinedName 返回 %q，期望 %q", got, "Sheet1!$A$1:$A$1")
	}
	if err := b2.Save(); err != nil {
		t.Fatal(err)
	}
	crossCheckWorkbook(t, fn, false, map[string]string{"MyRange": "Sheet1!$A$1:$A$1"})

	// 删除
	if err := b2.DeleteDefinedName("MyRange"); err != nil {
		t.Fatal(err)
	}
	if got := b2.GetDefinedName("MyRange"); got != "" {
		t.Fatalf("删除后 GetDefinedName 应返回空，却得到 %q", got)
	}
	if err := b2.Save(); err != nil {
		t.Fatal(err)
	}
	crossCheckWorkbook(t, fn, false, nil)
}

func TestP1ProtectDefinedNameCoexist(t *testing.T) {
	// 验证工作簿保护与命名区域（包括表级打印区域）可同时写入且均可被 openpyxl 严格解析。
	dir := t.TempDir()
	fn := filepath.Join(dir, "coexist.xlsx")

	b, err := Create()
	if err != nil {
		t.Fatal(err)
	}
	if err := b.SaveAs(fn); err != nil {
		t.Fatal(err)
	}
	ws, _ := b.Sheet("Sheet1")
	if err := ws.SetCellValue("A1", 1); err != nil {
		t.Fatal(err)
	}
	// 表级打印区域（走 print.go 的 helper）
	idx := ws.sheetLocalIndex()
	b.fileMap["xl/workbook.xml"] = setOrReplaceDefinedName(b.fileMap["xl/workbook.xml"], "_xlnm.Print_Area", idx, "Sheet1!$A$1:$A$1")
	// 工作簿级命名区域
	if err := b.SetDefinedName("GlobalName", "Sheet1!$A$1"); err != nil {
		t.Fatal(err)
	}
	// 工作簿结构保护
	if err := b.ProtectWorkbook(""); err != nil {
		t.Fatal(err)
	}
	if err := b.Save(); err != nil {
		t.Fatal(err)
	}

	py := openpyxlValidatePy()
	if py != "" && fileExists(py) {
		script := `
import openpyxl
wb = openpyxl.load_workbook(r'''` + fn + `''')
sec = wb.security
cur = bool(sec.lock_structure) if sec is not None else False
assert cur == True, 'lock_structure 不符: ' + str(cur)
assert wb.defined_names['GlobalName'].value == 'Sheet1!$A$1', '全局命名区域不符: ' + wb.defined_names['GlobalName'].value
# 打印区域为表级命名区域（带 localSheetId），openpyxl 经 ws.print_area 暴露
ws = wb['Sheet1']
pa = ws.print_area
assert pa is not None and '$A$1' in str(pa), '打印区域不符: ' + str(pa)
print('OK')
`
		out, err := exec.Command(py, "-c", script).CombinedOutput()
		if err != nil {
			t.Fatalf("openpyxl 共存交叉验证失败: %v\n%s", err, string(out))
		}
		if !strings.Contains(string(out), "OK") {
			t.Fatalf("共存交叉验证未输出 OK:\n%s", string(out))
		}
	} else {
		t.Skip("未找到 openpyxl 校验 venv，跳过交叉验证")
	}
}

func isValidHex4(s string) bool {
	if len(s) != 4 {
		return false
	}
	for _, c := range s {
		if !((c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f')) {
			return false
		}
	}
	return true
}
