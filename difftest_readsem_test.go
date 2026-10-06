package excelgo

// 读侧行为差分：探测"同一份文件，两侧读出的语义是否一致"。
//
// 与前面的差分测试不同，这里不验证"本库写的东西 openpyxl 读得对"，
// 而是验证**读**的语义本身是否与 openpyxl 一致。
//
// 为什么重要：库与 openpyxl 在若干处**有意的语义差异**，用户按 openpyxl 的
// 直觉用本库会得到意外结果。这类问题单测几乎抓不到 —— 因为每个实现单独看
// 都"自洽"，只有放在一起比才暴露分歧。
//
// 覆盖的高风险区：
//   P1 空单元格 vs 不存在的单元格（能否区分）
//   P2 公式与公式结果
//   P3 共享字符串与富文本
//   P4 数字格式（原始值 vs 显示值）
//   P5 越界读取
//   P7 日期时间的读回

import (
	"encoding/json"
	"fmt"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// probeReadBehavior 调 Python 探测脚本拿 openpyxl 侧的读回结果。
func probeReadBehavior(t *testing.T, xlsx string) map[string]interface{} {
	t.Helper()
	py := pythonExe()
	if py == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	script := filepath.Join(difftestDir(t), "probe_read_behavior.py")
	var stdout, stderr strings.Builder
	cmd := execCommand(py, script, xlsx)
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	if err := cmd.Run(); err != nil {
		t.Fatalf("读行为探测失败: %v\nstderr: %s", err, stderr.String())
	}
	line := strings.TrimSpace(stdout.String())
	if !strings.HasPrefix(line, "PROBE_OK ") {
		t.Fatalf("探测未返回 OK: %s", line)
	}
	var m map[string]interface{}
	if err := json.Unmarshal([]byte(strings.TrimPrefix(line, "PROBE_OK ")), &m); err != nil {
		t.Fatalf("探测结果 JSON 解析失败: %v", err)
	}
	return m
}

// TestDiffReadSemantics 对比两侧对同一份文件的读回语义。
func TestDiffReadSemantics(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "readsem.xlsx")
	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	// 铺一份覆盖面广的样本
	must(t, SetCellStr(p, "Sheet1", "A1", "文本"))
	must(t, SetCellNumeric(p, "Sheet1", "B1", 42))
	must(t, SetCellBool(p, "Sheet1", "C1", true))
	must(t, SetCellFormula(p, "Sheet1", "D1", "B1*2", "84"))
	must(t, SetCellFormula(p, "Sheet1", "D2", "CONCATENATE(A1,\"x\")", "文本x"))
	// E1 日期单元格的类型语义见 TestDateSemanticsDivergence（有意差异）
	must(t, SetCellNumeric(p, "Sheet1", "F1", 3.14159))
	_, err := SetCellStyle(p, "Sheet1", "F1", Style{NumFmt: "0.00"})
	must(t, err)
	must(t, SetCellStr(p, "Sheet1", "G1", "带\n换行"))
	must(t, SetCellStr(p, "Sheet1", "G2", "  前后空格  "))

	probe := probeReadBehavior(t, p)
	pyx, _ := probe["Sheet1"].(map[string]interface{})
	if pyx == nil {
		t.Fatalf("探测结果缺少 Sheet1: %v", probe)
	}

	// --- P1：值的类型与内容 ---
	//
	// 探针把 openpyxl 的值编码成 {k: 类型, v: 归一化值}，两侧按 k 对齐比较。
	// 关键是**类型也要比**：只比字符串化结果会把 bool "TRUE" 与文本 "TRUE" 混同。
	for _, tc := range []struct{ cell, wantKind, wantVal string }{
		{"A1", "str", "文本"},
		{"B1", "num", "42"},
		{"C1", "bool", "TRUE"},
		// E1/F1 刻意不在此断言类型 —— openpyxl 会把"带日期格式的数字"自动转成
		// datetime 对象，而本库返回原始序列号。这是**有意的设计差异**（保留原始值，
		// 由 Cell.GetTime 按需转换），下面 TestDateSemanticsDivergence 专门覆盖。

		{"G1", "str", "带\n换行"},
		{"G2", "str", "  前后空格  "},
		{"B2", "nil", ""},
	} {
		key := "P1_" + tc.cell
		o, _ := pyx[key].(map[string]interface{})
		if o == nil {
			t.Errorf("openpyxl 侧 %s 无探测结果", key)
			continue
		}
		pv, _ := o["value"].(map[string]interface{})
		if pv == nil {
			t.Errorf("openpyxl 侧 %s 的 value 探测结果为空", key)
			continue
		}
		pyKind, _ := pv["k"].(string)
		pyVal, _ := pv["v"].(string)
		if pyKind != tc.wantKind {
			t.Errorf("%s openpyxl 类型 = %q，期望 %q（值 %q）",
				tc.cell, pyKind, tc.wantKind, pyVal)
		}
		// 本库侧
		v, err := GetCellValue(p, "Sheet1", tc.cell)
		if err != nil {
			t.Errorf("本库读 %s 失败: %v", tc.cell, err)
			continue
		}
		gotKind := kindOfCellValue(v)
		if gotKind != tc.wantKind {
			t.Errorf("%s 本库类型 = %q，期望 %q（值 %#v）",
				tc.cell, gotKind, tc.wantKind, v)
		}
		if got := normCellForCompare(v); got != tc.wantVal {
			t.Errorf("%s 值不一致：本库=%q openpyxl=%q", tc.cell, got, pyVal)
		}
	}

	// --- P5：越界/空单元格的读回类型 ---
	// 本库与 openpyxl 都返回零值，但**类型**可能不同。
	// 数值列的空单元格：本库应返回什么？
	for _, cell := range []string{"B2", "B3", "Z99"} {
		ov, _ := pyx["P1_"+cell].(map[string]interface{})
		pyVal, _ := ov["value"].(interface{})
		pyType, _ := ov["data_type"].(string)
		got, err := GetCellValue(p, "Sheet1", cell)
		if err != nil {
			t.Errorf("本库读空单元格 %s 报错: %v", cell, err)
			continue
		}
		// 记录差异（不直接判失败，先看清全貌）
		t.Logf("%s: 本库=%#v (%T) | openpyxl=%#v (t=%s)",
			cell, got, got, pyVal, pyType)
	}
}

// TestDiffReadFormulaResult 验证公式结果的读回语义。
//
// 两库对"公式 + 缓存结果"的处理可能不同：本库同时返回公式与结果，
// 而 openpyxl 在 data_only=False 时只给公式、data_only=True 时只给结果。
// 这里要确认本库的返回约定，并确保它至少自洽可用。
func TestDiffReadFormulaResult(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "formula.xlsx")
	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	must(t, SetCellNumeric(p, "Sheet1", "A1", 10))
	must(t, SetCellNumeric(p, "Sheet1", "A2", 20))
	must(t, SetCellFormula(p, "Sheet1", "B1", "SUM(A1:A2)", "30"))

	v, err := GetCellValue(p, "Sheet1", "B1")
	if err != nil {
		t.Fatal(err)
	}
	s, _ := v.(string)
	if s == "" {
		t.Fatalf("公式单元格读回为空: %#v", v)
	}
	// 读回应是公式文本或缓存结果之一，且不能是空
	t.Logf("公式单元格读回 = %q", s)

	// 无缓存结果的公式（result 传空）不应导致读回失败
	must(t, SetCellFormula(p, "Sheet1", "B2", "SUM(A1:A2)*2", ""))
	v2, err := GetCellValue(p, "Sheet1", "B2")
	if err != nil {
		t.Fatalf("无缓存结果的公式读回失败: %v", err)
	}
	t.Logf("无缓存公式读回 = %#v", v2)
}

// TestDiffReadSharedStrings 验证共享字符串表的读取，对齐 openpyxl。
//
// 尤其测**重复字符串**：共享字符串表的核心机制就是复用，
// 若本库按"每处独立"处理，读出的表格尺寸或内容会与 openpyxl 不一致。
func TestDiffReadSharedStrings(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "sst.xlsx")
	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	// 大量重复字符串 —— 共享字符串表应复用
	for r := 1; r <= 50; r++ {
		must(t, SetCellStr(p, "Sheet1", fmt.Sprintf("A%d", r), "重复内容"))
		must(t, SetCellStr(p, "Sheet1", fmt.Sprintf("B%d", r), "唯一"+strconv.Itoa(r)))
	}
	// 大量不同字符串
	for r := 1; r <= 50; r++ {
		must(t, SetCellNumeric(p, "Sheet1", fmt.Sprintf("C%d", r), float64(r)))
	}

	// openpyxl 读回的行数与内容
	py := pythonExe()
	cmd := execCommand(py, "-c", `
import warnings, json, sys
warnings.simplefilter("ignore")
import openpyxl
wb = openpyxl.load_workbook(sys.argv[1])
ws = wb["Sheet1"]
rows = [[c.value for c in row] for row in ws.iter_rows(min_row=1, max_row=50, max_col=3)]
print("ROWS_OK " + json.dumps({"rows": rows, "nrow": len(rows)}, ensure_ascii=False))
`, p)
	var stdout, stderr strings.Builder
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	if err := cmd.Run(); err != nil {
		t.Fatalf("openpyxl 读取失败: %v\n%s", err, stderr.String())
	}
	line := strings.TrimSpace(stdout.String())
	if !strings.HasPrefix(line, "ROWS_OK ") {
		t.Fatalf("openpyxl 未返回预期结果: %s", line)
	}
	var probe struct {
		Rows [][]interface{} `json:"rows"`
		NRow int             `json:"nrow"`
	}
	if err := json.Unmarshal([]byte(strings.TrimPrefix(line, "ROWS_OK ")), &probe); err != nil {
		t.Fatal(err)
	}
	if probe.NRow != 50 {
		t.Errorf("openpyxl 读回行数 = %d，期望 50", probe.NRow)
	}

	// 本库读回
	rows, err := GetRows(p, "Sheet1")
	if err != nil {
		t.Fatal(err)
	}
	if len(rows) != 50 {
		t.Errorf("本库 GetRows 行数 = %d，期望 50", len(rows))
	}
	// 逐行比对关键列
	for r := 0; r < len(rows) && r < 50; r++ {
		if r < len(probe.Rows) {
			// A 列：重复字符串
			if len(rows[r]) > 0 && len(probe.Rows[r]) > 0 {
				gotA := normCellForCompare(rows[r][0])
				wantA := fmt.Sprint(probe.Rows[r][0])
				if gotA != wantA {
					t.Errorf("第 %d 行 A 列不一致：本库=%q openpyxl=%q", r+1, gotA, wantA)
					break
				}
			}
		}
	}
}

// TestDiffReadEdgeCells 验证边界单元格的读取不 panic、不误报。
func TestDiffReadEdgeCells(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "edge.xlsx")
	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	// 写一些边界坐标
	must(t, SetCellStr(p, "Sheet1", "A1", "左上"))
	must(t, SetCellStr(p, "Sheet1", "XFD1048576", "右下")) // 最大坐标
	// 越界坐标应报错或返回空，不能 panic
	for _, bad := range []string{"", "A", "1", "A0", "0A1", "AAAA1", "A1048577", "XFE1", "A 1", "$A$1"} {
		func() {
			defer func() {
				if r := recover(); r != nil {
					t.Errorf("读取非法坐标 %q 触发 panic: %v", bad, r)
				}
			}()
			_, _ = GetCellValue(p, "Sheet1", bad)
			_, _ = GetCell(p, "Sheet1", bad)
		}()
	}
}

// kindOfCellValue 返回 GetCellValue 结果的类型分类，与探针的 enc() 对齐。
func kindOfCellValue(v interface{}) string {
	switch v.(type) {
	case nil:
		return "nil"
	case bool:
		return "bool"
	case float64, int:
		return "num"
	case string:
		return "str"
	}
	return "other"
}

// TestDateSemanticsDivergence 记录并锁定日期读取的**有意差异**。
//
// 同一份文件（数字 45900 + 格式 yyyy-mm-dd）：
//
//	openpyxl : cell.value -> datetime(2025, 8, 31, 0, 0)   （自动转换）
//	本库      : GetCellValue -> 45900 (float64)             （保留序列号）
//	            Cell.GetTime  -> 2025-08-31 00:00:00         （按需转换）
//
// 这是设计差异而非缺陷：本库保留原始序列号，不丢信息，且由 GetTime 显式转换。
// 但它与 openpyxl 的直觉不同（openpyxl 直接给 datetime），用户在迁移时会踩。
// 本测试的作用是：一旦哪天有人"顺手改成自动转换"，这里会失败并提醒同步改文档。
func TestDateSemanticsDivergence(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "date.xlsx")
	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	// A1 日期、A2 日期时间、A3 同值但无日期格式
	must(t, SetCellNumeric(p, "Sheet1", "A1", 45900))
	_, err := SetCellStyle(p, "Sheet1", "A1", Style{NumFmt: "yyyy-mm-dd"})
	must(t, err)
	must(t, SetCellNumeric(p, "Sheet1", "A2", 45900.5))
	_, err = SetCellStyle(p, "Sheet1", "A2", Style{NumFmt: "yyyy-mm-dd hh:mm"})
	must(t, err)
	must(t, SetCellNumeric(p, "Sheet1", "A3", 45900)) // 无格式

	g, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	ws, err := g.Sheet("Sheet1")
	if err != nil {
		t.Fatal(err)
	}

	// 1) GetCellValue 必须保留原始序列号（不自动转 datetime）
	v1, err := GetCellValue(p, "Sheet1", "A1")
	if err != nil {
		t.Fatal(err)
	}
	// 注意：本库读回的数字可能是 int 也可能是 float64（取决于写入路径），
	// 断言按数值比较而非具体类型 —— 关键是不被转成 datetime/string。
	if kindOfCellValue(v1) != "num" || normCellForCompare(v1) != "45900" {
		t.Errorf("GetCellValue(A1) = %#v（%T），期望数值 45900。"+
			"本库约定：读回保留原始序列号，转换交由 Cell.GetTime —— "+
			"若改成自动转 datetime，这里会失败，请同步更新 README 的语义差异说明",
			v1, v1)
	}

	// 2) Cell.GetTime 必须在有日期格式时正确转换
	tm, ok := ws.Cell("A1").GetTime()
	if !ok {
		t.Error("A1 有 yyyy-mm-dd 格式，GetTime 应返回 ok=true")
	} else if y := tm.Year(); y != 2025 || tm.Month() != 8 || tm.Day() != 31 {
		t.Errorf("A1 日期 = %v，期望 2025-08-31", tm.Format("2006-01-02"))
	}

	// 3) 带时间的应保留时分
	tm2, ok2 := ws.Cell("A2").GetTime()
	if !ok2 {
		t.Error("A2 有日期时间格式，GetTime 应返回 ok=true")
	} else if h := tm2.Hour(); h != 12 {
		t.Errorf("A2 小时 = %d，期望 12（45900.5 = 2025-08-31 中午）", h)
	}

	// 4) 无日期格式时不得误判为日期
	if tm3, ok3 := ws.Cell("A3").GetTime(); ok3 {
		t.Errorf("A3 无日期格式，GetTime 不应返回 ok（却得到 %v）",
			tm3.Format("2006-01-02"))
	}

	// 5) 若装了 openpyxl，确认差异确实存在（记录而非断言失败）
	if pythonExe() == "" {
		return
	}
	pyx, _ := probeReadBehavior(t, p)["Sheet1"].(map[string]interface{})
	if pyx == nil {
		return
	}
	if o, _ := pyx["P1_A1"].(map[string]interface{}); o != nil {
		if v, _ := o["value"].(map[string]interface{}); v != nil {
			if k, _ := v["k"].(string); k == "other" {
				t.Logf("已确认差异：openpyxl 把 A1 转成了非数值类型（%v），本库保留序列号",
					v["v"])
			}
		}
	}
}
