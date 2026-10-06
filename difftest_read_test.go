package excelgo

// 读侧与往返差分：
//
//  1. **openpyxl 写 → excelgo 读**：验证本库能否正确解析 openpyxl 的正常产物。
//     这是兼容性缺口的主要暴露面 —— 用户的文件很可能来自 Excel/WPS/openpyxl，
//     而非本库。若本库读不出这些文件的语义，实用价值大打折扣。
//  2. **excelgo 写 → excelgo 读**：往返一致性，确保写入的语义能原样取回。
//
// 与写侧差分的区别：这里不比较"两侧产物是否相同"（那是写侧的事），
// 而是验证本库**对同一份文件**的解读与 openpyxl 是否一致。

import (
	"fmt"
	"os"
	"path/filepath"
	"strconv"
	"testing"
)

// readCase 一个"openpyxl 写、本库读"的差分场景。
type readCase struct {
	Name string
	// Ops 交给 openpyxl 的操作
	Ops string
	// ExcelGoRead 用本库读取并与 openpyxl 快照的对应字段比对。
	// 返回实际读到的值（归一化为字符串/数字），与 openpyxl 侧对齐即可。
	ExcelGoRead func(t *testing.T, f string, want map[string]interface{})
	SkipReason  string
}

// runOpenpyxlFile 让 openpyxl 产出一份文件，供本库读取。
func runOpenpyxlFile(t *testing.T, opsJSON string) string {
	t.Helper()
	dir := t.TempDir()
	out := filepath.Join(dir, "openpyxl_src.xlsx")
	runOpenpyxlOps(t, out, fmt.Sprintf(`{"out": %q, "ops": [%s]}`, out, opsJSON))
	if _, err := os.Stat(out); err != nil {
		t.Fatalf("openpyxl 未产出文件: %v", err)
	}
	return out
}

func TestDiffReadSide(t *testing.T) {
	cases := []readCase{
		{
			Name: "读-字符串",
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"hello"},
			       {"op":"set_str","sheet":"Sheet1","cell":"A2","value":"中文🚀"}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				assertCellEquals(t, f, "A1", "hello")
				assertCellEquals(t, f, "A2", "中文🚀")
			},
		},
		{
			Name: "读-数值",
			Ops: `{"op":"set_num","sheet":"Sheet1","cell":"A1","value":42},
			       {"op":"set_num","sheet":"Sheet1","cell":"A2","value":-3.5},
			       {"op":"set_num","sheet":"Sheet1","cell":"A3","value":0}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				assertCellEquals(t, f, "A1", "42")
				assertCellEquals(t, f, "A2", "-3.5")
				assertCellEquals(t, f, "A3", "0")
			},
		},
		{
			Name: "读-布尔",
			Ops: `{"op":"set_bool","sheet":"Sheet1","cell":"A1","value":true},
			       {"op":"set_bool","sheet":"Sheet1","cell":"A2","value":false}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				assertCellEquals(t, f, "A1", "TRUE")
				assertCellEquals(t, f, "A2", "FALSE")
			},
		},
		{
			Name: "读-公式",
			Ops:  `{"op":"set_formula","sheet":"Sheet1","cell":"A1","formula":"SUM(B1:B3)"}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				got, err := GetCellValue(f, "Sheet1", "A1")
				if err != nil {
					t.Fatalf("读公式失败: %v", err)
				}
				// 与 openpyxl 对齐：两侧都保留前导 "="（op.py 写入时补 "="，
				// 本库读回也带 "="，快照归一化不再剥离）。
				if s, ok := got.(string); !ok || s != "=SUM(B1:B3)" {
					t.Errorf("公式读回 = %#v，期望 =SUM(B1:B3)", got)
				}
			},
		},
		{
			Name: "读-多表与顺序",
			Ops: `{"op":"create_sheet","name":"表二"},
			       {"op":"create_sheet","name":"表三"},
			       {"op":"set_str","sheet":"表二","cell":"A1","value":"x"}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				names, err := GetSheetList(f)
				if err != nil {
					t.Fatalf("GetSheetList 失败: %v", err)
				}
				if len(names) != 3 || names[0] != "Sheet1" || names[1] != "表二" || names[2] != "表三" {
					t.Errorf("工作表顺序 = %q，期望 [Sheet1 表二 表三]", names)
				}
			},
		},
		{
			Name: "读-合并区域",
			Ops: `{"op":"merge","sheet":"Sheet1","ref":"A1:C1"},
			       {"op":"merge","sheet":"Sheet1","ref":"A3:B4"}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				got, err := GetMergeCells(f, "Sheet1")
				if err != nil {
					t.Fatalf("GetMergeCells 失败: %v", err)
				}
				m := map[string]bool{}
				for _, g := range got {
					m[g] = true
				}
				for _, want := range []string{"A1:C1", "A3:B4"} {
					if !m[want] {
						t.Errorf("缺少合并区域 %s，实际 %q", want, got)
					}
				}
			},
		},
		{
			Name: "读-列宽行高",
			Ops: `{"op":"col_width","sheet":"Sheet1","col":2,"width":25.5},
			       {"op":"row_height","sheet":"Sheet1","row":3,"height":28}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				w, err := GetColWidth(f, "Sheet1", 2)
				if err != nil {
					t.Fatalf("GetColWidth 失败: %v", err)
				}
				if diff := w - 25.5; diff > 0.01 || diff < -0.01 {
					t.Errorf("B 列宽 = %v，期望 25.5", w)
				}
			},
		},
		{
			Name: "读-批注",
			Ops:  `{"op":"comment","sheet":"Sheet1","cell":"A1","text":"批注内容","author":"作者"}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				got, err := GetComment(f, "Sheet1", "A1")
				if err != nil {
					t.Fatalf("GetComment 失败: %v", err)
				}
				if got == "" {
					t.Error("批注读回为空，期望包含「批注内容」")
				}
			},
		},
		{
			Name: "读-定义名称",
			Ops:  `{"op":"defined_name","name":"我的区域","refers_to":"Sheet1!$A$1:$C$10"}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				got, err := GetDefinedName(f, "我的区域")
				if err != nil {
					t.Fatalf("GetDefinedName 失败: %v", err)
				}
				if got == "" {
					t.Error("定义名称读回为空")
				}
			},
		},
		{
			Name: "读-GetRange",
			Ops:  `{"op":"set_range","sheet":"Sheet1","values":[["a","b"],["c","d"]]}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				got, err := GetRange(f, "Sheet1", "A1:B2")
				if err != nil {
					t.Fatalf("GetRange 失败: %v", err)
				}
				if len(got) != 2 || len(got[0]) != 2 {
					t.Fatalf("GetRange 形状 = %v，期望 2x2", got)
				}
				if got[0][0] != "a" || got[1][1] != "d" {
					t.Errorf("GetRange 内容 = %v，期望 [[a b] [c d]]", got)
				}
			},
		},
		{
			Name: "读-样式",
			Ops: `{"op":"set_style","sheet":"Sheet1","cell":"A1",
			       "style":{"bold":true,"numfmt":"0.00"}}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				g, err := Open(f)
				if err != nil {
					t.Fatalf("Open 失败: %v", err)
				}
				ws, err := g.Sheet("Sheet1")
				if err != nil {
					t.Fatalf("Sheet 失败: %v", err)
				}
				st, err := ws.GetStyle("A1")
				if err != nil {
					t.Fatalf("GetStyle 失败: %v", err)
				}
				if st.Font == nil || !st.Font.Bold {
					t.Errorf("粗体未读回: %+v", st.Font)
				}
				if st.NumFmt != "0.00" {
					t.Errorf("数字格式 = %q，期望 0.00", st.NumFmt)
				}
			},
		},
		{
			Name: "读-工作表可见性",
			Ops: `{"op":"create_sheet","name":"暗表"},
			       {"op":"sheet_visible","name":"暗表","state":"hidden"}`,
			ExcelGoRead: func(t *testing.T, f string, want map[string]interface{}) {
				st, err := GetSheetVisible(f, "暗表")
				if err != nil {
					t.Fatalf("GetSheetVisible 失败: %v", err)
				}
				if st != "hidden" {
					t.Errorf("可见性 = %q，期望 hidden", st)
				}
			},
		},
	}

	for _, c := range cases {
		c := c
		t.Run(c.Name, func(t *testing.T) {
			if c.SkipReason != "" {
				t.Skipf("不可比：%s", c.SkipReason)
			}
			if pythonExe() == "" {
				t.Skip("未找到 Python/openpyxl")
			}
			src := runOpenpyxlFile(t, c.Ops)
			c.ExcelGoRead(t, src, nil)
		})
	}
}

// assertCellEquals 断言单元格读回值与期望一致（按本库 GetCellValue 的返回类型归一）。
func assertCellEquals(t *testing.T, f, cell, want string) {
	t.Helper()
	got, err := GetCellValue(f, "Sheet1", cell)
	if err != nil {
		t.Fatalf("读 %s 失败: %v", cell, err)
	}
	if normCellForCompare(got) != want {
		t.Errorf("%s 读回 = %q，期望 %q", cell, normCellForCompare(got), want)
	}
}

// normCellForCompare 把 GetCellValue 的返回值转为可与 openpyxl 快照对齐的字符串。
func normCellForCompare(v interface{}) string {
	switch n := v.(type) {
	case nil:
		return ""
	case string:
		return n
	case bool:
		if n {
			return "TRUE"
		}
		return "FALSE"
	case float64:
		if n == float64(int(n)) {
			return strconv.Itoa(int(n))
		}
		return strconv.FormatFloat(n, 'f', -1, 64)
	case int:
		return strconv.Itoa(n)
	}
	return fmt.Sprint(v)
}

// TestDiffRoundTrip excelgo 写 → excelgo 读：验证往返语义不丢失。
// 差分价值：与 openpyxl 快照对比，能抓出"写进去但读不回来"的不对称。
func TestDiffRoundTrip(t *testing.T) {
	dir := t.TempDir()
	f := filepath.Join(dir, "rt.xlsx")

	if err := Create2(f); err != nil {
		t.Fatal(err)
	}
	// 写入各类数据
	must(t, SetCellStr(f, "Sheet1", "A1", "文本"))
	must(t, SetCellNumeric(f, "Sheet1", "A2", 3.14))
	must(t, SetCellBool(f, "Sheet1", "A3", true))
	must(t, MergeCells(f, "Sheet1", "C1:D1"))
	must(t, SetColWidth(f, "Sheet1", 2, 20))
	must(t, newSheetHelper(f, "第二表"))
	must(t, SetCellStr(f, "第二表", "A1", "另一表"))

	// 与 openpyxl 读同一份文件的结果对比
	snap := runSnapshot(t, f)
	sheets, _ := snap["sheets"].(map[string]interface{})
	s1, _ := sheets["Sheet1"].(map[string]interface{})
	if s1 == nil {
		t.Fatalf("快照缺少 Sheet1: %v", snap)
	}
	cells, _ := s1["cells"].(map[string]interface{})
	for _, tc := range []struct{ cell, wantT, wantV string }{
		{"A1", "str", "文本"},
		{"A3", "bool", "true"},
	} {
		c, ok := cells[tc.cell].(map[string]interface{})
		if !ok {
			t.Errorf("%s 未读回（往返丢失）", tc.cell)
			continue
		}
		if c["t"] != tc.wantT {
			t.Errorf("%s 类型 = %v，期望 %s", tc.cell, c["t"], tc.wantT)
		}
		got, _ := c["v"].(string)
		if tc.cell == "A3" {
			// 布尔在 JSON 里是 bool 而非 string
			if c["v"] != true {
				t.Errorf("A3 值 = %v，期望 true", c["v"])
			}
			continue
		}
		if got != tc.wantV {
			t.Errorf("%s 值 = %q，期望 %q", tc.cell, got, tc.wantV)
		}
	}
	// 数值应为 3.14
	if c, ok := cells["A2"].(map[string]interface{}); ok {
		if v, _ := c["v"].(float64); v != 3.14 {
			t.Errorf("A2 值 = %v，期望 3.14（往返精度损失）", c["v"])
		}
	}
	if _, ok := sheets["第二表"]; !ok {
		t.Error("第二表 往返后丢失")
	}
}

// newSheetHelper 包装 NewSheet，返回 error 以便 must 断言。
func newSheetHelper(f, name string) error {
	_, err := NewSheet(f, name)
	return err
}
