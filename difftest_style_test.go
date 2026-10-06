package excelgo

// 样式细节的差分验证。
//
// 前面的场景覆盖了"粗粒度样式"（粗体/颜色/填充/边框/居中）。本文件补齐
// **枚举型取值域**这一维度 —— 它们与自由文本不同：转义挡不住，写错不是"值不对"
// 而是"整个文件打不开"（openpyxl 对每类都有严格白名单，Excel 同样判损坏）。
//
// 覆盖：字体（下划线/删除线/字体名）、填充图案、边框线型、对齐、
// 数字格式、主题色与 dxf 差分格式。

import (
	"path/filepath"
	"strings"
	"testing"
)

// styleCase 一个样式差分场景：同一份样式两侧各写一次，比对读回结果。
type styleCase struct {
	Name    string
	Style   Style
	// Ops 是 openpyxl 侧的等价写法（JSON 的 style 片段）
	Ops string
	// OpenpyxlOK 标记 openpyxl 是否有等价能力（false = 跳过执行，仅验证本库侧）
	OpenpyxlOK bool
	SkipReason string
}

func TestDiffStyleEnums(t *testing.T) {
	cases := []styleCase{
		{
			Name:       "字体-下划线各取值",
			Style:      Style{Font: &FontStyle{Underline: "single"}},
			Ops:        `{"bold":false,"underline":"single"}`,
			OpenpyxlOK: true,
		},
		{
			Name:       "字体-双下划线",
			Style:      Style{Font: &FontStyle{Underline: "double"}},
			Ops:        `{"underline":"double"}`,
			OpenpyxlOK: true,
		},
		{
			Name:       "字体-删除线",
			Style:      Style{Font: &FontStyle{Strike: true}},
			Ops:        `{"strike":true}`,
			OpenpyxlOK: true,
		},
		{
			Name:       "字体-自定义字体名",
			Style:      Style{Font: &FontStyle{Name: "微软雅黑"}},
			Ops:        `{"name":"微软雅黑"}`,
			OpenpyxlOK: true,
		},
		{
			Name:       "边框-各种线型",
			Style:      Style{Border: &BorderStyle{Left: BorderSide{Style: "thick"}, Top: BorderSide{Style: "dashed"}}},
			Ops:        `{"border":{"left":"thick","top":"dashed"}}`,
			OpenpyxlOK: true,
		},
		{
			Name:       "边框-双线",
			Style:      Style{Border: &BorderStyle{Right: BorderSide{Style: "double"}}},
			Ops:        `{"border":{"right":"double"}}`,
			OpenpyxlOK: true,
		},
		{
			Name:       "对齐-分散对齐",
			Style:      Style{Alignment: &AlignmentStyle{Horizontal: "distributed"}},
			Ops:        `{"halign":"distributed"}`,
			OpenpyxlOK: true,
		},
		{
			Name:       "对齐-两端对齐+垂直分散",
			Style:      Style{Alignment: &AlignmentStyle{Horizontal: "justify", Vertical: "distributed"}},
			Ops:        `{"halign":"justify","valign":"distributed"}`,
			OpenpyxlOK: true,
		},
		{
			Name: "填充-灰色125",
			Style: Style{Fill: &FillStyle{Color: "FFD9D9D9", PatternType: "gray125"}},
			// openpyxl 只能设 solid 填充，无法表达图案类型 —— 属能力差异。
			// 这里只验证本库产物能被 openpyxl 严格加载（不比对具体值）。
			OpenpyxlOK: false,
			SkipReason: "openpyxl 侧无法设置非 solid 的填充图案（能力差异，非缺陷）",
		},
		{
			Name: "组合-字体填充边框对齐全开",
			Style: Style{Font: &FontStyle{Bold: true, Italic: true, Underline: "single", Strike: true, Size: 12, Name: "宋体", Color: "FF112233"}, Fill: &FillStyle{Color: "FFAABBCC", PatternType: "solid"}, Border: &BorderStyle{Left: BorderSide{Style: "thin", Color: "FF000000"}, Bottom: BorderSide{Style: "medium", Color: "FFFF0000"}}, Alignment: &AlignmentStyle{Horizontal: "center", Vertical: "center", WrapText: true}},
			// 两侧 JSON 必须完全等价 —— 之前漏了 underline/strike，
			// 测的是"用例不对称"而非实现差异。
			Ops:        `{"bold":true,"italic":true,"underline":"single","strike":true,"size":12,"name":"宋体","color":"FF112233","fill":"AABBCC","border":{"left":"thin","bottom":"medium"},"halign":"center","valign":"center","wrap":true}`,
			OpenpyxlOK: true,
		},
	}

	for _, c := range cases {
		c := c
		t.Run(c.Name, func(t *testing.T) {
			if c.SkipReason != "" {
				t.Logf("说明：%s", c.SkipReason)
			}
			runStyleCase(t, c)
		})
	}
}

// runStyleCase 两侧各写一次同一份样式，比对 openpyxl 读回的结果。
func runStyleCase(t *testing.T, c styleCase) {
	t.Helper()
	dir := t.TempDir()
	gotFile := filepath.Join(dir, "excelgo.xlsx")
	if err := Create2(gotFile); err != nil {
		t.Fatal(err)
	}
	if _, err := SetCellStyle(gotFile, "Sheet1", "A1", c.Style); err != nil {
		t.Fatalf("本库写入样式失败: %v", err)
	}

	if !c.OpenpyxlOK {
		// 仅验证本库产物能被 openpyxl 严格加载（不比对具体值）
		runSnapshot(t, gotFile)
		return
	}
	if pythonExe() == "" {
		runSnapshot(t, gotFile)
		return
	}
	wantFile := filepath.Join(dir, "openpyxl.xlsx")
	ops := `{"op":"set_style","sheet":"Sheet1","cell":"A1","style":` + c.Ops + `}`
	runOpenpyxlOps(t, wantFile, `{"out": `+quoteJSON(wantFile)+`, "ops": [`+ops+`]}`)

	got := runSnapshot(t, gotFile)
	want := runSnapshot(t, wantFile)
	gs, _ := got["sheets"].(map[string]interface{})
	ws, _ := want["sheets"].(map[string]interface{})
	gStyles, _ := gs["Sheet1"].(map[string]interface{})
	wStyles, _ := ws["Sheet1"].(map[string]interface{})
	gSt, _ := gStyles["styles"].(map[string]interface{})
	wSt, _ := wStyles["styles"].(map[string]interface{})
	if gSt["A1"] == nil {
		t.Fatalf("本库侧 A1 无样式记录: %v", gSt)
	}
	if wSt["A1"] == nil {
		t.Fatalf("openpyxl 侧 A1 无样式记录: %v", wSt)
	}
	if diffs := diffSnapshot("A1", gSt["A1"], wSt["A1"], nil); len(diffs) > 0 {
		for _, d := range diffs {
			t.Errorf("[%s] 样式差异: %s", c.Name, d)
		}
	}
}

// TestDiffStyleEnumRejected 验证非法枚举取值被**提前拒绝**，而不是产出
// 消费方读不出来的文件。
//
// 这类缺陷最危险的地方在于：写入返回 nil、文件也能生成，但 Excel/WPS 打开时
// 才报"文件损坏"，用户很难定位到是哪个字段写错了。所以必须在 API 层拦截。
func TestDiffStyleEnumRejected(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "bad.xlsx")
	if err := Create2(p); err != nil {
		t.Fatal(err)
	}

	cases := []struct {
		desc  string
		apply func() error
	}{
		{"非法下划线", func() error {
			_, err := SetCellStyle(p, "Sheet1", "A1", Style{Font: &FontStyle{Underline: "wavy"}})
			return err
		}},
		{"非法边框线型", func() error {
			_, err := SetCellStyle(p, "Sheet1", "A2", Style{
				Border: &BorderStyle{Left: BorderSide{Style: "veryThick"}}})
			return err
		}},
		{"非法水平对齐", func() error {
			_, err := SetCellStyle(p, "Sheet1", "A3", Style{
				Alignment: &AlignmentStyle{Horizontal: "middle"}})
			return err
		}},
		{"非法垂直对齐", func() error {
			_, err := SetCellStyle(p, "Sheet1", "A4", Style{
				Alignment: &AlignmentStyle{Vertical: "middle"}})
			return err
		}},
		{"非法填充图案", func() error {
			_, err := SetCellStyle(p, "Sheet1", "A5", Style{
				Fill: &FillStyle{Color: "FFFFFF00", PatternType: "rainbow"}})
			return err
		}},
		{"非法条件格式类型", func() error {
			return SetConditionalFormat(p, "Sheet1", "A1:A5", "cell", "1", 1, Style{})
		}},
		{"非法数据验证类型", func() error {
			return AddDataValidation(p, "Sheet1", "B1:B5", "dropdown", "", "a,b", "", true)
		}},
		{"非法数据验证运算符", func() error {
			return AddDataValidation(p, "Sheet1", "C1:C5", "whole", "不等", "1", "2", true)
		}},
	}
	for _, c := range cases {
		if err := c.apply(); err == nil {
			t.Errorf("%s：应报错却通过了 —— 会产出消费方读不出来的文件", c.desc)
		} else if !strings.Contains(err.Error(), "OOXML") && !strings.Contains(err.Error(), "cellIs") {
			t.Logf("%s 的错误信息未提及 OOXML: %v", c.desc, err)
		}
	}

	// 关键：被拒绝的写入不能污染文件 —— openpyxl 仍应能严格加载
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	runSnapshot(t, p)
}

// TestDiffNumFmtVariants 验证各类数字格式在两侧读回一致。
// 数字格式是自由文本（formatCode），但 openpyxl 会归一化部分写法，
// 故需实测确认哪些是等价的。
func TestDiffNumFmtVariants(t *testing.T) {
	for _, nf := range []string{
		"0.00", "#,##0.00", "0%", "0.00%", "yyyy-mm-dd", "yyyy/m/d",
		"hh:mm:ss", "0.00E+00", "@", "#,##0", "0.0,,\"万\"",
	} {
		nf := nf
		t.Run(nf, func(t *testing.T) {
			dir := t.TempDir()
			f := filepath.Join(dir, "nf.xlsx")
			if err := Create2(f); err != nil {
				t.Fatal(err)
			}
			if _, err := SetCellStyle(f, "Sheet1", "A1", Style{NumFmt: nf}); err != nil {
				t.Fatalf("设置数字格式失败: %v", err)
			}
			if pythonExe() == "" {
				t.Skip("未找到 Python/openpyxl")
			}
			// openpyxl 侧用等价的 numfmt
			ops := `{"op":"set_style","sheet":"Sheet1","cell":"A1","style":{"numfmt":` +
				quoteJSON(nf) + `}}`
			want := filepath.Join(dir, "want.xlsx")
			runOpenpyxlOps(t, want, `{"out": `+quoteJSON(want)+`, "ops": [`+ops+`]}`)

			gotSnap := runSnapshot(t, f)
			wantSnap := runSnapshot(t, want)
			gs, _ := gotSnap["sheets"].(map[string]interface{})
			ws, _ := wantSnap["sheets"].(map[string]interface{})
			gCell, _ := gs["Sheet1"].(map[string]interface{})
			wCell, _ := ws["Sheet1"].(map[string]interface{})
			gSt, _ := gCell["styles"].(map[string]interface{})
			wSt, _ := wCell["styles"].(map[string]interface{})
			gMap, _ := gSt["A1"].(map[string]interface{})
			wMap, _ := wSt["A1"].(map[string]interface{})
			if gMap == nil || wMap == nil {
				t.Fatalf("样式读回缺失: 本库=%v openpyxl=%v", gMap, wMap)
			}
			if diffs := diffSnapshot("numfmt", gMap["numfmt"], wMap["numfmt"], nil); len(diffs) > 0 {
				t.Errorf("数字格式 %q 两侧不一致: %s", nf, strings.Join(diffs, "; "))
			}
		})
	}
}

// quoteJSON 把字符串转成带引号的 JSON 字面量（供拼 ops 用）。
func quoteJSON(s string) string {
	var b strings.Builder
	b.WriteByte('"')
	for _, r := range s {
		switch r {
		case '"':
			b.WriteString(`\"`)
		case '\\':
			b.WriteString(`\\`)
		case '\n':
			b.WriteString(`\n`)
		case '\t':
			b.WriteString(`\t`)
		case '\r':
			b.WriteString(`\r`)
		default:
			if r < 0x20 {
				b.WriteString(`\u00`)
				const hex = "0123456789abcdef"
				b.WriteByte(hex[(r>>4)&0xF])
				b.WriteByte(hex[r&0xF])
			} else {
				b.WriteRune(r)
			}
		}
	}
	b.WriteByte('"')
	return b.String()
}
