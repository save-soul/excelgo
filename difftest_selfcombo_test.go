package excelgo

import (
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// TestSelfAuthoredCombinations 把**本库自己造**的功能叠加到同一张表上，
// 再跑复合操作。
//
// 与 difftest_combo_test.go 的分工：那边用 openpyxl 生成输入，
// 验证的是"搬运既有文件时能否保真"；这遍验证的是"本库造出来的东西
// 自己能不能扛住后续的复制/合并/流式读"。
//
// 值得单独一层的原因：本库新增的图表创建、命名样式、大纲折叠、数组公式、
// 条件格式多规则**从未在组合场景测过**。而此前抓到的每一个严重缺陷
// （dxfId 悬空、rId 被覆盖、多级链断裂）都是组合场景暴露的 ——
// 单项功能各自都对，叠在一起就出问题。
func TestSelfAuthoredCombinations(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()

	// 造一张"什么都有"的表
	src := filepath.Join(dir, "self.xlsx")
	mustNoErr(t, buildKitchenSinkWorkbook(t, src))

	// 前置：源文件必须确实什么都有，否则下面的断言全是空转
	assertKitchenSinkComplete(t, src)

	t.Run("CopySheet", func(t *testing.T) {
		dst := filepath.Join(t.TempDir(), "cp.xlsx")
		copyFileForTest(t, src, dst)
		g, err := Open(dst)
		mustNoErr(t, err)
		if _, err := g.CopySheet("Sheet1", "副本"); err != nil {
			t.Fatalf("CopySheet 失败: %v", err)
		}
		mustNoErr(t, g.Save())

		// openpyxl 必须能严格解析（结构错乱在这里暴露）
		runSnapshot(t, dst)
		// 副本必须带着全套要素
		assertKitchenSinkComplete(t, dst, "副本")
		assertNoDanglingRels(t, dst)
		assertNoDanglingSheetRefs(t, dst)
		assertTableIdentityUnique(t, dst)
	})

	t.Run("CopySheetTo", func(t *testing.T) {
		dst := filepath.Join(t.TempDir(), "cp2.xlsx")
		g, err := Open(src)
		mustNoErr(t, err)
		mustNoErr(t, g.CopySheetTo(dst, "Sheet1", "搬过来"))
		runSnapshot(t, dst)
		assertKitchenSinkComplete(t, dst, "搬过来")
		assertNoDanglingRels(t, dst)
		assertNoDanglingSheetRefs(t, dst)
		assertTableIdentityUnique(t, dst)
	})

	t.Run("Merge", func(t *testing.T) {
		dst := filepath.Join(t.TempDir(), "merged.xlsx")
		b, err := Create()
		mustNoErr(t, err)
		mustNoErr(t, b.SaveAs(dst))
		mustNoErr(t, b.Merge([]SourceRef{{Workbook: src, Sheet: "Sheet1"}}))
		mustNoErr(t, b.Save())

		runSnapshot(t, dst) // 悬空 dxfId / rId 缺失会在这里抛
		// 合并进来的表名是 "Sheet1_merge1"
		assertKitchenSinkComplete(t, dst, "Sheet1_merge1")
		assertNoDanglingRels(t, dst)
		assertNoDanglingSheetRefs(t, dst)
		assertTableIdentityUnique(t, dst)
	})

	t.Run("StreamReader", func(t *testing.T) {
		sb, err := OpenReader(src)
		mustNoErr(t, err)
		defer sb.Close()
		it, err := sb.StreamRows("Sheet1")
		mustNoErr(t, err)
		defer it.Close()
		rows := 0
		for it.Next() {
			rows++
			// 值必须与常规读一致 —— 组合文件里最容易出问题的是含样式/图表的表
			if rows == 2 {
				want, err := GetCell(src, "Sheet1", "A2")
				mustNoErr(t, err)
				got := ""
				if len(it.Cells()) > 0 {
					got = it.Cells()[0].Value
				}
				if got != want {
					t.Errorf("流式读 A2 = %q，常规读 = %q", got, want)
				}
			}
		}
		mustNoErr(t, it.Error())
		if rows != 10 {
			t.Errorf("流式迭代到 %d 行，期望 10", rows)
		}
	})

	t.Run("二次组合：副本再复制一次", func(t *testing.T) {
		// 组合操作叠加组合操作 —— 双写冲突这类问题只在第二层暴露
		dst := filepath.Join(t.TempDir(), "twice.xlsx")
		copyFileForTest(t, src, dst)
		g, err := Open(dst)
		mustNoErr(t, err)
		if _, err := g.CopySheet("Sheet1", "第一份"); err != nil {
			t.Fatal(err)
		}
		if _, err := g.CopySheet("第一份", "第二份"); err != nil {
			t.Fatal(err)
		}
		mustNoErr(t, g.Save())

		runSnapshot(t, dst)
		assertKitchenSinkComplete(t, dst, "第一份")
		assertKitchenSinkComplete(t, dst, "第二份")
		assertNoDanglingRels(t, dst)
		assertNoDanglingSheetRefs(t, dst)
		assertTableIdentityUnique(t, dst)
	})
}

// buildKitchenSinkWorkbook 造一张把本库各项能力都放上去的工作表。
//
// 刻意让各要素**互相引用**（图表引用表格区域、命名样式被条件格式复用、
// 表格区域里有数组公式），而不是各自独立 —— 独立要素只能验证"单项存在"，
// 互相引用才能验证"索引与关系都对得上"。
func buildKitchenSinkWorkbook(t *testing.T, path string) error {
	t.Helper()
	b, err := Create()
	if err != nil {
		return err
	}
	if err := b.SaveAs(path); err != nil {
		return err
	}
	ws, err := b.Sheet("Sheet1")
	if err != nil {
		return err
	}
	// 注意：**包级 API**（AddTable / AddDataValidationList / AddChart / SetNamedStyle）
	// 各自独立地 open→改→save 这个文件。若在最后再 b.Save()，
	// 会用**旧的内存副本**覆盖掉它们刚写的内容 ——
	// 表现为这些要素在 sheet XML 里凭空消失，且所有调用都返回 nil（无报错）。
	// 因此：本文件全部用**对象式方法**（ws.Xxx / b.Xxx），不用包级 API，
	// 最后一次性 Save。
	_ = b

	// --- 数据：A 列分类、B 列数值、D 列图表数据源 ---
	for i := 1; i <= 10; i++ {
		r := strconv.Itoa(i)
		ws.SetCellStr("A"+r, "类"+r)
		ws.SetCellNumeric("B"+r, float64(i*10))
		ws.SetCellStr("D"+r, "项"+r)
		ws.SetCellNumeric("E"+r, float64(i)*1.5)
	}

	// --- 命名样式（工作簿级） ---
	headStyle, err := ws.NewNamedStyle("表头样式", Style{
		Font:      &FontStyle{Bold: true, Color: "FFFFFFFF", Size: 12},
		Fill:      &FillStyle{Color: "FF4472C4", PatternType: "solid"},
		Alignment: &AlignmentStyle{Horizontal: "center", Vertical: "center"},
	})
	if err != nil {
		return err
	}
	if err := ws.SetNamedStyle("A1:E1", headStyle); err != nil {
		return err
	}

	// --- 表格（区域含命名样式与数组公式） ---
	if err := ws.AddTable("D1:E10", "数据表"); err != nil {
		return err
	}

	// --- 数组公式（放在表格区域内，考验 ref 与单元格样式协同） ---
	if err := ws.SetCellArrayFormula("F2", "F2:F10", "E2:E10*2", ""); err != nil {
		return err
	}

	// --- 条件格式：多种规则类型 + 多区域 ---
	if err := ws.SetConditionalFormatRules("B1:B10", []ConditionalFormatRule{
		{Type: "cellIs", Operator: "lessThan", Formula: "30",
			Style: &Style{Fill: &FillStyle{Color: "FFFFFF00", PatternType: "solid"}}},
		{Type: "cellIs", Operator: "greaterThan", Formula: "80",
			Style: &Style{Font: &FontStyle{Color: "FF0000FF", Bold: true}}},
	}); err != nil {
		return err
	}
	if err := ws.SetConditionalFormatRules("E1:E10", []ConditionalFormatRule{
		{Type: "colorScale", ColorScale: []ColorScalePoint{
			{Type: "min", Color: "FF63BE7B"},
			{Type: "percentile", Value: "50", Color: "FFFFEB84"},
			{Type: "max", Color: "FFF8696B"},
		}},
		{Type: "dataBar", Color: "FF638EC6"},
	}); err != nil {
		return err
	}

	// --- 数据验证 ---
	if err := ws.AddDataValidationList("A2:A10",
		[]string{"类1", "类2", "类3"}, true); err != nil {
		return err
	}

	// --- 图表（引用表格区域） ---
	if err := ws.AddChart(ChartOptions{
		Type:   ChartLine,
		Title:  "趋势",
		Anchor: "H2",
		Series: []ChartSeries{
			{NameRef: "Sheet1!$E$1", Categories: "Sheet1!$D$2:$D$10", Values: "Sheet1!$E$2:$E$10"},
			{Name: "第二组", Categories: "Sheet1!$D$2:$D$10", Values: "Sheet1!$B$2:$B$10"},
		},
	}); err != nil {
		return err
	}

	// --- 第二个图表（考验多图表共存） ---
	if err := ws.AddChart(ChartOptions{
		Type:     ChartBar,
		Title:    "对比",
		Anchor:   "H20",
		BarDir:   "col",
		Grouping: "clustered",
		Series: []ChartSeries{
			{NameRef: "Sheet1!$B$1", Categories: "Sheet1!$A$2:$A$10", Values: "Sheet1!$B$2:$B$10"},
		},
	}); err != nil {
		return err
	}

	// --- 合并单元格 + 打印区域 + 大纲折叠 ---
	if err := ws.MergeCells("A12:C12"); err != nil {
		return err
	}
	if err := ws.SetProps(SheetProps{PrintArea: "A1:F10"}); err != nil {
		return err
	}

	// --- 大纲折叠（第 5~9 行） ---
	if err := ws.GroupRows(5, 9, 1); err != nil {
		return err
	}
	if err := ws.CollapseRows(5, 9); err != nil {
		return err
	}

	return b.Save()
}

// assertKitchenSinkComplete 断言工作簿的各要素齐全。
//
// sheetName 为空时查第一张表。这是"组合完整性"的正向守卫 ——
// 只断言"没报错"不够，要确认真的是需要的都在。
func assertKitchenSinkComplete(t *testing.T, p string, sheetName ...string) {
	t.Helper()
	target := "Sheet1"
	if len(sheetName) > 0 {
		target = sheetName[0]
	}

	// 1) openpyxl 严格解析
	snap := runSnapshot(t, p)
	sheets, _ := snap["sheets"].(map[string]interface{})
	ws, _ := sheets[target].(map[string]interface{})
	if ws == nil {
		t.Errorf("快照里找不到工作表 %s（现有: %v）", target, sheetKeys(sheets))
		return
	}

	// 2) 图表
	if n := chartCountOfSheet(t, p, target); n < 2 {
		t.Errorf("工作表 %s 的图表数 = %d，期望 ≥2（两个图表都应存在）", target, n)
	}

	// 3) 表格
	f := findSheetFileFor(t, p, target)
	if n := len(rIdsIn(readPartStr(t, p, f))); n == 0 {
		t.Errorf("工作表 %s 没有任何关系引用（表格/图表都没了？）", target)
	}

	// 4) 条件格式。快照里是**扁平**结构：每条规则一个元素
	//    {range, type, formula}，不是按区域嵌套 rules 数组。
	cf, _ := ws["conditional_formats"].([]interface{})
	if len(cf) < 4 {
		t.Errorf("工作表 %s 的条件格式规则数 = %d，期望 ≥4", target, len(cf))
	}
	// 规则类型必须覆盖 cellIs / colorScale / dataBar 三类
	seenType := map[string]bool{}
	for _, c := range cf {
		if m, ok := c.(map[string]interface{}); ok {
			if ty, _ := m["type"].(string); ty != "" {
				seenType[ty] = true
			}
		}
	}
	for _, want := range []string{"cellIs", "colorScale", "dataBar"} {
		if !seenType[want] {
			t.Errorf("工作表 %s 缺少 %s 条件格式规则（现有: %v）", target, want, seenType)
		}
	}

	// 5) 数组公式（t="array" + ref）
	if !strings.Contains(readPartStr(t, p, f), `<f t="array"`) {
		t.Errorf("工作表 %s 的数组公式丢失（缺 t=\"array\"）", target)
	}

	// 6) 命名样式
	styles := readPartStr(t, p, "xl/styles.xml")
	if !strings.Contains(styles, "表头样式") {
		t.Errorf("命名样式 表头样式 丢失")
	}
	if !strings.Contains(styles, "<cellStyles") {
		t.Errorf("cellStyles 区块丢失")
	}

	// 7) 数据验证
	if !strings.Contains(readPartStr(t, p, f), "<dataValidation") {
		t.Errorf("工作表 %s 的数据验证丢失", target)
	}

	// 8) 合并单元格
	if !strings.Contains(readPartStr(t, p, f), "<mergeCell") {
		t.Errorf("工作表 %s 的合并单元格丢失", target)
	}

	// 9) 大纲折叠（outlineLevel + hidden）
	sheet := readPartStr(t, p, f)
	if !strings.Contains(sheet, `outlineLevel="1"`) {
		t.Errorf("工作表 %s 的大纲层级丢失", target)
	}
	if !strings.Contains(sheet, `hidden="1"`) {
		t.Errorf("工作表 %s 的折叠隐藏标记丢失", target)
	}
}

// sheetKeys 取快照里 sheets 映射的键，便于失败时报告。
func sheetKeys(m map[string]interface{}) []string {
	out := make([]string, 0, len(m))
	for k := range m {
		out = append(out, k)
	}
	return out
}
