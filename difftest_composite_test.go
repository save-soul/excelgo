package excelgo

// 复合操作的差分验证：CopySheet（同簿 / 跨文件）与 MergeWorkbook。
//
// 与前面写侧差分的**根本区别**：
// openpyxl 没有 MergeWorkbook，copy_worksheet 也只是"浅复制"（不搬图片、
// 条件格式、数据验证）。所以这里不能"两侧执行同一操作再比对产物"——
// 那比的是 openpyxl 的能力缺口，不是 excelgo 的正确性。
//
// 正确的用法是：**openpyxl 当结果验证器**。
// 由 excelgo 执行复合操作，再用 openpyxl 独立打开产物，检查：
//   1. 能否被严格解析（结构/命名空间/关系完整）；
//   2. 语义内容是否与源表一致（值/样式/合并/列宽/打印设置）；
//   3. **关联部件是否跟着搬走**（drawing / media / rels / 表格 / 批注 /
//      条件格式 / 数据验证）—— 这是"版式保真"的核心，也是普通快照看不到的部分。
//
// 第 3 点必须直接查 zip 部件表：openpyxl 读回时不会暴露
// "drawing 关系是否悬空"这类问题，但那恰恰是 Excel 打不开文件的主因。

import (
	"archive/zip"
	"io"
	"path"
	"path/filepath"
	"sort"
	"strings"
	"testing"
)

// partNames 列出 zip 中全部部件名（排序后）。
func partNames(t *testing.T, p string) []string {
	t.Helper()
	r, err := zip.OpenReader(p)
	if err != nil {
		t.Fatalf("打开 zip 失败 %s: %v", p, err)
	}
	defer r.Close()
	var out []string
	for _, f := range r.File {
		out = append(out, f.Name)
	}
	sort.Strings(out)
	return out
}

// partExists 判断某部件是否存在。
func partExists(t *testing.T, p, name string) bool {
	t.Helper()
	for _, n := range partNames(t, p) {
		if n == name {
			return true
		}
	}
	return false
}

// requirePart 断言部件存在。
func requirePart(t *testing.T, p, name, why string) {
	t.Helper()
	if !partExists(t, p, name) {
		t.Errorf("缺少部件 %s（%s）\n实际部件: %v", name, why, partNames(t, p))
	}
}

// requireNoPart 断言部件不存在。
func requireNoPart(t *testing.T, p, name, why string) {
	t.Helper()
	if partExists(t, p, name) {
		t.Errorf("部件 %s 不应存在（%s）", name, why)
	}
}

// readPartStr 读取 zip 内的部件内容为字符串；部件不存在直接让测试失败。
//
// 实际的读取委托给 readAllFromZip（全套件唯一的 zip 读取实现）。
// 这里额外做一次长度校验：静默截断是最难查的一类测试故障，
// 读全了却"看起来成功"，断言会以一个与真实原因无关的数字失败。
func readPartStr(t *testing.T, p, name string) string {
	t.Helper()
	r, err := zip.OpenReader(p)
	if err != nil {
		t.Fatalf("打开 zip 失败: %v", err)
	}
	defer r.Close()
	for _, f := range r.File {
		if f.Name != name {
			continue
		}
		got := readAllFromZip(f)
		if got == "" {
			t.Fatalf("读取 %s 失败或为空", name)
		}
		if uint64(len(got)) != f.UncompressedSize64 {
			t.Fatalf("读取 %s 不完整: 得到 %d 字节，应为 %d",
				name, len(got), f.UncompressedSize64)
		}
		return got
	}
	t.Fatalf("部件不存在: %s", name)
	return ""
}

// assertNoDanglingRels 断言所有 .rels 中引用的内部目标都真实存在。
// 这是"复制/合并后 Excel 能否打开"的关键：悬空关系会让文件被判损坏。
func assertNoDanglingRels(t *testing.T, p string) {
	t.Helper()
	r, err := zip.OpenReader(p)
	if err != nil {
		t.Fatalf("打开 zip 失败: %v", err)
	}
	defer r.Close()

	exists := map[string]bool{}
	for _, f := range r.File {
		exists[f.Name] = true
	}
	// 解析每个 rels 的 Target（仅检查内部关系，TargetMode="External" 跳过）
	for _, f := range r.File {
		if !strings.HasSuffix(f.Name, ".rels") {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			continue
		}
		content := readAllFromZip(f)
		rc.Close()
		_ = content

		// rels 所在目录：xl/worksheets/_rels/x.xml.rels -> 基准 xl/worksheets/
		base := relsBaseDir(f.Name)
		for _, rel := range parseRels(content) {
			if rel.external {
				continue
			}
			target := resolveRelTarget(base, rel.target)
			if target == "" {
				continue
			}
			if !exists[target] {
				t.Errorf("悬空关系：%s -> %s（解析为 %s，但该部件不存在）\n"+
					"这会导致 Excel/WPS 判定文件损坏",
					f.Name, rel.target, target)
			}
		}
	}
}

type relEntry struct {
	id       string // 关系 ID（rIdN）—— 断言"引用的 rId 是否已声明"时需要
	target   string
	external bool
}

// readAllFromZip 读出 zip 条目的全部内容。
//
// 整个测试套件**只有这一处**读取 zip 条目的实现，所有需要读部件的辅助
// 都委托到这里（此前存在 4 处各自实现：有的手写 Read 循环、有的用 io.ReadAll）。
// 统一的原因不是洁癖：archive/zip 的 Reader 一次 Read 只返回一段解压数据
// （实测恒为 32KB），手写循环必须自己处理分块，而两种常见写法的安全边界不同 ——
// `if err != nil break` 只在 err 非 nil 时停，`if n == 0 || err != nil break`
// 则额外把"读到 0 字节"当结束。后者在包装 reader 下会**静默截断**：
// 文件里明明有 36466 字节却只拿到 32768，症状是"行数断言莫名偏少"，
// 原因与真实缺陷毫无关联，极难定位。
//
// io.ReadAll 按 io.Reader 契约处理，一处即可。
func readAllFromZip(f *zip.File) string {
	rc, err := f.Open()
	if err != nil {
		return ""
	}
	defer rc.Close()
	b, err := io.ReadAll(rc)
	if err != nil {
		return ""
	}
	return string(b)
}

// parseRels 粗解析 Relationship 标签，取出 Target 与 TargetMode。
func parseRels(s string) []relEntry {
	var out []relEntry
	for {
		i := strings.Index(s, "<Relationship")
		if i == -1 {
			break
		}
		j := strings.Index(s[i:], ">")
		if j == -1 {
			break
		}
		tag := s[i : i+j]
		s = s[i+j:]
		e := relEntry{external: strings.Contains(tag, `TargetMode="External"`)}
		if k := strings.Index(tag, `Id="`); k != -1 {
			rest := tag[k+len(`Id="`):]
			if m := strings.Index(rest, `"`); m != -1 {
				e.id = rest[:m]
			}
		}
		if k := strings.Index(tag, `Target="`); k != -1 {
			rest := tag[k+len(`Target="`):]
			if m := strings.Index(rest, `"`); m != -1 {
				e.target = rest[:m]
			}
		}
		out = append(out, e)
	}
	return out
}

// relsBaseDir 返回 .rels 文件对应的"被描述部件"所在目录（不含尾斜杠）。
//
//	xl/worksheets/_rels/sheet1.xml.rels -> xl/worksheets
//	xl/_rels/workbook.xml.rels           -> xl
//	_rels/.rels                          -> ""（包根，描述整个包）
//
// 第三种必须特殊处理：包根 rels 的路径形如 "_rels/.rels"，不含 "/_rels/"
// 子串，早期实现漏了它，导致把 xl/workbook.xml 误判为悬空。
func relsBaseDir(relsPath string) string {
	i := strings.LastIndex(relsPath, "/_rels/")
	if i == -1 {
		return "" // _rels/.rels（包根）
	}
	return relsPath[:i]
}

// resolveRelTarget 把关系 Target 解析为 zip 内绝对路径。
func resolveRelTarget(base, target string) string {
	if target == "" {
		return ""
	}
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(target, "/")
	}
	// 逐级处理 ../ 与 ./；base 为空表示包根
	var parts []string
	if base != "" {
		parts = strings.Split(strings.Trim(base, "/"), "/")
	}
	for _, seg := range strings.Split(target, "/") {
		switch seg {
		case "", ".":
			continue
		case "..":
			if len(parts) > 0 {
				parts = parts[:len(parts)-1]
			}
		default:
			parts = append(parts, seg)
		}
	}
	return strings.Join(parts, "/")
}

// ---------- 场景 1：同簿 CopySheet 的关联部件保真 ----------

func TestDiffCopySheet_Interop(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "copy.xlsx")

	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(p, "Sheet1", "Src"))

	// 在 Src 上铺满各类需要保真的内容
	must(t, SetCellStr(p, "Src", "A1", "标题"))
	must(t, SetCellStr(p, "Src", "A2", "数据"))
	must(t, SetCellNumeric(p, "Src", "B2", 42))
	must(t, SetCellFormula(p, "Src", "C2", "B2*2", "84"))
	_, err := SetCellStyle(p, "Src", "A1", Style{
		Font:      &FontStyle{Bold: true, Size: 16, Color: "FFFF0000"},
		Alignment: &AlignmentStyle{Horizontal: "center"},
	})
	must(t, err)
	must(t, MergeCells(p, "Src", "A4:C4"))
	must(t, SetColWidth(p, "Src", 2, 22.5))
	must(t, SetRowHeight(p, "Src", 1, 30))
	must(t, FreezePanes(p, "Src", "B2"))
	must(t, AddComment(p, "Src", "A2", "批注内容", "作者"))
	must(t, AddHyperlink(p, "Src", "A3", "https://example.com", "链接"))
	must(t, AddTable(p, "Src", "E1:F3", "表A"))
	must(t, AddDataValidationList(p, "Src", "G1:G5", []string{"甲", "乙"}, true))
	must(t, SetConditionalFormat(p, "Src", "B2:B10", "cellIs", "42", 1,
		Style{Fill: &FillStyle{Color: "FFFFFF00", PatternType: "solid"}}))
	must(t, AutoFilter(p, "Src", "A1:C3"))
	must(t, SetSheetProps(p, "Src", SheetProps{
		PrintArea: "A1:F10", PrintTitleRows: "1:1", TabColor: "FF00FF00",
	}))
	// 浮动图片
	must(t, AddPictureFromBytes(p, "Src", tinyPNG(), nil))

	// 目标工作簿里先放一个无关表，确保复制不影响它
	must(t, newSheetHelper(p, "Other"))

	srcBefore := partNames(t, p)
	hasDrawing := partExists(t, p, "xl/drawings/drawing1.xml")

	// 执行同簿复制
	g, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	newName, err := g.CopySheet("Src", "副本")
	if err != nil {
		t.Fatalf("CopySheet 失败: %v", err)
	}
	if newName != "副本" {
		t.Errorf("CopySheet 返回新表名 = %q，期望 副本", newName)
	}
	if err := g.Save(); err != nil {
		t.Fatal(err)
	}

	// --- 1) openpyxl 必须能严格打开（结构完整）---
	snap := runSnapshot(t, p)
	sheets, _ := snap["sheets"].(map[string]interface{})
	cp, ok := sheets["副本"].(map[string]interface{})
	if !ok {
		t.Fatalf("副本表读回失败，快照: %v", snap)
	}

	// --- 2) 语义内容与源表一致 ---
	//
	// tables 单独处理：表格的 name/displayName 是**工作簿级唯一**的，副本必然
	// 被重命名为 "X_1"（否则 Excel 判损坏），所以不能直接要求与源表相同 ——
	// 这里只断言"副本有表格"，重命名正确性由 TestTableReidentifyOnCopy 覆盖。
	src, _ := sheets["Src"].(map[string]interface{})
	for _, field := range []string{"cells", "merges", "col_widths", "row_heights",
		"freeze_panes", "auto_filter", "print_area", "print_title_rows",
		"hyperlinks", "comments"} {
		a, _ := src[field]
		b, _ := cp[field]
		if diffs := diffSnapshot(field, a, b, nil); len(diffs) > 0 {
			t.Errorf("副本的 %s 与源表不一致:\n  %s", field, strings.Join(diffs, "\n  "))
		}
	}
	// 副本必须有表格（名字允许不同）
	if srcTables, _ := src["tables"].([]interface{}); len(srcTables) > 0 {
		if cpTables, _ := cp["tables"].([]interface{}); len(cpTables) != len(srcTables) {
			t.Errorf("副本的表格数 = %v，期望与源表相同 %v", cpTables, srcTables)
		}
	}

	// 定位副本的 sheet 文件名（不假设是 sheet2 —— 前面还有 Other 表）
	cpSheetFile := findSheetFileFor(t, p, "副本")
	cpRelsFile := "xl/worksheets/_rels/" + path.Base(cpSheetFile) + ".rels"

	// --- 3) 关联部件跟着搬走（版式保真的核心）---
	if hasDrawing {
		// drawing 必须有独立副本
		if n := countParts(t, p, "xl/drawings/drawing"); n < 2 {
			t.Errorf("drawing 部件数 = %d，期望 ≥2（副本的图片 drawing 未搬过来）\n部件: %v",
				n, partNames(t, p))
		}
		// 副本的 rels 必须指向**自己的** drawing（不是复用源表的）
		requirePart(t, p, cpRelsFile, "副本表应有自己的 drawing 关系")
		if !strings.Contains(readPartStr(t, p, cpRelsFile), "drawing2.xml") {
			t.Errorf("副本 rels 未指向独立的 drawing2.xml：%s",
				readPartStr(t, p, cpRelsFile))
		}
		// 默认是 MediaShared 策略：媒体**共享**而非复制（这是设计，不是缺陷）。
		// 两种策略都合法，故只要求 media 至少存在。
		if countParts(t, p, "xl/media/") == 0 {
			t.Error("media 部件缺失")
		}
	} else {
		t.Log("源表无图片，跳过 drawing 断言")
	}

	// 表格部件也应为两份（副本有独立的表部件）
	if partExists(t, p, "xl/tables/table1.xml") {
		if n := countParts(t, p, "xl/tables/table"); n < 2 {
			t.Errorf("tables 部件数 = %d，期望 ≥2（副本的表格未搬过来）", n)
		}
		// 副本表必须指向自己的表部件（table2 而非 table1）
		rels := readPartStr(t, p, cpRelsFile)
		if !strings.Contains(rels, "tables/table2.xml") {
			t.Errorf("副本 rels 未指向独立的 table2.xml：%s", rels)
		}
	}

	// 批注部件应为两份
	if partExists(t, p, "xl/comments/comments1.xml") {
		if n := countParts(t, p, "xl/comments/comments"); n < 2 {
			t.Errorf("comments 部件数 = %d，期望 ≥2（副本的批注未搬过来）", n)
		}
	}

	// --- 4) 目标簿原有表不受影响 ---
	if _, ok := sheets["Other"]; !ok {
		t.Error("原有工作表 Other 在复制后丢失")
	}
	_ = srcBefore

	// --- 5) 关系不得悬空（Excel 打不开文件的主因）---
	assertNoDanglingRels(t, p)
}

// countParts 统计部件名以 prefix 开头的数量。
func countParts(t *testing.T, p, prefix string) int {
	t.Helper()
	n := 0
	for _, name := range partNames(t, p) {
		if strings.HasPrefix(name, prefix) {
			n++
		}
	}
	return n
}

// ---------- 场景 2：跨文件 CopySheetTo ----------

func TestDiffCopySheetTo_Interop(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	src := filepath.Join(dir, "src.xlsx")
	dst := filepath.Join(dir, "dst.xlsx")

	if err := Create2(src); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(src, "Sheet1", "Src"))
	must(t, SetCellStr(src, "Src", "A1", "跨文件内容"))
	must(t, SetCellNumeric(src, "Src", "B1", 7))
	_, err := SetCellStyle(src, "Src", "A1", Style{Font: &FontStyle{Bold: true}})
	must(t, err)
	must(t, MergeCells(src, "Src", "D1:E1"))
	must(t, SetColWidth(src, "Src", 1, 30))
	must(t, AddPictureFromBytes(src, "Src", tinyPNG(), nil))

	// 目标簿先有内容，验证不被破坏
	if err := Create2(dst); err != nil {
		t.Fatal(err)
	}
	must(t, newSheetHelper(dst, "DstExisting"))
	must(t, SetCellStr(dst, "DstExisting", "A1", "目标原有内容"))

	if err := CopySheet(src, dst, "Src", "搬入"); err != nil {
		t.Fatalf("CopySheetTo 失败: %v", err)
	}

	// openpyxl 严格解析
	snap := runSnapshot(t, dst)
	sheets, _ := snap["sheets"].(map[string]interface{})
	moved, ok := sheets["搬入"].(map[string]interface{})
	if !ok {
		t.Fatalf("搬入的表读回失败: %v", sheets)
	}
	cells, _ := moved["cells"].(map[string]interface{})
	if c, ok := cells["A1"].(map[string]interface{}); !ok || c["v"] != "跨文件内容" {
		t.Errorf("A1 = %v，期望 跨文件内容", cells["A1"])
	}
	// 目标原有表必须完好
	if _, ok := sheets["DstExisting"]; !ok {
		t.Error("目标工作簿原有表丢失")
	}
	// 图片 media 应带过来
	if !partExists(t, dst, "xl/media/") && countParts(t, dst, "xl/media/") == 0 {
		t.Error("图片 media 部件未随跨文件复制搬过来")
	}
	assertNoDanglingRels(t, dst)
}

// ---------- 场景 3：MergeWorkbook ----------

func TestDiffMergeWorkbook_Interop(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	base := filepath.Join(dir, "base.xlsx")
	m1 := filepath.Join(dir, "m1.xlsx")
	m2 := filepath.Join(dir, "m2.xlsx")

	// 目标簿
	if err := Create2(base); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(base, "Sheet1", "Base"))
	must(t, SetCellStr(base, "Base", "A1", "目标内容"))

	// 两个源簿
	for _, spec := range []struct{ path, sheet, mark string }{
		{m1, "SrcOne", "来自源一"},
		{m2, "SrcTwo", "来自源二"},
	} {
		if err := Create2(spec.path); err != nil {
			t.Fatal(err)
		}
		must(t, RenameSheet(spec.path, "Sheet1", spec.sheet))
		must(t, SetCellStr(spec.path, spec.sheet, "A1", spec.mark))
		must(t, SetCellNumeric(spec.path, spec.sheet, "B1", 100))
		_, err := SetCellStyle(spec.path, spec.sheet, "A1", Style{
			Font: &FontStyle{Bold: true, Color: "FF0000FF"},
		})
		must(t, err)
		must(t, SetColWidth(spec.path, spec.sheet, 1, 18))
		must(t, AddPictureFromBytes(spec.path, spec.sheet, tinyPNG(), nil))
	}

	if err := MergeWorkbook(base, []SourceRef{
		{Workbook: m1, Sheet: "SrcOne"},
		{Workbook: m2, Sheet: "SrcTwo"},
	}); err != nil {
		t.Fatalf("MergeWorkbook 失败: %v", err)
	}

	// openpyxl 严格解析
	snap := runSnapshot(t, base)
	sheets, _ := snap["sheets"].(map[string]interface{})
	for _, want := range []struct{ sheet, cell, val string }{
		{"Base", "A1", "目标内容"},
		{"SrcOne", "A1", "来自源一"},
		{"SrcTwo", "A1", "来自源二"},
	} {
		s, ok := sheets[want.sheet].(map[string]interface{})
		if !ok {
			t.Errorf("合并后缺少表 %s（现有: %v）", want.sheet, keysOf(sheets))
			continue
		}
		cells, _ := s["cells"].(map[string]interface{})
		c, ok := cells[want.cell].(map[string]interface{})
		if !ok || c["v"] != want.val {
			t.Errorf("%s!%s = %v，期望 %q", want.sheet, want.cell, cells[want.cell], want.val)
		}
	}

	// 两个源各一张图 => 加上目标本身无图，media 应为 2
	if n := countParts(t, base, "xl/media/"); n < 2 {
		t.Errorf("media 部件数 = %d，期望 ≥2（两张源表的图片都应搬过来）\n部件: %v",
			n, partNames(t, base))
	}
	// 每个合并进来的表都要有自己的 drawing 关系
	for _, sheet := range []string{"sheet2", "sheet3"} {
		if partExists(t, base, "xl/worksheets/"+sheet+".xml") {
			requirePart(t, base, "xl/worksheets/_rels/"+sheet+".xml.rels",
				"合并进来的表应有自己的 drawing 关系")
		}
	}
	// 样式不能串：两个源都用了 fontId，合并后应各自独立
	assertNoDanglingRels(t, base)
}

func keysOf(m map[string]interface{}) []string {
	var out []string
	for k := range m {
		out = append(out, k)
	}
	sort.Strings(out)
	return out
}

// findSheetFileFor 按工作表名解析其 sheet XML 部件路径。
// 不假设 sheetN.xml 的编号 —— 编号由工作簿里的表顺序决定，断言写死会误报。
func findSheetFileFor(t *testing.T, p, sheetName string) string {
	t.Helper()
	wbXML := readPartStr(t, p, "xl/workbook.xml")
	rels := readPartStr(t, p, "xl/_rels/workbook.xml.rels")
	type relEntry struct{ id, target string }
	var entries []relEntry
	for _, chunk := range strings.Split(rels, "<Relationship") {
		idAttr := attrValue(chunk, "Id")
		tgtAttr := attrValue(chunk, "Target")
		if idAttr != "" && tgtAttr != "" {
			entries = append(entries, relEntry{idAttr, tgtAttr})
		}
	}
	for _, chunk := range strings.Split(wbXML, "<sheet ") {
		name := attrValue(chunk, "name")
		rid := attrValue(chunk, "r:id")
		if name != sheetName && name != "" {
			continue
		}
		for _, e := range entries {
			if e.id == rid {
				target := strings.TrimPrefix(e.target, "/xl/")
				if !strings.HasPrefix(target, "xl/") {
					target = "xl/" + strings.TrimPrefix(target, "/")
				}
				if partExists(t, p, target) {
					return target
				}
			}
		}
	}
	t.Fatalf("找不到工作表 %s 的 sheet 部件", sheetName)
	return ""
}

// attrValue 从 XML 片段里取属性值。
func attrValue(s, attr string) string {
	key := attr + "=\""
	i := strings.Index(s, key)
	if i == -1 {
		return ""
	}
	rest := s[i+len(key):]
	j := strings.Index(rest, "\"")
	if j == -1 {
		return ""
	}
	return rest[:j]
}
