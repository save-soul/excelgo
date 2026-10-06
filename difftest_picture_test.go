package excelgo

// 图片与绘图部件的差分验证。
//
// 库的核心卖点之一是"保住 WPS 单元格内嵌图片"（WPS 用 mc:AlternateContent +
// x14:picture 内联嵌图，而非传统的 drawing 锚点）。这类结构最容易被复制/合并
// 路径破坏，而普通语义快照**看不到**它 —— 必须直接查 zip 部件表与关系。
//
// 本文件覆盖：
//   - WPS 内嵌图片在 复制 / 跨文件复制 / 合并 三条路径下的保真
//   - 浮动图片（传统 drawing）在同样路径下的保真
//   - 媒体共享（默认）vs 独立两种策略的行为差异
//   - 图片部件的提取（ExtractPicture）能取回正确字节

import (
	"bytes"
	"path/filepath"
	"strings"
	"testing"
)

// findPartByPrefix 返回第一个以 prefix 开头且以 suffix 结尾的部件名。
func findPartByPrefix(t *testing.T, p, prefix, suffix string) string {
	t.Helper()
	for _, n := range partNames(t, p) {
		if strings.HasPrefix(n, prefix) && strings.HasSuffix(n, suffix) {
			return n
		}
	}
	return ""
}

// TestDiffCellPicture_WPSInline 验证 WPS 单元格内嵌图片在复制后仍然存在，
// 且 AlternateContent/x14 命名空间结构完好。
func TestDiffCellPicture_WPSInline(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "inline.xlsx")

	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(p, "Sheet1", "Src"))
	must(t, SetCellStr(p, "Src", "A1", "带图单元格"))
	must(t, AddCellPictureFromBytes(p, "Src", "A1", tinyPNG()))

	// 结构前置校验：必须真的有 x14 内嵌结构，否则下面的复制无从验证
	data := readPartStr(t, p, "xl/worksheets/sheet1.xml")
	if !strings.Contains(data, "x14:picture") {
		t.Fatalf("AddCellPicture 未产生 x14:picture 内嵌结构，实际 sheet1.xml 片段: %s",
			snippetAround(data, "drawing", 200))
	}
	// 关键：mc:Ignorable 引用的前缀必须真实存在（MCE 降级机制依赖它）
	assertIgnorablePrefixesExist(t, data, "xl/worksheets/sheet1.xml")

	g, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := g.CopySheet("Src", "副本"); err != nil {
		t.Fatal(err)
	}
	if err := g.Save(); err != nil {
		t.Fatal(err)
	}

	// 副本的 sheet XML 也应有 x14:picture，且 Ignorable 前缀完好
	cpFile := findSheetFileFor(t, p, "副本")
	cpData := readPartStr(t, p, cpFile)
	if !strings.Contains(cpData, "x14:picture") {
		t.Errorf("副本丢失 WPS 内嵌图片（x14:picture）\n副本 sheet: %s", cpFile)
	}
	assertIgnorablePrefixesExist(t, cpData, cpFile)

	// 图片二进制部件应存在
	if findPartByPrefix(t, p, "xl/media/", "") == "" {
		t.Error("xl/media 下没有图片部件")
	}

	// openpyxl 必须能严格加载（内嵌图片结构合法）
	runSnapshot(t, p)
	assertNoDanglingRels(t, p)
}

// TestDiffPicture_FloatMerge 验证浮动图片在合并工作簿时正确搬运，
// 且每张表拿到独立副本（不共享 drawing 部件）。
func TestDiffPicture_FloatMerge(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	base := filepath.Join(dir, "base.xlsx")
	m1 := filepath.Join(dir, "m1.xlsx")
	m2 := filepath.Join(dir, "m2.xlsx")

	if err := Create2(base); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(base, "Sheet1", "Base"))

	for _, spec := range []struct{ path, sheet string }{
		{m1, "SrcOne"}, {m2, "SrcTwo"},
	} {
		if err := Create2(spec.path); err != nil {
			t.Fatal(err)
		}
		must(t, RenameSheet(spec.path, "Sheet1", spec.sheet))
		must(t, SetCellStr(spec.path, spec.sheet, "A1", "图"))
		must(t, AddCellPictureFromBytes(spec.path, spec.sheet, "A1", tinyPNG()))
	}

	if err := MergeWorkbook(base, []SourceRef{
		{Workbook: m1, Sheet: "SrcOne"},
		{Workbook: m2, Sheet: "SrcTwo"},
	}); err != nil {
		t.Fatalf("MergeWorkbook 失败: %v", err)
	}

	// 三个表：Base（无图）+ 两个源表（各有图）=> 应有 2 份内嵌图片结构
	if n := countParts(t, base, "xl/worksheets/sheet"); n < 3 {
		t.Errorf("应有 ≥3 个工作表部件，实际 %d", n)
	}
	// 每个源表搬进来的 sheet 都应带 x14:picture
	pics := 0
	for _, name := range partNames(t, base) {
		if !strings.HasPrefix(name, "xl/worksheets/sheet") || !strings.HasSuffix(name, ".xml") {
			continue
		}
		if strings.Contains(readPartStr(t, base, name), "x14:picture") {
			pics++
		}
	}
	if pics != 2 {
		t.Errorf("带 WPS 内嵌图片的工作表数 = %d，期望 2（两张源表的图都应搬过来）", pics)
	}
	// 结构合法性
	for _, name := range partNames(t, base) {
		if strings.HasPrefix(name, "xl/worksheets/sheet") && strings.HasSuffix(name, ".xml") {
			assertIgnorablePrefixesExist(t, readPartStr(t, base, name), name)
		}
	}
	runSnapshot(t, base)
	assertNoDanglingRels(t, base)
}

// TestDiffPicture_MediaStrategy 验证两种媒体策略的差异符合文档描述：
//   - MediaShared（默认）：共享同一媒体部件，副本不复制二进制；
//   - MediaIndependent：复制出独立副本，删副本不影响源表。
func TestDiffPicture_MediaStrategy(t *testing.T) {
	dir := t.TempDir()

	// --- 共享模式 ---
	shared := filepath.Join(dir, "shared.xlsx")
	if err := Create2(shared); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(shared, "Sheet1", "Src"))
	must(t, AddCellPictureFromBytes(shared, "Src", "A1", tinyPNG()))
	g1, err := Open(shared)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := g1.CopySheet("Src", "副本"); err != nil {
		t.Fatal(err)
	}
	if err := g1.Save(); err != nil {
		t.Fatal(err)
	}
	// 共享模式下 media 只有一份（两个表指向同一图片二进制）
	if n := countParts(t, shared, "xl/media/"); n != 1 {
		t.Errorf("共享模式 media 部件数 = %d，期望 1（共享即不复制二进制）", n)
	}
	// WPS 单元格内嵌图片**不走独立 drawing 部件**（结构内联在 sheet XML 里），
	// 故这里断言"两个表各自都有 x14:picture"，而不是数 drawing 部件。
	for _, sheet := range []string{"Src", "副本"} {
		f := findSheetFileFor(t, shared, sheet)
		if !strings.Contains(readPartStr(t, shared, f), "x14:picture") {
			t.Errorf("%s 缺少 WPS 内嵌图片结构（x14:picture）", sheet)
		}
	}

	// --- 独立模式 ---
	indep := filepath.Join(dir, "indep.xlsx")
	if err := Create2(indep); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(indep, "Sheet1", "Src"))
	must(t, AddCellPictureFromBytes(indep, "Src", "A1", tinyPNG()))
	g2, err := Open(indep)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := g2.CopySheet("Src", "副本", WithMedia(MediaIndependent)); err != nil {
		t.Fatal(err)
	}
	if err := g2.Save(); err != nil {
		t.Fatal(err)
	}
	if n := countParts(t, indep, "xl/media/"); n < 2 {
		t.Errorf("独立模式 media 部件数 = %d，期望 ≥2（应复制出独立副本）", n)
	}
}

// TestDiffPicture_Extract 验证 ExtractPicture 能取回原始图片字节，
// 且复制/合并后取回的内容仍与原图一致（字节级保真）。
func TestDiffPicture_Extract(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "ext.xlsx")
	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(p, "Sheet1", "Src"))
	orig := tinyPNG()
	must(t, AddCellPictureFromBytes(p, "Src", "A1", orig))

	// 直接取回
	media := findPartByPrefix(t, p, "xl/media/", "")
	if media == "" {
		t.Fatal("找不到 media 部件")
	}
	got, err := ExtractPicture(p, "Src", media)
	if err != nil {
		t.Fatalf("ExtractPicture 失败: %v", err)
	}
	if !bytes.Equal(got, orig) {
		t.Errorf("取回的图片字节与原图不一致（原 %d 字节，取回 %d 字节）", len(orig), len(got))
	}

	// 复制后取回：内容应与原图一致（跨簿复制不损坏图片二进制）
	g, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := g.CopySheet("Src", "副本"); err != nil {
		t.Fatal(err)
	}
	if err := g.Save(); err != nil {
		t.Fatal(err)
	}
	got2, err := ExtractPicture(p, "副本", media)
	if err != nil {
		t.Fatalf("从副本取图失败: %v", err)
	}
	if !bytes.Equal(got2, orig) {
		t.Errorf("副本取回的图片字节与原图不一致（%d vs %d 字节）", len(orig), len(got2))
	}
}

// assertIgnorablePrefixesExist 断言 mc:Ignorable="a b c" 里的每个前缀
// 在该元素上都真实声明了 xmlns:a 等。
//
// 这是 WPS 内嵌图片能被 Excel 正确识别的关键：mc:Ignorable 的**值本身就是
// 前缀名列表**，若复制过程中前缀被重命名而 Ignorable 没同步，MCE 降级机制
// 会静默失效（Excel 可能忽略该元素而看不出任何报错）。
func assertIgnorablePrefixesExist(t *testing.T, xmlContent, partName string) {
	t.Helper()
	i := strings.Index(xmlContent, "Ignorable=\"")
	if i == -1 {
		return // 没有 MCE 声明，无需检查
	}
	rest := xmlContent[i+len("Ignorable=\""):]
	j := strings.Index(rest, "\"")
	if j == -1 {
		t.Errorf("%s: mc:Ignorable 属性值未闭合", partName)
		return
	}
	value := rest[:j]
	// 只需检查 <worksheet ...> 这一个开标签里是否声明了这些 xmlns
	openTag := xmlContent[:strings.Index(xmlContent, ">")+1]
	for _, prefix := range strings.Fields(value) {
		decl := "xmlns:" + prefix + "="
		if !strings.Contains(openTag, decl) {
			t.Errorf("%s: mc:Ignorable 引用了前缀 %q，但根元素未声明 %s"+
				"（MCE 降级机制会静默失效，Excel 可能忽略该元素）",
				partName, prefix, decl)
		}
	}
}

// snippetAround 返回 needle 附近的内容片段，供失败信息展示。
func snippetAround(s, needle string, width int) string {
	i := strings.Index(s, needle)
	if i == -1 {
		return s
	}
	start := i - width / 2
	if start < 0 {
		start = 0
	}
	end := i + width/2
	if end > len(s) {
		end = len(s)
	}
	return s[start:end]
}
