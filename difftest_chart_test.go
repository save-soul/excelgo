package excelgo

// 图表部件链的差分验证。
//
// 图表是**多级关系链**，是最后一个未验证的关联部件环节：
//
//	sheet1.xml  --<drawing r:id>-->  drawing1.xml
//	drawing1.xml  --<c:chart r:id>-->  chart1.xml、chart2.xml
//	chart1.xml  --<c:externalData>-->  （可选：嵌入工作簿）
//
// 单层关系（sheet -> media）此前已验证过，但多级链要额外成立：
//  1. 搬运函数必须**递归**下去，不能只处理一层 rels；
//  2. 每层 rels 的 Target 都要被正确重写（绝对/相对两种形式都要处理）；
//  3. 图表内部的数据缓存（c:numCache / c:strCache）必须完整保留 ——
//     丢了缓存 Excel 打开会显示"图表数据不可用"。
//
// 本库自身不创建图表（只负责搬运），所以测试方式是：用 openpyxl 生成真实图表文件，
// 再让本库复制/合并，检查图表链是否完整 —— 这正是用户实际的使用场景。

import (
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// makeChartWorkbook 用 openpyxl 生成含多级图表链的 xlsx。
func makeChartWorkbook(t *testing.T, outFile string) {
	t.Helper()
	py := pythonExe()
	if py == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	script := filepath.Join(difftestDir(t), "make_charts.py")
	outB, err := execCommand(py, script, outFile).CombinedOutput()
	if err != nil {
		t.Fatalf("生成图表文件失败: %v\n%s", err, outB)
	}
	if !strings.Contains(string(outB), "BUILD_OK") {
		t.Fatalf("生成图表文件未返回 BUILD_OK: %s", outB)
	}
	if !partExists(t, outFile, "xl/charts/chart1.xml") {
		t.Fatalf("生成的文件里没有图表部件: %v", partNames(t, outFile))
	}
}

// chartCount 统计 xl/charts/chart*.xml 数量。
func chartCount(t *testing.T, p string) int {
	t.Helper()
	return countParts(t, p, "xl/charts/chart")
}

// assertChartDataRefIntact 断言图表的数据引用表达式（<c:f>）完整保留。
//
// 注意：**不能断言 numCache/strCache**。实测 openpyxl 自身也不写数据缓存
// （只写 <c:numRef><c:f>Sheet!$B$2:$B$7</c:f>），Excel 打开时会依引用重算。
// 真正要保的是**引用表达式**：它指回原工作表的单元格区域，丢了图表就取不到数据。
func assertChartDataRefIntact(t *testing.T, p string) {
	t.Helper()
	n := 0
	for i := 1; i <= chartCount(t, p); i++ {
		name := "xl/charts/chart" + strconv.Itoa(i) + ".xml"
		if !partExists(t, p, name) {
			continue
		}
		data := readPartStr(t, p, name)
		// openpyxl 用默认命名空间，标签是 <f>（不是 <c:f>）
		if strings.Contains(data, "<f>") {
			n++
		} else {
			t.Errorf("%s 丢失数据引用（<f>）—— Excel 打开后图表取不到数据", name)
		}
	}
	if chartCount(t, p) > 0 && n == 0 {
		t.Errorf("所有图表都丢失了数据引用（chart 数=%d）", chartCount(t, p))
	}
}

// assertChartChainComplete 断言每条 sheet->drawing->chart 链在 rels 上闭合：
// sheet 的 rels 指向的 drawing，其 rels 指向的 chart 部件必须真实存在。
//
// 这是多级链特有的失效方式：一层关系修好了，二层没修。
func assertChartChainComplete(t *testing.T, p string) {
	t.Helper()
	// drawing 部件的 rels 指向的 chart
	for _, name := range partNames(t, p) {
		if !strings.HasPrefix(name, "xl/drawings/_rels/") {
			continue
		}
		base := relsBaseDir(name)
		rels := readPartStr(t, p, name)
		for _, e := range parseRels(rels) {
			if e.external || strings.Contains(e.target, "media/") {
				continue
			}
			if !strings.Contains(e.target, "chart") {
				continue
			}
			resolved := resolveRelTarget(base, e.target)
			if !partExists(t, p, resolved) {
				t.Errorf("图表链断裂：%s -> %s（解析为 %s，该部件不存在）",
					name, e.target, resolved)
			}
		}
	}
	// sheet 的 rels 指向的 drawing
	for _, name := range partNames(t, p) {
		if !strings.HasPrefix(name, "xl/worksheets/_rels/") {
			continue
		}
		base := relsBaseDir(name)
		for _, e := range parseRels(readPartStr(t, p, name)) {
			if e.external || !strings.Contains(e.target, "drawings/") {
				continue
			}
			resolved := resolveRelTarget(base, e.target)
			if !partExists(t, p, resolved) {
				t.Errorf("drawing 链断裂：%s -> %s（解析为 %s，不存在）",
					name, e.target, resolved)
			}
		}
	}
}

// TestDiffChart_CopySheet 验证复制含图表的工作表时，多级图表链完整搬运。
func TestDiffChart_CopySheet(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "chart.xlsx")
	makeChartWorkbook(t, p)

	// 前置：确认源文件确实是多级链
	beforeCharts := chartCount(t, p)
	beforeDrawings := countParts(t, p, "xl/drawings/drawing")
	if beforeCharts < 3 || beforeDrawings < 2 {
		t.Fatalf("源文件图表结构不符预期: charts=%d drawings=%d",
			beforeCharts, beforeDrawings)
	}
	assertChartChainComplete(t, p)

	g, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := g.CopySheet("Data", "Data副本"); err != nil {
		t.Fatalf("CopySheet 失败: %v", err)
	}
	if err := g.Save(); err != nil {
		t.Fatal(err)
	}

	// 1) "Data" 表有 2 个图表（drawing1 指向 chart1+chart2），
	//    "Second" 表有 1 个（drawing2 -> chart3）。只复制 Data，
	//    故期望 chart 数 = 3 + 2 = 5、drawing 数 = 2 + 1 = 3。
	//    这里不写死"翻倍"—— 翻倍取决于被复制表各自带几个图。
	wantCharts := chartCountOfSheet(t, p, "Data") + beforeCharts
	if n := chartCount(t, p); n != wantCharts {
		t.Errorf("图表部件数 = %d，期望 %d（源 %d + 被复制表的 %d）",
			n, wantCharts, beforeCharts, wantCharts-beforeCharts)
	}
	if n := countParts(t, p, "xl/drawings/drawing"); n < beforeDrawings {
		t.Errorf("drawing 部件数 = %d，期望 ≥%d（至少要保留源表的）", n, beforeDrawings)
	}

	// 2) 关系链闭合（核心）
	assertChartChainComplete(t, p)

	// 3) 数据缓存保留
	assertChartDataRefIntact(t, p)

	// 4) 副本的 sheet 必须指向**自己那份** drawing，不能与源表共享
	cpSheet := findSheetFileFor(t, p, "Data副本")
	cpRels := "xl/worksheets/_rels/" + filepathBase(cpSheet) + ".rels"
	if !partExists(t, p, cpRels) {
		t.Fatalf("副本缺少 rels: %s（现有: %v）", cpRels, partNames(t, p))
	}
	cpRelsData := readPartStr(t, p, cpRels)
	srcRelsData := readPartStr(t, p, "xl/worksheets/_rels/sheet1.xml.rels")
	if cpRelsData == srcRelsData {
		t.Errorf("副本 rels 与源表完全相同（未重写 Target）:\n副本: %s", cpRelsData)
	}

	// 5) Content_Types 必须为所有 chart 部件声明 override
	ct := readPartStr(t, p, "[Content_Types].xml")
	for i := 1; chartCount(t, p) >= i; i++ {
		if !strings.Contains(ct, "/xl/charts/chart"+strconv.Itoa(i)+".xml") {
			t.Errorf("[Content_Types].xml 缺少 chart%d.xml 的 Override —— "+
				"Excel 无法识别该图表部件", i)
		}
	}

	// 6) 关系不得悬空 + openpyxl 必须能读
	assertNoDanglingRels(t, p)
	runSnapshot(t, p)
}

// TestDiffChart_MergeWorkbook 验证合并含图表的源工作簿时图表链完整。
func TestDiffChart_MergeWorkbook(t *testing.T) {
	dir := t.TempDir()
	base := filepath.Join(dir, "base.xlsx")
	src := filepath.Join(dir, "src.xlsx")

	if err := Create2(base); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(base, "Sheet1", "Base"))
	must(t, SetCellStr(base, "Base", "A1", "目标"))

	// 源簿带图表
	makeChartWorkbook(t, src)
	// 源簿共 3 个图表，但只合并 "Data" 表（带 2 个），
	// 故期望搬过来的是 Data 自己的那几个，不是全搬。
	wantCharts := chartCountOfSheet(t, src, "Data")
	if wantCharts == 0 {
		t.Fatal("源簿 Data 表没有图表，测试前提不成立")
	}

	if err := MergeWorkbook(base, []SourceRef{{Workbook: src, Sheet: "Data"}}); err != nil {
		t.Fatalf("MergeWorkbook 失败: %v", err)
	}

	// 合并进来的表应带着它的图表部件
	if n := chartCount(t, base); n != wantCharts {
		t.Errorf("合并后图表部件数 = %d，期望 %d（Data 表原有的图表数）", n, wantCharts)
	}
	assertChartChainComplete(t, base)
	assertChartDataRefIntact(t, base)
	assertNoDanglingRels(t, base)
	runSnapshot(t, base)
}

// TestDiffChart_MediaIndependent 验证独立媒体策略下，图表链里的 chart 部件
// 也拿到独立副本（而不是与源表共享同一份 chartN.xml）。
//
// 共享 chartN.xml 意味着改副本的图表会同时改源表 —— 违背"独立"语义。
func TestDiffChart_MediaIndependent(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "chartind.xlsx")
	makeChartWorkbook(t, p)
	before := chartCount(t, p)

	g, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := g.CopySheet("Data", "副本", WithMedia(MediaIndependent)); err != nil {
		t.Fatalf("CopySheet 失败: %v", err)
	}
	if err := g.Save(); err != nil {
		t.Fatal(err)
	}

	// 无论共享还是独立策略，chart 都必须拿到独立副本
	// （共享 chartN.xml 意味着改副本的图表会连带改源表，违背复制语义）
	wantCharts := chartCountOfSheet(t, p, "Data") + before
	if n := chartCount(t, p); n != wantCharts {
		t.Errorf("独立媒体模式下图表部件数 = %d，期望 %d（源 %d + 被复制表的 %d）",
			n, wantCharts, before, wantCharts-before)
	}
	assertChartChainComplete(t, p)
	assertNoDanglingRels(t, p)
}

// filepathBase 是 filepath.Base 的本地别名（避免测试里到处 import path）。
func filepathBase(p string) string {
	if i := strings.LastIndex(p, "/"); i >= 0 {
		return p[i+1:]
	}
	if i := strings.LastIndex(p, `\`); i >= 0 {
		return p[i+1:]
	}
	return p
}

// chartCountOfSheet 统计指定工作表引用的图表数量（经 sheet -> drawing -> chart 链）。
func chartCountOfSheet(t *testing.T, p, sheet string) int {
	t.Helper()
	sheetFile := findSheetFileFor(t, p, sheet)
	relsFile := "xl/worksheets/_rels/" + filepathBase(sheetFile) + ".rels"
	if !partExists(t, p, relsFile) {
		return 0
	}
	n := 0
	for _, e := range parseRels(readPartStr(t, p, relsFile)) {
		if e.external || !strings.Contains(e.target, "drawings/") {
			continue
		}
		drawing := resolveRelTarget("xl/worksheets", e.target)
		dRels := "xl/drawings/_rels/" + filepathBase(drawing) + ".rels"
		if !partExists(t, p, dRels) {
			continue
		}
		for _, c := range parseRels(readPartStr(t, p, dRels)) {
			if !c.external && strings.Contains(c.target, "chart") {
				n++
			}
		}
	}
	return n
}
