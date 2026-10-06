package excelgo

// 图表创建。
//
// 本库原先只能**搬运**图表部件（复制/合并时保持完整），不能新建 —— 这是
// 相对 openpyxl 最主要的能力缺口。本文件补齐创建侧。
//
// 图表的部件链（与此前修过的多级链对应）：
//
//	sheet.xml  --<drawing r:id>-->  drawingN.xml
//	drawingN.xml --<c:chart r:id>-->  chartN.xml
//
// 因此创建图表要生成三个部件 + 两层 rels + Content_Types override，
// 与"搬运"共用同一套关系管理设施。

import (
	"fmt"
	"path"
	"regexp"
	"strconv"
	"strings"
)

// ChartType 图表类型。
type ChartType string

// 支持的图表类型。取值为 OOXML 的 ST_Chart 家族成员（去掉命名空间前缀）。
const (
	ChartBar      ChartType = "barChart"
	ChartLine     ChartType = "lineChart"
	ChartPie      ChartType = "pieChart"
	ChartScatter  ChartType = "scatterChart"
	ChartArea     ChartType = "areaChart"
	ChartDoughnut ChartType = "doughnutChart"
	ChartRadar    ChartType = "radarChart"
	ChartBubble   ChartType = "bubbleChart"
	ChartSurface  ChartType = "surfaceChart"
	ChartBar3D    ChartType = "bar3DChart"
	ChartLine3D   ChartType = "line3DChart"
	ChartPie3D    ChartType = "pie3DChart"
)

// validChartTypes 是合法图表类型集合。
var validChartTypes = map[ChartType]bool{
	ChartBar: true, ChartLine: true, ChartPie: true, ChartScatter: true,
	ChartArea: true, ChartDoughnut: true, ChartRadar: true, ChartBubble: true,
	ChartSurface: true, ChartBar3D: true, ChartLine3D: true, ChartPie3D: true,
}

// chartTypeNeedsGrouping 标记哪些类型需要 <grouping> 子元素。
// 饼图/雷达/散点/气泡等没有分组概念。
var chartTypeNeedsGrouping = map[ChartType]bool{
	ChartBar: true, ChartLine: true, ChartArea: true,
	ChartBar3D: true, ChartLine3D: true,
}

// ChartSeries 是一个数据系列。
type ChartSeries struct {
	// Name 系列名（如 "2026年"）。为空则引用首行单元格。
	Name string
	// NameRef 系列名所在单元格（如 "Sheet1!$B$1"），优先于 Name。
	NameRef string
	// Categories 分类轴引用（形如 "Sheet1!$A$2:$A$7"）
	Categories string
	// Values 数值引用（形如 "Sheet1!$B$2:$B$7"）
	Values string
}

// ChartOptions 是图表选项。
type ChartOptions struct {
	// Title 图表标题
	Title string
	// Type 图表类型（默认 ChartBar）
	Type ChartType
	// BarDir 条形图方向："col"（柱状，默认）或 "bar"（条形）
	BarDir string
	// Grouping 分组方式："clustered"（默认）/ "stacked" / "percentStacked"
	Grouping string
	// Series 数据系列
	Series []ChartSeries
	// Anchor 图表在工作表中的锚点单元格（如 "E2"）
	Anchor string
	// Width/Height 图表尺寸（像素，默认 480x290，与 Excel 默认一致）
	Width  int
	Height int
	// ShowLegend 是否显示图例（默认 true）
	ShowLegend *bool
	// LegendPos 图例位置："r"/"l"/"t"/"b"/"tr"，默认 "r"
	LegendPos string
}

// AddChart 在 sheetRef 的 opts.Anchor 位置创建图表。
//
// 示意：
//
//	err := excelgo.AddChart("wb.xlsx", "Sheet1", excelgo.ChartOptions{
//	    Type:   excelgo.ChartLine,
//	    Title:  "月度趋势",
//	    Anchor: "E2",
//	    Series: []excelgo.ChartSeries{{
//	        Name:       "产值",
//	        Categories: "Sheet1!$A$2:$A$7",
//	        Values:     "Sheet1!$B$2:$B$7",
//	    }},
//	})
func AddChart(filename, sheetRef string, opts ChartOptions) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.AddChart(opts)
	})
}

// AddChart 在本工作表的 opts.Anchor 位置创建图表。
func (s *WorkSheet) AddChart(opts ChartOptions) error {
	if err := validateChartOptions(&opts); err != nil {
		return err
	}
	fm := s.fm()

	// 1) 生成 chart 部件
	chartNum := nextFreeNumber(fm, "xl/charts/chart", ".xml")
	chartPath := fmt.Sprintf("xl/charts/chart%d.xml", chartNum)
	fm[chartPath] = []byte(buildChartXML(opts))
	fm["[Content_Types].xml"] = insertOverrideInContentTypes(fm["[Content_Types].xml"],
		"/"+chartPath,
		"application/vnd.openxmlformats-officedocument.drawingml.chart+xml")

	// 2) drawing 部件（本表已有则复用其编号，没有则新建）
	drawingPath, drawingRelsPath := ensureDrawingPart(fm, s)
	if drawingPath == "" {
		return fmt.Errorf("无法为工作表 %q 建立 drawing 部件", s.Name())
	}

	// 3) drawing -> chart 的关系
	dRels := loadRels(fm, drawingRelsPath)
	if dRels == nil {
		dRels = &Relationships{XMLName: xmlNameRelationships()}
	}
	chartRID := fmt.Sprintf("rId%d", getMaxRId(dRels)+1)
	dRels.Relationship = append(dRels.Relationship, Relationship{
		ID:     chartRID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
		Target: "/" + chartPath,
	})
	fm[drawingRelsPath] = serializeRelationships(dRels)

	// 4) drawing 里加锚点
	anchorIdx := len(drawingRelsEntries(fm, drawingRelsPath))
	drawing := string(fm[drawingPath])
	drawing = insertChartAnchor(drawing, newRelIDFrom(chartRID), opts, anchorIdx)
	fm[drawingPath] = []byte(drawing)
	fm["[Content_Types].xml"] = insertOverrideInContentTypes(fm["[Content_Types].xml"],
		"/"+drawingPath,
		"application/vnd.openxmlformats-officedocument.drawing+xml")

	// 5) sheet -> drawing 的关系 + <drawing r:id>
	ws := string(s.ws())
	sheetRelsPath := "xl/worksheets/_rels/" + path.Base(s.file) + ".rels"
	sRels := loadRels(fm, sheetRelsPath)
	if sRels == nil {
		sRels = &Relationships{XMLName: xmlNameRelationships()}
	}
	drawingRID := ""
	for _, r := range sRels.Relationship {
		if strings.Contains(r.Type, "/drawing") {
			drawingRID = r.ID
			break
		}
	}
	if drawingRID == "" {
		drawingRID = fmt.Sprintf("rId%d", getMaxRId(sRels)+1)
		sRels.Relationship = append(sRels.Relationship, Relationship{
			ID:     drawingRID,
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
			Target: "/" + drawingPath,
		})
		fm[sheetRelsPath] = serializeRelationships(sRels)
	}
	ws = attachDrawingToSheet(ws, drawingRID)
	s.setWS(ws)
	return nil
}

// drawingRelsEntries 统计 drawing rels 里的关系数（用于生成锚点索引）。
func drawingRelsEntries(fm map[string][]byte, relsPath string) []string {
	r := loadRels(fm, relsPath)
	if r == nil {
		return nil
	}
	var out []string
	for _, rel := range r.Relationship {
		out = append(out, rel.ID)
	}
	return out
}

// validateChartOptions 校验并补默认值。
func validateChartOptions(opts *ChartOptions) error {
	if opts.Type == "" {
		opts.Type = ChartBar
	}
	if !validChartTypes[opts.Type] {
		var valid []string
		for t := range validChartTypes {
			valid = append(valid, string(t))
		}
		return fmt.Errorf("不支持的图表类型 %q，可选：%s",
			opts.Type, strings.Join(valid, ", "))
	}
	if opts.BarDir == "" {
		opts.BarDir = "col"
	}
	if opts.BarDir != "col" && opts.BarDir != "bar" {
		return fmt.Errorf("无效的条形方向 %q，应为 col（柱状）或 bar（条形）", opts.BarDir)
	}
	// 分组方式的合法值**因图表类型而异**：
	//   - 条形图有 clustered（簇状）/ stacked / percentStacked
	//   - 折线/面积图只有 standard / stacked / percentStacked
	//   - 饼图/雷达/散点等没有分组概念
	// 写错（如给折线图写 clustered）会让消费方直接拒绝加载整个工作簿。
	if opts.Grouping == "" {
		if opts.Type == ChartBar || opts.Type == ChartBar3D {
			opts.Grouping = "clustered"
		} else if chartTypeNeedsGrouping[opts.Type] {
			opts.Grouping = "standard"
		}
	}
	if chartTypeNeedsGrouping[opts.Type] {
		valid := []string{"standard", "stacked", "percentStacked"}
		kind := "该图表类型"
		if opts.Type == ChartBar || opts.Type == ChartBar3D {
			valid = append(valid, "clustered")
			kind = "条形图"
		}
		ok := false
		for _, v := range valid {
			if opts.Grouping == v {
				ok = true
				break
			}
		}
		if !ok {
			return fmt.Errorf("无效的分组方式 %q：%s 只接受 %s",
				opts.Grouping, kind, strings.Join(valid, " / "))
		}
	}
	if opts.Width <= 0 {
		opts.Width = 480
	}
	if opts.Height <= 0 {
		opts.Height = 290
	}
	if opts.Anchor == "" {
		opts.Anchor = "E2"
	}
	if newCellRef(opts.Anchor) == "" {
		return fmt.Errorf("无效的锚点单元格: %q", opts.Anchor)
	}
	if len(opts.Series) == 0 {
		return fmt.Errorf("至少需要一个数据系列")
	}
	for i, se := range opts.Series {
		if se.Values == "" {
			return fmt.Errorf("第 %d 个系列缺少 Values（数值引用，如 Sheet1!$B$2:$B$7）", i+1)
		}
		if se.Name == "" && se.NameRef == "" {
			return fmt.Errorf("第 %d 个系列缺少名称（给 Name 或 NameRef 其一）", i+1)
		}
	}
	if opts.LegendPos == "" {
		opts.LegendPos = "r"
	}
	switch opts.LegendPos {
	case "r", "l", "t", "b", "tr":
	default:
		return fmt.Errorf("无效的图例位置 %q，应为 r / l / t / b / tr", opts.LegendPos)
	}
	return nil
}

// ensureDrawingPart 返回该表 drawing 部件与 rels 的路径；没有则新建。
func ensureDrawingPart(fm map[string][]byte, s *WorkSheet) (string, string) {
	// 先看该表是否已有 drawing 关系
	sheetRelsPath := "xl/worksheets/_rels/" + path.Base(s.file) + ".rels"
	if r := loadRels(fm, sheetRelsPath); r != nil {
		for _, rel := range r.Relationship {
			if !strings.Contains(rel.Type, "/drawing") {
				continue
			}
			p := resolveTarget("xl/worksheets", rel.Target)
			if fileExistsInMap(fm, p) {
				return p, path.Join(path.Dir(p), "_rels", path.Base(p)+".rels")
			}
		}
	}
	// 新建 drawing1.xml
	num := nextFreeNumber(fm, "xl/drawings/drawing", ".xml")
	p := fmt.Sprintf("xl/drawings/drawing%d.xml", num)
	// 根元素必须用 "带结束标签" 的形式（非自闭合），否则后续追加的锚点会落在
	// 根元素之外，产出 junk after document element 的坏 XML。
	fm[p] = []byte(xmlDecl +
		`<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ` +
		`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`</xdr:wsDr>`)
	return p, path.Join(path.Dir(p), "_rels", path.Base(p)+".rels")
}

// attachDrawingToSheet 在 sheet 里挂上 <drawing r:id="..."/>。
func attachDrawingToSheet(ws, rid string) string {
	ws = ensureWorksheetNamespaces(ws)
	tag := `<drawing r:id="` + rid + `"/>`
	if strings.Contains(ws, "<drawing ") {
		return regexp.MustCompile(`<drawing\b[^>]*/>`).ReplaceAllString(ws, tag)
	}
	// <drawing> 必须在 sheetData 之后、pageMargins 等之前；直接放在 </worksheet> 前最稳
	if i := strings.LastIndex(ws, "</worksheet>"); i != -1 {
		return ws[:i] + tag + ws[i:]
	}
	return ws + tag
}

// insertChartAnchor 在 drawing 里追加一个 twoCellAnchor 图表锚点。
func insertChartAnchor(drawing string, rID relID, opts ChartOptions, idx int) string {
	col, row, _ := refCoords(opts.Anchor)
	// 由像素尺寸估算跨几列/行（近似即可，Excel 打开后可拖拽调整）
	colSpan := opts.Width / 64
	if colSpan < 1 {
		colSpan = 1
	}
	rowSpan := opts.Height / 20
	if rowSpan < 1 {
		rowSpan = 1
	}
	anchor := `<xdr:twoCellAnchor editAs="oneCell">` +
		`<xdr:from><xdr:col>` + strconv.Itoa(col) + `</xdr:col><xdr:colOff>0</xdr:colOff>` +
		`<xdr:row>` + strconv.Itoa(row) + `</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
		`<xdr:to><xdr:col>` + strconv.Itoa(col+colSpan) + `</xdr:col><xdr:colOff>0</xdr:colOff>` +
		`<xdr:row>` + strconv.Itoa(row+rowSpan) + `</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
		`<xdr:graphicFrame macro="">` +
		`<xdr:nvGraphicFramePr>` +
		`<xdr:cNvPr id="` + strconv.Itoa(idx+2) + `" name="Chart ` + strconv.Itoa(idx+1) + `"/>` +
		`<xdr:cNvGraphicFramePr/>` +
		`</xdr:nvGraphicFramePr>` +
		`<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>` +
		`<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
		`<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
		`r:id="` + string(rID) + `"/>` +
		`</a:graphicData></a:graphic>` +
		`</xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`
	if i := strings.Index(drawing, "</xdr:wsDr>"); i != -1 {
		return drawing[:i] + anchor + drawing[i:]
	}
	return drawing + anchor
}

// buildChartXML 生成 chart 部件内容。
func buildChartXML(opts ChartOptions) string {
	var b strings.Builder
	b.WriteString(xmlDecl)
	b.WriteString(`<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" `)
	b.WriteString(`xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" `)
	b.WriteString(`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)

	// 标题
	if opts.Title != "" {
		b.WriteString(`<c:chart><c:title><c:tx><c:rich>`)
		b.WriteString(`<a:bodyPr/><a:p><a:r><a:t>` + safeText(opts.Title) + `</a:t></a:r></a:p>`)
		b.WriteString(`</c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`)
	} else {
		b.WriteString(`<c:chart><c:autoTitleDeleted val="1"/>`)
	}

	// 图例
	showLegend := true
	if opts.ShowLegend != nil {
		showLegend = *opts.ShowLegend
	}
	if showLegend {
		b.WriteString(`<c:legend><c:legendPos val="` + safeAttr(opts.LegendPos) +
			`"/><c:overlay val="0"/></c:legend>`)
	}

	b.WriteString(`<c:plotArea><c:layout/>`)
	b.WriteString(`<c:` + string(opts.Type) + `>`)

	// 方向与分组（部分类型需要）
	if opts.Type == ChartBar || opts.Type == ChartBar3D {
		b.WriteString(`<c:barDir val="` + safeAttr(opts.BarDir) + `"/>`)
	}
	if chartTypeNeedsGrouping[opts.Type] {
		g := opts.Grouping
		if g == "standard" {
			g = "standard"
		}
		b.WriteString(`<c:grouping val="` + safeAttr(g) + `"/>`)
	}
	if opts.Type == ChartScatter || opts.Type == ChartBubble {
		b.WriteString(`<c:varyColors val="0"/>`)
	}

	// 系列
	for i, se := range opts.Series {
		b.WriteString(`<c:ser>`)
		b.WriteString(`<c:idx val="` + strconv.Itoa(i) + `"/>`)
		b.WriteString(`<c:order val="` + strconv.Itoa(i) + `"/>`)
		// 名称
		if se.NameRef != "" {
			b.WriteString(`<c:tx><c:strRef><c:f>` + safeText(se.NameRef) +
				`</c:f></c:strRef></c:tx>`)
		} else if se.Name != "" {
			b.WriteString(`<c:tx><c:v>` + safeText(se.Name) + `</c:v></c:tx>`)
		}
		// 分类轴
		if se.Categories != "" {
			b.WriteString(`<c:cat><c:strRef><c:f>` + safeText(se.Categories) +
				`</c:f></c:strRef></c:cat>`)
		}
		// 数值轴
		b.WriteString(`<c:val><c:numRef><c:f>` + safeText(se.Values) +
			`</c:f></c:numRef></c:val>`)
		b.WriteString(`</c:ser>`)
	}
	b.WriteString(`</c:` + string(opts.Type) + `>`)
	b.WriteString(`</c:plotArea>`)
	b.WriteString(`<c:plotVisOnly val="1"/>`)
	b.WriteString(`</c:chart>`)
	b.WriteString(`</c:chartSpace>`)
	return b.String()
}

// xmlDecl 是 XML 声明头（等价于 encoding/xml 的 xml.Header，此处内联以避免
// 在字符串拼接里引入包级变量 —— 那会让污点护栏误判为"带污数据进入 XML"）。
const xmlDecl = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
