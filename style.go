package excelgo

// style.go 提供单元格样式设置能力（纯 Go、不依赖 Office）。
//
// 通过 Style 聚合字体 / 填充 / 边框 / 对齐 / 数字格式等选项，把对应条目写入 xl/styles.xml
// （font/fill/border/numFmt 均按内容去重，避免无限膨胀），生成一条 cellXfs 并返回其 s 索引；
// 再把目标单元格的 s 属性设为该索引（保留其余属性与数据）。
//
// 复用了 styles.go 中的正则块操作原语（extractBlock / extractItems / attrOf / eqElem /
// xfKey / buildXF / replaceBlockText / ensureStyleSheetNS / safeAttr），因此同样不破坏
// styles.xml 其他既有内容。

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// ---------- 样式选项类型 ----------

// FontStyle 字体选项。零值（不含任何设置）表示沿用工作簿默认字体（fontId=0）。
type FontStyle struct {
	Bold      bool
	Italic    bool
	Underline string // "" 无下划线 / "single" / "double" / "singleAccounting" / "doubleAccounting"
	Strike    bool   // 删除线
	Size      float64
	Name      string // 字体名（如 "微软雅黑"），空=默认
	Color     string // ARGB 十六进制（如 "FFFF0000" 红），空=默认
}

// isEmpty 判断字体是否未设置任何属性（此时复用默认 fontId=0）。
func (f *FontStyle) isEmpty() bool {
	if f == nil {
		return true
	}
	return !f.Bold && !f.Italic && f.Underline == "" && !f.Strike && f.Size == 0 && f.Name == "" && f.Color == ""
}

// FillStyle 单元格填充。Color 为空表示无填充（fillId=0）。
type FillStyle struct {
	Color       string // ARGB 前景色（如 "FFFFFF00" 黄）
	PatternType string // 填充图案，默认 "solid"
}

// BorderSide 单边边框。Style 为空表示无边。
type BorderSide struct {
	Style string // "thin"/"medium"/"thick"/"double"/"dashDot"/"hair"/""(无)
	Color string // ARGB 边框色，空=默认黑
}

// BorderStyle 四边 + 对角线边框。全部为空表示无边框（borderId=0）。
type BorderStyle struct {
	Left, Right, Top, Bottom, Diagonal BorderSide
}

func (b *BorderStyle) isEmpty() bool {
	if b == nil {
		return true
	}
	empty := func(s BorderSide) bool { return s.Style == "" }
	return empty(b.Left) && empty(b.Right) && empty(b.Top) && empty(b.Bottom) && empty(b.Diagonal)
}

// AlignmentStyle 对齐方式。
type AlignmentStyle struct {
	Horizontal string // "left"/"center"/"right"/"general"/"fill"/"justify"/"centerContinuous"
	Vertical   string // "top"/"center"/"bottom"/"justify"
	WrapText   bool   // 自动换行
}

// Style 单元格完整样式聚合。任一子项为零值/空时复用样式表中对应默认索引。
type Style struct {
	Font      *FontStyle
	Fill      *FillStyle
	Border    *BorderStyle
	Alignment *AlignmentStyle
	NumFmt    string // 数字格式代码（如 "0.00"、"yyyy-mm-dd"、"0.00%"），空=常规
	// NumFmtId 为该格式对应的 numFmtId（读取样式时填充）。
	// GetTime 依此判定是否为日期格式，比仅凭 NumFmt 文本更可靠。
	NumFmtId int
}

// ---------- 元素构造 ----------

func buildFont(f *FontStyle) string {
	var b strings.Builder
	b.WriteString("<font>")
	if f.Bold {
		b.WriteString("<b/>")
	}
	if f.Italic {
		b.WriteString("<i/>")
	}
	if f.Underline != "" {
		b.WriteString(`<u val="` + safeAttr(f.Underline) + `"/>`)
	}
	if f.Strike {
		b.WriteString("<strike/>")
	}
	if f.Size != 0 {
		b.WriteString(`<sz val="` + strconv.FormatFloat(f.Size, 'f', -1, 64) + `"/>`)
	} else {
		b.WriteString(`<sz val="11"/>`)
	}
	if f.Color != "" {
		b.WriteString(`<color rgb="` + safeAttr(f.Color) + `"/>`)
	}
	if f.Name != "" {
		b.WriteString(`<name val="` + safeAttr(f.Name) + `"/>`)
	}
	b.WriteString("</font>")
	return b.String()
}

// validPatternTypes 是 OOXML 规范（ST_PatternType）认可的填充图案。
// 取值域来自 openpyxl 的严格校验集（该库会拒绝域外取值，导致工作簿无法加载）。
var validPatternTypes = map[string]bool{
	"none": true, "solid": true, "darkDown": true, "darkGray": true,
	"darkGrid": true, "darkHorizontal": true, "darkTrellis": true,
	"darkUp": true, "darkVertical": true, "gray0625": true, "gray125": true,
	"lightDown": true, "lightGray": true, "lightGrid": true,
	"lightHorizontal": true, "lightTrellis": true, "lightUp": true,
	"lightVertical": true, "mediumGray": true,
}

// validatePatternType 校验填充图案类型，返回可直接写入 XML 的安全取值。
// 该值原样写入 patternType 属性，域外取值会让 openpyxl/Excel 拒绝加载工作簿。
func validatePatternType(pt string) (string, error) {
	if pt == "" {
		return "solid", nil
	}
	if !validPatternTypes[pt] {
		return "", fmt.Errorf("无效的填充图案 %q，合法的 OOXML 取值如："+
			"solid / none / gray125 / darkGray / lightGray / mediumGray / "+
			"darkGrid / lightGrid / darkTrellis / lightTrellis"+
			"（写错会导致 openpyxl/Excel 拒绝加载工作簿）", pt)
	}
	return pt, nil
}

func buildFill(f *FillStyle) string {
	pt := f.PatternType
	if pt == "" {
		pt = "solid"
	}
	return `<fill><patternFill patternType="` + safeAttr(pt) + `">` +
		`<fgColor rgb="` + safeAttr(f.Color) + `"/>` +
		`<bgColor indexed="64"/>` +
		`</patternFill></fill>`
}

func buildBorder(b *BorderStyle) string {
	var sb strings.Builder
	sb.WriteString("<border>")
	writeBorderSide(&sb, "left", b.Left)
	writeBorderSide(&sb, "right", b.Right)
	writeBorderSide(&sb, "top", b.Top)
	writeBorderSide(&sb, "bottom", b.Bottom)
	writeBorderSide(&sb, "diagonal", b.Diagonal)
	sb.WriteString("</border>")
	return sb.String()
}

func writeBorderSide(sb *strings.Builder, name borderSideName, side BorderSide) {
	if side.Style == "" {
		sb.WriteString("<" + string(name) + "/>")
		return
	}
	sb.WriteString("<" + string(name) + ` style="` + safeAttr(side.Style) + `">`)
	if side.Color != "" {
		sb.WriteString(`<color rgb="` + safeAttr(side.Color) + `"/>`)
	}
	sb.WriteString("</" + string(name) + ">")
}

func buildXFForStyle(s Style, fontId, fillId, borderId, numFmtId styleIndex) string {
	var b strings.Builder
	b.WriteString(`<xf numFmtId="` + string(numFmtId) + `" fontId="` + string(fontId) +
		`" fillId="` + string(fillId) + `" borderId="` + string(borderId) + `" xfId="0"`)
	b.WriteString(` applyFont="1" applyFill="1" applyBorder="1" applyNumberFormat="1"`)
	if s.Alignment != nil {
		b.WriteString(` applyAlignment="1"`)
	}
	if s.Alignment != nil {
		var ab strings.Builder
		ab.WriteString("<alignment")
		if s.Alignment.Horizontal != "" {
			ab.WriteString(` horizontal="` + safeAttr(s.Alignment.Horizontal) + `"`)
		}
		if s.Alignment.Vertical != "" {
			ab.WriteString(` vertical="` + safeAttr(s.Alignment.Vertical) + `"`)
		}
		if s.Alignment.WrapText {
			ab.WriteString(` wrapText="1"`)
		}
		ab.WriteString(`/>`)
		b.WriteString(">" + ab.String() + "</xf>")
	} else {
		b.WriteString(`/>`)
	}
	return b.String()
}

// ---------- 样式表构建器 ----------

type styleBuilder struct {
	fonts   []string
	fills   []string
	borders []string
	numFmts []string
	cellXfs []string
}

// applyStyleToStylesXML 在已有 styles.xml 字节上应用样式 s，返回新字节与新 cellXf 索引。
func applyStyleToStylesXML(stylesXML []byte, s Style) ([]byte, int, error) {
	// 枚举型字段（填充图案/下划线/边框线型/对齐）原样写入 XML，域外取值会让
	// 消费方拒绝加载工作簿，必须在写入前拦截
	if err := validateStyleEnums(s); err != nil {
		return nil, 0, err
	}
	b := &styleBuilder{
		fonts:   extractItems(extractBlock(string(stylesXML), "fonts"), "font"),
		fills:   extractItems(extractBlock(string(stylesXML), "fills"), "fill"),
		borders: extractItems(extractBlock(string(stylesXML), "borders"), "border"),
		cellXfs: extractItems(extractBlock(string(stylesXML), "cellXfs"), "xf"),
	}
	for _, nf := range extractItems(extractBlock(string(stylesXML), "numFmts"), "numFmt") {
		b.numFmts = append(b.numFmts, nf)
	}

	// 容错：至少保证 1 字体 / 2 填充 / 1 边框 / 1 cellXf
	if len(b.fonts) == 0 {
		b.fonts = append(b.fonts, defaultFont())
	}
	if len(b.fills) == 0 {
		b.fills = append(b.fills, noneFill(), gray125Fill())
	} else if len(b.fills) == 1 {
		b.fills = append(b.fills, gray125Fill())
	}
	if len(b.borders) == 0 {
		b.borders = append(b.borders, defaultBorder())
	}
	if len(b.cellXfs) == 0 {
		b.cellXfs = append(b.cellXfs, defaultXF())
	}

	// 解析各组件索引（按内容去重）
	fontId := "0"
	if !s.Font.isEmpty() {
		fontId = itoa(ensureItem(&b.fonts, buildFont(s.Font)))
	}
	fillId := "0"
	if s.Fill != nil && s.Fill.Color != "" {
		// 图案类型是原样写入 XML 的枚举值，域外取值会让消费方拒绝加载工作簿
		cp := *s.Fill
		pt, err := validatePatternType(cp.PatternType)
		if err != nil {
			return nil, 0, err
		}
		cp.PatternType = pt
		fillId = itoa(ensureItem(&b.fills, buildFill(&cp)))
	}
	borderId := "0"
	if !s.Border.isEmpty() {
		borderId = itoa(ensureItem(&b.borders, buildBorder(s.Border)))
	}
	numFmtId := "0"
	if s.NumFmt != "" {
		id, err := b.ensureNumFmt(s.NumFmt)
		if err != nil {
			return nil, 0, err
		}
		numFmtId = id
	}

	// 生成 cellXf，去重后返回索引
	newXF := buildXFForStyle(s, idx(fontId), idx(fillId), idx(borderId), idx(numFmtId))
	if idx, ok := findEqualXF(b.cellXfs, newXF); ok {
		return rebuildStyles(stylesXML, b), idx, nil
	}
	idx := len(b.cellXfs)
	b.cellXfs = append(b.cellXfs, newXF)
	return rebuildStyles(stylesXML, b), idx, nil
}

// ensureItem 把 elem 去重追加到 items，返回最终下标（已存在则复用）。
func ensureItem(items *[]string, elem string) int {
	norm := normalizeWhitespace(elem)
	for i, it := range *items {
		if normalizeWhitespace(it) == norm {
			return i
		}
	}
	*items = append(*items, elem)
	return len(*items) - 1
}

// ensureNumFmt 确保数字格式代码已存在于 numFmts（去重），返回 id（自定义从 164 起）。
func (b *styleBuilder) ensureNumFmt(code string) (string, error) {
	for _, e := range b.numFmts {
		if attrOf(e, "formatCode") == code {
			return attrOf(e, "numFmtId"), nil
		}
	}
	id := 164
	for b.hasNumFmtID(itoa(id)) {
		id++
	}
	b.numFmts = append(b.numFmts, `<numFmt numFmtId="`+itoa(id)+`" formatCode="`+safeAttr(code)+`"/>`)
	return itoa(id), nil
}

func (b *styleBuilder) hasNumFmtID(id string) bool {
	for _, e := range b.numFmts {
		if attrOf(e, "numFmtId") == id {
			return true
		}
	}
	return false
}

// rebuildStyles 用构建器内最新切片重建 styles.xml 各块（必要时插入缺失块）。
func rebuildStyles(stylesXML []byte, b *styleBuilder) []byte {
	out := string(stylesXML)
	out = ensureStyleSheetNS(out)
	if !strings.HasPrefix(out, `<?xml`) {
		out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + out
	}
	out = upsertBlock(out, "numFmts", b.numFmts, "numFmt", func(n int) string {
		return `<numFmts count="` + itoa(n) + `">`
	})
	out = upsertBlock(out, "fonts", b.fonts, "font", func(n int) string {
		return `<fonts count="` + itoa(n) + `">`
	})
	out = upsertBlock(out, "fills", b.fills, "fill", func(n int) string {
		return `<fills count="` + itoa(n) + `">`
	})
	out = upsertBlock(out, "borders", b.borders, "border", func(n int) string {
		return `<borders count="` + itoa(n) + `">`
	})
	out = upsertBlock(out, "cellXfs", b.cellXfs, "xf", func(n int) string {
		return `<cellXfs count="` + itoa(n) + `">`
	})
	return []byte(out)
}

// upsertBlock 若 styles 含该块则替换；否则插入到 <fonts> 之前（保持 OOXML 子元素合法顺序）。
func upsertBlock(styles string, tag xmlTagName, items []string, itemTag xmlTagName, makeOpen func(int) string) string {
	block := extractBlock(styles, tag)
	body := makeOpen(len(items))
	for _, it := range items {
		body += it
	}
	body += `</` + string(tag) + `>`
	if block == "" {
		anchor := "<fonts>"
		idx := strings.Index(styles, anchor)
		if idx == -1 {
			idx = strings.Index(styles, "<styleSheet")
			if idx != -1 {
				// 插在 <styleSheet ...> 起始标签之后
				rootEnd := strings.Index(styles[idx:], ">")
				idx += rootEnd + 1
			} else {
				return styles // 无法定位，原样返回
			}
		}
		return styles[:idx] + body + styles[idx:]
	}
	return replaceBlockText(styles, tag, func() string { return body })
}

// ---------- 默认条目 ----------

func defaultFont() string {
	return `<font><sz val="11"/><name val="Calibri"/><family val="2"/></font>`
}
func noneFill() string {
	return `<fill><patternFill patternType="none"/></fill>`
}
func gray125Fill() string {
	return `<fill><patternFill patternType="gray125"/></fill>`
}
func defaultBorder() string {
	return `<border><left/><right/><top/><bottom/><diagonal/></border>`
}
func defaultXF() string {
	return `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`
}

// defaultStylesXML 返回最小合法的默认 styles.xml（不含自定义样式，供缺失时创建）。
func defaultStylesXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		`<fonts count="1">` + defaultFont() + `</fonts>` +
		`<fills count="2">` + noneFill() + gray125Fill() + `</fills>` +
		`<borders count="1">` + defaultBorder() + `</borders>` +
		`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
		`<cellXfs count="1">` + defaultXF() + `</cellXfs>` +
		`</styleSheet>`
}

// ---------- 对外 API ----------

// SetCellStyle 把样式 s 应用到 sheetRef 工作表的 cell 单元格（只改 s 索引，保留数据与其它属性）。
// 若工作簿尚无 styles.xml，会自动创建默认样式表。返回被赋予的 s 索引（通常用不到）。
func SetCellStyle(filename, sheetRef, cell string, s Style) (int, error) {
	return SetCellStyleRange(filename, sheetRef, cell, s)
}

// SetCellStyleRange 把样式 s 应用到 sheetRef 工作表的 rangeRef 矩形区域内每个单元格。
// rangeRef 可单格（如 "A1"）或区域（如 "A1:C3"）。所有受影响单元格共享同一 s 索引。
func SetCellStyleRange(filename, sheetRef, rangeRef string, s Style) (int, error) {
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return 0, err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return 0, err
	}
	file, _, _, _, err := locateSheetInMap(fileMap, wb, sheetRef)
	if err != nil {
		return 0, err
	}
	idx, err := setStyleRangeInMap(fileMap, file, rangeRef, s)
	if err != nil {
		return 0, err
	}
	if err := writeMapToZip(filename, fileMap); err != nil {
		return 0, err
	}
	return idx, nil
}

// setStyleRangeInMap 在内存 fileMap 上把样式 s 应用到 file 工作表的 rangeRef 区域
// （不复写磁盘）。返回被赋予的 s 索引。
func setStyleRangeInMap(fileMap map[string][]byte, file, rangeRef string, s Style) (int, error) {
	// 确保 styles.xml 存在
	stylesPath := "xl/styles.xml"
	if _, ok := fileMap[stylesPath]; !ok {
		fileMap[stylesPath] = []byte(defaultStylesXML())
		fileMap["[Content_Types].xml"] = insertOverrideInContentTypes(
			fileMap["[Content_Types].xml"],
			"/"+stylesPath,
			"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml",
		)
	}
	newStyles, idx, err := applyStyleToStylesXML(fileMap[stylesPath], s)
	if err != nil {
		return 0, err
	}
	fileMap[stylesPath] = newStyles

	// 解析区域并逐格设置 s 索引
	c1, r1, c2, r2, err := parseRangeRef(rangeRef)
	if err != nil {
		return 0, err
	}
	for row := r1; row <= r2; row++ {
		for col := c1; col <= c2; col++ {
			cellRef := colNumToLetters(col) + strconv.Itoa(row+1)
			if err := setCellStyleAttrInMap(fileMap, file, cellRef, idx); err != nil {
				return 0, err
			}
		}
	}
	return idx, nil
}

// setCellStyleAttrInMap 在内存 map 上把 cell 的 s 属性设为 sIndex（保留其它属性/数据）。
// 单元格不存在则创建（仅带 s 索引、空值）。
func setCellStyleAttrInMap(fileMap map[string][]byte, file, cell string, sIndex int) error {
	// 单元格坐标先校验规范化：非法坐标直接拒绝，不写入 XML
	cref := newCellRef(cell)
	if cref == "" {
		return fmt.Errorf("无效的单元格引用: %q", cell)
	}
	cell = string(cref)
	ws := string(fileMap[file])
	sAttr := fmt.Sprintf(` s="%d"`, sIndex)
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + regexp.QuoteMeta(cell) + `"[^>]*/?>(?:.*?</c>)?`)
	loc := re.FindStringIndex(ws)
	if loc != nil {
		oldCell := ws[loc[0]:loc[1]]
		newCell := injectSAttr(oldCell, sAttr)
		ws = ws[:loc[0]] + newCell + ws[loc[1]:]
	} else {
		_, rowNum, err := parseCellRef(cell)
		if err != nil {
			return err
		}
		newCell := injectSAttr(`<c r="`+cell+`"/>`, sAttr)
		ws = insertCellIntoRow(part(ws), rowNum+1, frag(newCell))
	}
	fileMap[file] = []byte(ws)
	return nil
}

// styleIndex 表示 styles.xml 中的样式索引（fontId/fillId/borderId/numFmtId）。
// 由库内 buildStyles 递增生成，不含用户数据；写入 XML 属性时无需转义。
type styleIndex string

// idx 标记一个样式索引值。
func idx(s string) styleIndex { return styleIndex(s) }

// borderSideName 表示边框的边名常量（left/right/top/bottom/diagonal），
// 来自库内固定集合，非用户输入。
type borderSideName string

// ---------- 样式枚举值的合法性校验 ----------
//
// 下面这些属性都是**原样写入 XML 的枚举值**，不是自由文本。域外取值不会被
// 转义拦住，而是会让消费方直接拒绝加载整个工作簿（openpyxl 对每类都有严格
// 白名单，Excel 同样判损坏）。因此必须在写入前拦截。

// validUnderlines 是 OOXML 规范（ST_UnderlineValues）认可的下划线取值。
var validUnderlines = map[string]bool{
	"single": true, "double": true, "singleAccounting": true, "doubleAccounting": true,
}

// validBorderStyles 是 OOXML 规范（ST_BorderStyle）认可的边框线型。
var validBorderStyles = map[string]bool{
	"thin": true, "medium": true, "thick": true, "double": true, "dashed": true,
	"dotted": true, "thickDashDot": true, "dashDot": true, "dashDotDot": true,
	"mediumDashDot": true, "mediumDashDotDot": true, "mediumUpDash": true,
	"hair": true, "mediumDashed": true, "dashDotStroked": true, "slantDashDot": true,
	"none": true,
}

// validHorizontals 是 OOXML 规范（ST_HorizontalAlignment）认可的水平对齐取值。
var validHorizontals = map[string]bool{
	"left": true, "center": true, "right": true, "general": true,
	"fill": true, "justify": true, "centerContinuous": true, "distributed": true,
}

// validVerticals 是 OOXML 规范（ST_VerticalAlignment）认可��垂直对齐取值。
var validVerticals = map[string]bool{
	"top": true, "center": true, "bottom": true, "justify": true, "distributed": true,
}

// validateUnderline 校验下划线取值（空串表示无下划线，合法）。
func validateUnderline(u string) error {
	if u == "" || validUnderlines[u] {
		return nil
	}
	return fmt.Errorf("无效的下划线类型 %q，合法的 OOXML 取值：single / double / "+
		"singleAccounting / doubleAccounting（写错会导致 openpyxl/Excel 拒绝加载工作簿）", u)
}

// validateBorderStyle 校验边框线型（空串表示无边框，合法）。
func validateBorderStyle(s string) error {
	if s == "" || validBorderStyles[s] {
		return nil
	}
	return fmt.Errorf("无效的边框线型 %q，合法的 OOXML 取值如：thin / medium / thick / "+
		"double / dashed / dotted / hair（写错会导致 openpyxl/Excel 拒绝加载工作簿）", s)
}

// validateAlignment 校验水平与垂直对齐取值（空串表示默认，合法）。
func validateAlignment(a *AlignmentStyle) error {
	if a == nil {
		return nil
	}
	if a.Horizontal != "" && !validHorizontals[a.Horizontal] {
		return fmt.Errorf("无效的水平对齐 %q，合法的 OOXML 取值：left / center / right / "+
			"general / fill / justify / centerContinuous / distributed", a.Horizontal)
	}
	if a.Vertical != "" && !validVerticals[a.Vertical] {
		return fmt.Errorf("无效的垂直对齐 %q，合法的 OOXML 取值：top / center / bottom / "+
			"justify / distributed", a.Vertical)
	}
	return nil
}

// validateFont 校验字体里的枚举取值。
func validateFont(f *FontStyle) error {
	if f == nil {
		return nil
	}
	return validateUnderline(f.Underline)
}

// validateBorder 校验边框各边的线型取值。
func validateBorder(b *BorderStyle) error {
	if b == nil {
		return nil
	}
	for _, side := range []struct {
		name  string
		style string
	}{
		{"Left", b.Left.Style}, {"Right", b.Right.Style},
		{"Top", b.Top.Style}, {"Bottom", b.Bottom.Style},
		{"Diagonal", b.Diagonal.Style},
	} {
		if err := validateBorderStyle(side.style); err != nil {
			return fmt.Errorf("边框 %s: %w", side.name, err)
		}
	}
	return nil
}

// validateStyleEnums 一次性校验 Style 里所有枚举型字段。
// 供所有写入样式的入口调用，避免各路径漏检。
func validateStyleEnums(s Style) error {
	if err := validateFont(s.Font); err != nil {
		return err
	}
	if err := validateBorder(s.Border); err != nil {
		return err
	}
	return validateAlignment(s.Alignment)
}
