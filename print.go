package excelgo

// print.go 实现页面设置与打印属性（页边距、纸张/方向、缩放适配、打印区域、页眉页脚、
// 重复打印行）的读写，配套扩展 SheetProps 并提供 GetProps 读取闭环。
//
// 说明：
//   - pageMargins / pageSetup / headerFooter 属于 worksheet XML 的子元素；
//   - 打印区域（_xlnm.Print_Area）与重复打印行（_xlnm.Print_Titles）以 definedName
//     形式存在于 workbook.xml，localSheetId 为该表在 <sheets> 中的 0 基序号。

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// PageMargins 页边距（单位：英寸）。
type PageMargins struct {
	Left   float64
	Right  float64
	Top    float64
	Bottom float64
	Header float64
	Footer float64
}

// PageSetup 页面设置。
type PageSetup struct {
	PaperSize   int    // 纸张：9=A4，8=A3，1=Letter 等（0=不改变）
	Orientation string // portrait / landscape（空=不改变）
	Scale       int    // 打印缩放百分比（0=不改变）
	FitToWidth  int    // 适应页宽（0=不使用）
	FitToHeight int    // 适应页高（0=不使用）
}

// HeaderFooter 页眉页脚（支持奇偶页）。
type HeaderFooter struct {
	OddHeader  string
	OddFooter  string
	EvenHeader string
	EvenFooter string
}

// sheetLocalIndex 返回工作表在 <sheets> 中的 0 基序号（localSheetId）。
func (s *WorkSheet) sheetLocalIndex() int {
	wb := s.book.wbStruct()
	for i, sh := range wb.Sheets.Sheet {
		if sh.Name == s.Name() {
			return i
		}
	}
	return -1
}

// tabColor 读取该工作表在 worksheet XML 的 <sheetPr><tabColor> 中的标签颜色（ARGB 十六进制，无则空）。
func (s *WorkSheet) tabColor() string {
	ws := string(s.ws())
	if m := regexp.MustCompile(`(?s)<tabColor\b[^>]*\brgb="([^"]*)"`).FindStringSubmatch(ws); m != nil {
		return m[1]
	}
	return ""
}

// GetProps 读取工作表属性（含打印/页面设置）。
func (s *WorkSheet) GetProps() (SheetProps, error) {
	props := SheetProps{}
	// 网格线 / 缩放
	if m := regexp.MustCompile(`(?s)<sheetView\b[^>]*\bshowGridLines="([^"]*)"`).FindStringSubmatch(string(s.ws())); m != nil {
		v := m[1] == "1"
		props.ShowGridLines = &v
	}
	if m := regexp.MustCompile(`(?s)<sheetView\b[^>]*\bzoomScale="(\d+)"`).FindStringSubmatch(string(s.ws())); m != nil {
		if n, err := strconv.Atoi(m[1]); err == nil {
			props.Zoom = n
		}
	}
	// 标签色
	props.TabColor = s.tabColor()

	// 页边距
	if pm := readPageMargins(string(s.ws())); pm != nil {
		props.PageMargins = pm
	}
	// 页面设置
	if ps := readPageSetup(string(s.ws())); ps != nil {
		props.PageSetup = ps
	}
	// 页眉页脚
	if hf := readHeaderFooter(string(s.ws())); hf != nil {
		props.HeaderFooter = hf
	}
	// 打印区域 / 重复行
	idx := s.sheetLocalIndex()
	if idx >= 0 {
		wbXML := string(s.fm()["xl/workbook.xml"])
		if r := readDefinedName(wbXML, "_xlnm.Print_Area", idx); r != "" {
			props.PrintArea = stripSheetPrefix(r, s.Name())
		}
		if r := readDefinedName(wbXML, "_xlnm.Print_Titles", idx); r != "" {
			props.PrintTitleRows = stripSheetPrefixRows(r, s.Name())
		}
	}
	return props, nil
}

// GetProps 读取 filename 中 sheetRef 的工作表属性。
func GetProps(filename, sheetRef string) (SheetProps, error) {
	b, err := Open(filename)
	if err != nil {
		return SheetProps{}, err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return SheetProps{}, err
	}
	return ws.GetProps()
}

// ---------- worksheet 级写入 ----------

func setPageMarginsInWS(ws string, pm *PageMargins) string {
	block := ""
	if pm != nil {
		block = fmt.Sprintf(`<pageMargins left="%g" right="%g" top="%g" bottom="%g" header="%g" footer="%g"/>`,
			pm.Left, pm.Right, pm.Top, pm.Bottom, pm.Header, pm.Footer)
	}
	ws = regexp.MustCompile(`(?s)<pageMargins\b[^>]*/?>`).ReplaceAllString(ws, block)
	if block != "" && !strings.Contains(ws, "<pageMargins") {
		ws = insertBeforeWorksheetClose(ws, block)
	}
	return ws
}

func setPageSetupInWS(ws string, ps *PageSetup) string {
	block := ""
	if ps != nil {
		var b strings.Builder
		b.WriteString("<pageSetup")
		if ps.PaperSize != 0 {
			fmt.Fprintf(&b, ` paperSize="%d"`, ps.PaperSize)
		}
		if ps.Orientation != "" {
			fmt.Fprintf(&b, ` orientation="%s"`, safeAttr(ps.Orientation))
		}
		if ps.Scale != 0 {
			fmt.Fprintf(&b, ` scale="%d"`, ps.Scale)
		}
		if ps.FitToWidth != 0 || ps.FitToHeight != 0 {
			fmt.Fprintf(&b, ` fitToWidth="%d" fitToHeight="%d"`, ps.FitToWidth, ps.FitToHeight)
		}
		b.WriteString("/>")
		block = b.String()
		// 保留旧标签里本库**不管理**的属性 —— 尤其 r:id，它是指向
		// xl/printerSettings/*.bin 的关系。整体替换会把它抹掉，
		// 部件和关系都还在（不报错），但 worksheet 不再引用 → 打印机设置静默丢失。
		if old := regexp.MustCompile(`(?s)<pageSetup\b[^>]*?/?>`).FindString(ws); old != "" {
			block = keepForeignAttrs(old, block, pageSetupManagedAttrs)
		}
	}
	ws = regexp.MustCompile(`(?s)<pageSetup\b[^>]*/?>`).ReplaceAllString(ws, block)
	if block != "" && !strings.Contains(ws, "<pageSetup") {
		ws = insertBeforeWorksheetClose(ws, block)
	}
	return ws
}

func setHeaderFooterInWS(ws string, hf *HeaderFooter) string {
	block := ""
	if hf != nil {
		var b strings.Builder
		b.WriteString("<headerFooter>")
		if hf.OddHeader != "" {
			fmt.Fprintf(&b, "<oddHeader>%s</oddHeader>", safeText(hf.OddHeader))
		}
		if hf.OddFooter != "" {
			fmt.Fprintf(&b, "<oddFooter>%s</oddFooter>", safeText(hf.OddFooter))
		}
		if hf.EvenHeader != "" {
			fmt.Fprintf(&b, "<evenHeader>%s</evenHeader>", safeText(hf.EvenHeader))
		}
		if hf.EvenFooter != "" {
			fmt.Fprintf(&b, "<evenFooter>%s</evenFooter>", safeText(hf.EvenFooter))
		}
		b.WriteString("</headerFooter>")
		block = b.String()
	}
	ws = regexp.MustCompile(`(?s)<headerFooter\b.*?</headerFooter>`).ReplaceAllString(ws, block)
	if block != "" && !strings.Contains(ws, "<headerFooter") {
		ws = insertBeforeWorksheetClose(ws, block)
	}
	return ws
}

// insertBeforeWorksheetClose 在 </worksheet> 之前插入块（用于 pageMargins/pageSetup/headerFooter）。
// 若存在 drawing，则插在 drawing 之前以更接近 OOXML 顺序（printOptions→pageMargins→pageSetup→headerFooter→drawing）。
func insertBeforeWorksheetClose(ws, block string) string {
	if idx := strings.LastIndex(ws, "<drawing"); idx != -1 {
		return ws[:idx] + block + ws[idx:]
	}
	if idx := strings.LastIndex(ws, "</worksheet>"); idx != -1 {
		return ws[:idx] + block + ws[idx:]
	}
	return ws + block
}

// ---------- worksheet 级读取 ----------

func readPageMargins(ws string) *PageMargins {
	m := regexp.MustCompile(`(?s)<pageMargins\b[^>]*/>`).FindStringSubmatch(ws)
	if m == nil {
		return nil
	}
	tag := m[0]
	f := func(attr string) float64 {
		if v := attrOf(tag, attr); v != "" {
			if n, err := strconv.ParseFloat(v, 64); err == nil {
				return n
			}
		}
		return 0
	}
	pm := &PageMargins{
		Left:   f("left"),
		Right:  f("right"),
		Top:    f("top"),
		Bottom: f("bottom"),
		Header: f("header"),
		Footer: f("footer"),
	}
	return pm
}

func readPageSetup(ws string) *PageSetup {
	m := regexp.MustCompile(`(?s)<pageSetup\b[^>]*/>`).FindStringSubmatch(ws)
	if m == nil {
		return nil
	}
	tag := m[0]
	ps := &PageSetup{}
	if v := attrOf(tag, "paperSize"); v != "" {
		ps.PaperSize, _ = strconv.Atoi(v)
	}
	ps.Orientation = attrOf(tag, "orientation")
	if v := attrOf(tag, "scale"); v != "" {
		ps.Scale, _ = strconv.Atoi(v)
	}
	if v := attrOf(tag, "fitToWidth"); v != "" {
		ps.FitToWidth, _ = strconv.Atoi(v)
	}
	if v := attrOf(tag, "fitToHeight"); v != "" {
		ps.FitToHeight, _ = strconv.Atoi(v)
	}
	return ps
}

func readHeaderFooter(ws string) *HeaderFooter {
	m := regexp.MustCompile(`(?s)<headerFooter\b.*?</headerFooter>`).FindStringSubmatch(ws)
	if m == nil {
		return nil
	}
	inner := m[0]
	hf := &HeaderFooter{}
	get := func(tag string) string {
		if mm := regexp.MustCompile(`(?s)<` + tag + `>(.*?)</` + tag + `>`).FindStringSubmatch(inner); mm != nil {
			return unescapeXML(mm[1])
		}
		return ""
	}
	hf.OddHeader = get("oddHeader")
	hf.OddFooter = get("oddFooter")
	hf.EvenHeader = get("evenHeader")
	hf.EvenFooter = get("evenFooter")
	return hf
}

// ---------- workbook 级 definedName 辅助 ----------

// setOrReplaceDefinedName 在 workbook.xml 中设置或替换指定 name+localSheetId 的 definedName。
// name 做 XML 转义后再匹配与写入（name 可能来自调用方输入）。
func setOrReplaceDefinedName(wbXML []byte, name string, localSheetId int, value string) []byte {
	content := string(wbXML)
	re := regexp.MustCompile(`(?s)<definedName\s+name="` + regexp.QuoteMeta(safeText(name)) + `"\s+localSheetId="` + strconv.Itoa(localSheetId) + `"[^>]*>[^<]*</definedName>`)
	newTag := fmt.Sprintf(`<definedName name="%s" localSheetId="%d">%s</definedName>`, safeText(name), localSheetId, safeText(value))
	if m := re.FindString(content); m != "" {
		// 用精确匹配子串替换，避免 regexp 替换串把 $ 当作反向引用/expansion
		return []byte(strings.Replace(content, m, newTag, 1))
	}
	// 移除空占位 <definedNames />
	content = strings.ReplaceAll(content, "<definedNames />", "")
	tag := newTag
	if idx := strings.LastIndex(content, "</definedNames>"); idx != -1 {
		return []byte(content[:idx] + tag + content[idx:])
	}
	if idx := strings.LastIndex(content, "</sheets>"); idx != -1 {
		closeEnd := idx + len("</sheets>")
		block := `<definedNames>` + tag + `</definedNames>`
		return []byte(content[:closeEnd] + block + content[closeEnd:])
	}
	return []byte(content + `<definedNames>` + tag + `</definedNames>`)
}

// readDefinedName 读取指定 name+localSheetId 的 definedName 值（含表名前缀，如 "Sheet1!$A$1:$D$10"）。
func readDefinedName(wbXML, name string, localSheetId int) string {
	re := regexp.MustCompile(`(?s)<definedName\s+name="` + regexp.QuoteMeta(name) + `"\s+localSheetId="` + strconv.Itoa(localSheetId) + `"[^>]*>([^<]*)</definedName>`)
	if m := re.FindStringSubmatch(wbXML); m != nil {
		return unescapeXML(m[1])
	}
	return ""
}

// ---------- 范围规范化辅助 ----------

// absRange 把范围（如 "A1:D10"、"A1"）规整为带 $ 的绝对引用（"$A$1:$D$10"）。
// 已含 $ 的部分保持不变；空串原样返回。
//
// 若 ref 不是「纯」单元格/行列引用（含有括号、函数名、多段引用等），则原样返回，
// 避免把 "SUM(A1:B2)" 之类表达式错误改写成 "$SUM:$B$2"。
func absRange(ref string) string {
	ref = strings.TrimSpace(ref)
	if ref == "" {
		return ref
	}
	parts := strings.Split(ref, ":")
	// 只接受 1~2 段的单元格/行列引用，其余（表达式、多段联合引用、带表名前缀等）原样返回
	if len(parts) > 2 {
		return ref
	}
	for i, p := range parts {
		abs, ok := absCellRef(p)
		if !ok {
			// 不是合法的单元格/行列引用，整段放弃改写
			return ref
		}
		parts[i] = abs
	}
	return strings.Join(parts, ":")
}

// 合法引用：$?列字母$?数字（"A1"/"$A$1"）、$?数字（"1"/"$1"）、$?列字母（"A"/"$A"）
var reAbsCellRef = regexp.MustCompile(`^\$?(?:[A-Za-z]{1,3}\$?\d{1,7}|\d{1,7}|[A-Za-z]{1,3})$`)

// absCellRef 把单个单元格/行列引用规整为绝对引用：列与行前各补 $（若原已有则保留）。
// 支持纯行（"1"->"$1"）、纯列（"A"->"$A"）、行列混合（"A1"->"$A$1"、"$A1"->"$A$1"、
// "$A$1"->"$A$1"）。第二个返回值为 false 表示 c 不是合法引用，调用方应放弃改写。
func absCellRef(c string) (string, bool) {
	c = strings.TrimSpace(c)
	if c == "" {
		return c, false
	}
	if !reAbsCellRef.MatchString(c) {
		return c, false
	}
	// 解析可选 $ 前缀、列字母、可选 $、行数字
	idx := 0
	hasColDollar := false
	hasRowDollar := false
	if idx < len(c) && c[idx] == '$' {
		if idx+1 < len(c) && c[idx+1] >= 'A' && c[idx+1] <= 'Z' {
			hasColDollar = true
			idx++
		} else if idx+1 < len(c) && c[idx+1] >= '0' && c[idx+1] <= '9' {
			hasRowDollar = true
			idx++
		} else {
			return c, false // 孤立的 $
		}
	}
	colStart := idx
	for idx < len(c) && c[idx] >= 'A' && c[idx] <= 'Z' {
		idx++
	}
	colLetters := c[colStart:idx]
	if idx < len(c) && c[idx] == '$' {
		hasRowDollar = true
		idx++
	}
	rowStart := idx
	for idx < len(c) && c[idx] >= '0' && c[idx] <= '9' {
		idx++
	}
	rowDigits := c[rowStart:idx]
	if idx != len(c) {
		return c, false // 有多余尾随字符
	}

	var b strings.Builder
	if colLetters != "" {
		b.WriteString("$")
		b.WriteString(colLetters)
	} else if hasColDollar {
		b.WriteString("$")
	}
	if rowDigits != "" {
		b.WriteString("$")
		b.WriteString(rowDigits)
	} else if hasRowDollar && colLetters != "" {
		b.WriteString("$")
	}
	return b.String(), true
}

// absRangeRows 把重复打印行/列范围（如 "1:1"、"A:A"）规整为带 $ 的绝对引用。
func absRangeRows(ref string) string {
	return absRange(ref)
}

// stripSheetPrefix 去掉 "Sheet1!$A$1:$D$10" 中的表名前缀，返回 "A1:D10"。
func stripSheetPrefix(rng, sheetName string) string {
	prefix := sheetName + "!"
	if strings.HasPrefix(rng, prefix) {
		return rng[len(prefix):]
	}
	return rng
}

// stripSheetPrefixRows 去掉 "Sheet1!$1:$1" 中的表名前缀，返回 "1:1"。
func stripSheetPrefixRows(rng, sheetName string) string {
	prefix := sheetName + "!"
	if strings.HasPrefix(rng, prefix) {
		return rng[len(prefix):]
	}
	return rng
}

// ---------- 属性保留：避免改页面设置时丢掉外部引用 ----------

// keepForeignAttrs 从旧标签里挑出**本库不管理**的属性，原样拼到新标签上。
//
// 为什么必须做：<pageSetup> 上挂着 r:id="rIdN"，它是指向
// xl/printerSettings/printerSettingsN.bin 的关系。打印机设置二进制是
// Excel/WPS 保存工作簿时写下的打印机专属块（纸张、驱动、边距微调），
// 只存在于真实文件里，无法由本库重建。
//
// 早期实现用正则把整个 <pageSetup .../> 替换掉，于是改一次纸张或方向就把
// r:id 抹了：部件与关系都还在（不报错），但 worksheet 不再引用它 ——
// 打印机设置**静默丢失**，用户只会发现"打印出来不是我设的版式"。
//
// 与其逐个记忆哪些属性要保留，不如反向定义：本库只覆盖下列少数属性，
// 其余（r:id、firstPageNumber、copies、blackAndWhite…）一律原样搬运。
func keepForeignAttrs(oldTag, newTag string, managed []string) string {
	if oldTag == "" {
		return newTag
	}
	var extra strings.Builder
	for _, m := range regexp.MustCompile(`([A-Za-z:]+)\s*=\s*"([^"]*)"`).
		FindAllStringSubmatch(oldTag, -1) {
		attr, val := m[1], m[2]
		skip := false
		for _, k := range managed {
			if attr == k {
				skip = true
				break
			}
		}
		if skip {
			continue
		}
		if strings.Contains(newTag, " "+attr+`="`) {
			continue // 新标签已有该属性
		}
		extra.WriteString(" " + attr + `="` + safeAttr(val) + `"`)
	}
	if extra.Len() == 0 {
		return newTag
	}
	// 插到自闭合斜杠之前
	if strings.HasSuffix(newTag, "/>") {
		return strings.TrimSuffix(newTag, "/>") + extra.String() + "/>"
	}
	if strings.HasSuffix(newTag, ">") {
		return strings.TrimSuffix(newTag, ">") + extra.String() + ">"
	}
	return newTag + extra.String()
}

// pageSetupManagedAttrs 是 setPageSetupInWS 会覆盖的属性，其余一律保留。
var pageSetupManagedAttrs = []string{
	"paperSize", "orientation", "scale", "fitToWidth", "fitToHeight",
}
