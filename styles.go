package excelgo

// styles.go 提供跨工作簿合并工作表时所需的 xl/styles.xml 合并与去重能力。
//
// 背景：每个 .xlsx 都有独立的 styles.xml，单元格的 s 属性是一个索引，指向
// cellXfs（单元格格式）数组。把源表的 worksheet 原样搬进目标工作簿时，若不合并
// styles，源单元格的 s 索引会指向目标里完全不同的格式，导致整片格式错乱。
//
// 策略（合并去重，避免样式数量溢出）：
//   - 只处理源工作表「实际用到」的 cellXfs 索引（从 worksheet 的 s 属性收集），
//     而不是把源 styles 全部搬过去；
//   - 对每条用到的源 cellXfs，先在目标 cellXfs 中寻找「语义等价」的条目并复用其
//     索引（去重）；找不到才追加；
//   - cellXfs 引用的 fontId/fillId/borderId/numFmtId 同样按内容去重后再引用，
//     以防字体/填充/边框/数字格式无谓膨胀。
//
// 实现说明：styles.xml 含默认命名空间，Go 的 encoding/xml 在结构体字段未声明同一
// 命名空间时无法匹配带命名空间的元素，且 ,innerxml 在复杂场景下行为不稳。因此本
// 实现改用「正则块提取 + 字符串比对」的方式，直接操作各段落的原始 XML 文本，
// 既不依赖命名空间解析，也能精确保留每个 font/fill/border/numFmt 的原始内容。

import (
	"fmt"
	"regexp"
	"strings"
)

// 注：本文件不使用 encoding/xml，改用正则块操作以规避默认命名空间解析问题。

// ---------- 块提取 ----------

// extractBlock 提取 styles.xml 中名为 tag 的顶层块（<tag ...>...</tag>）的原始文本。
// 返回 "" 表示不存在该块。
func extractBlock(stylesXML string, tag xmlTagName) string {
	t := regexp.QuoteMeta(string(tag))
	openRe := regexp.MustCompile(`(?s)<` + t + `(\s[^>]*)?>`)
	loc := openRe.FindStringIndex(stylesXML)
	if loc == nil {
		return ""
	}
	openEnd := loc[1]
	closeRe := regexp.MustCompile(`(?s)</` + t + `>`)
	cls := closeRe.FindStringIndex(stylesXML[openEnd:])
	if cls == nil {
		return ""
	}
	closeEnd := openEnd + cls[1]
	return stylesXML[loc[0]:closeEnd]
}

// extractItems 提取某块内部所有名为 item 的子元素原始文本（按出现顺序）。
// 同时支持带闭合标签（<item>...</item>）与自闭合（<item .../>）两种形式。
func extractItems(block string, item xmlTagName) []string {
	if block == "" {
		return nil
	}
	// 匹配 <item ...> 或 <item .../>（含自闭合）
	it := regexp.QuoteMeta(string(item))
	openRe := regexp.MustCompile(`(?s)<` + it + `(\s[^>]*)?(/?>)`)
	var out []string
	pos := 0
	for {
		loc := openRe.FindStringIndex(block[pos:])
		if loc == nil {
			break
		}
		seg := block[pos+loc[0] : pos+loc[1]]
		isSelfClose := strings.HasSuffix(strings.TrimSpace(seg), "/>")
		absOpen := pos + loc[1]
		if isSelfClose {
			out = append(out, seg)
			pos = absOpen
			continue
		}
		// 非自闭合：找匹配闭合标签
		closeRe := regexp.MustCompile(`(?s)</` + it + `>`)
		cls := closeRe.FindStringIndex(block[absOpen:])
		if cls == nil {
			break
		}
		absClose := absOpen + cls[1]
		out = append(out, block[pos+loc[0]:absClose])
		pos = absClose
	}
	return out
}

// attrOf 从元素原始文本中提取指定属性值（字符串）。
func attrOf(elem, attr string) string {
	re := regexp.MustCompile(`\b` + attr + `="([^"]*)"`)
	m := re.FindStringSubmatch(elem)
	if m == nil {
		return ""
	}
	return m[1]
}

// xfKey 提取一条 <xf> 的语义关键字段，用于等价比较。
// 比较：numFmtId/fontId/fillId/borderId/xfId/apply*(布尔)/alignment 内部文本。
func xfKey(xf string) string {
	numFmt := attrOf(xf, "numFmtId")
	font := attrOf(xf, "fontId")
	fill := attrOf(xf, "fillId")
	border := attrOf(xf, "borderId")
	xfId := attrOf(xf, "xfId")
	applyFont := attrOf(xf, "applyFont")
	applyFill := attrOf(xf, "applyFill")
	applyBorder := attrOf(xf, "applyBorder")
	applyAlign := attrOf(xf, "applyAlignment")
	applyNum := attrOf(xf, "applyNumberFormat")
	// alignment 内部（若存在）：提取 <alignment .../> 原始片段
	alignRe := regexp.MustCompile(`(?s)<alignment\b.*?/>`)
	align := alignRe.FindString(xf)
	return strings.Join([]string{numFmt, font, fill, border, xfId, applyFont, applyFill, applyBorder, applyAlign, applyNum, align}, "|")
}

// eqElem 判断两个同类型元素（font/fill/border/numFmt）原始文本是否「内容相同」。
// 对 numFmt 还需比较 numFmtId（自定义格式按 formatCode 去重，但索引可能不同，
// 直接用完整文本比对更稳妥——不过 numFmt 合并走单独逻辑）。
func eqElem(a, b string) bool {
	na := normalizeWhitespace(a)
	nb := normalizeWhitespace(b)
	return na == nb
}

// normalizeWhitespace 归一化元素文本中的空白（属性间、标签间），用于稳健比较。
// 注意：仅折叠标签外与属性分隔的空白，不改变属性值。简单实现：折叠连续空白为单空格。
func normalizeWhitespace(s string) string {
	return wsRe.ReplaceAllString(s, " ")
}

var wsRe = regexp.MustCompile(`\s+`)

// ---------- 合并器（正则版） ----------

// StylesMerger 把源 styles 合并进目标 styles，并提供 s 索引重映射。
type StylesMerger struct {
	target string // 目标 styles.xml 原始文本（保留其余段落）
	// 下面缓存各块当前条目（用于去重比较与追加）
	tFonts   []string
	tFills   []string
	tBorders []string
	tNumFmts []string // 以 "id|formatCode" 形式缓存，便于按 formatCode 去重
	tCellXfs []string
	// tStyleXfs 是**命名样式**的底层 xf 表（与 cellXfs 并列的另一张表）。
	// 必须一并合并：<cellStyle xfId="N"> 指向的是这里，不是 cellXfs。
	tStyleXfs []string
	// styleXfRemap 记录「源 cellStyleXfs 下标 -> 目标下标」，供合并命名样式清单用。
	styleXfRemap map[int]int
}

// NewStylesMerger 构造合并器。
func NewStylesMerger(targetStyles, sourceStyles []byte) (*StylesMerger, error) {
	ts := string(targetStyles)
	m := &StylesMerger{
		target:   ts,
		tFonts:   extractItems(extractBlock(ts, tagName("fonts")), "font"),
		tFills:   extractItems(extractBlock(ts, tagName("fills")), "fill"),
		tBorders: extractItems(extractBlock(ts, tagName("borders")), "border"),
		tCellXfs: extractItems(extractBlock(ts, tagName("cellXfs")), "xf"),
	}
	// cellStyleXfs 至少要有一条默认 xf（OOXML 要求非空）
	m.tStyleXfs = extractItems(extractBlock(ts, tagName("cellStyleXfs")), "xf")
	if len(m.tStyleXfs) == 0 {
		m.tStyleXfs = []string{`<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>`}
	}
	// 缓存目标自定义 numFmt：以 formatCode 为键
	for _, nf := range extractItems(extractBlock(ts, tagName("numFmts")), "numFmt") {
		id := attrOf(nf, "numFmtId")
		fc := attrOf(nf, "formatCode")
		m.tNumFmts = append(m.tNumFmts, id+"|"+fc)
	}
	_ = sourceStyles
	return m, nil
}

// Merge 把源中 usedSrcXfs（相对于源 cellXfs 数组的索引）合并进目标，
// 返回「源 cellXfs 索引 -> 目标 cellXfs 索引」的映射。
// 需要 sourceStyles 原始字节以确定每条源 cellXfs 的内容及其引用的 font/fill/border/numFmt。
func (m *StylesMerger) Merge(sourceStyles []byte, usedSrcXfs []int) (map[int]int, error) {
	ss := string(sourceStyles)
	srcFonts := extractItems(extractBlock(ss, tagName("fonts")), "font")
	srcFills := extractItems(extractBlock(ss, tagName("fills")), "fill")
	srcBorders := extractItems(extractBlock(ss, tagName("borders")), "border")
	srcNumFmts := extractItems(extractBlock(ss, tagName("numFmts")), "numFmt")
	srcCellXfs := extractItems(extractBlock(ss, tagName("cellXfs")), "xf")

	// 自定义 numFmt 源索引 -> 目标 id（去重映射）
	srcNumFmtToTarget := make(map[string]string)
	for _, nf := range srcNumFmts {
		id := attrOf(nf, "numFmtId")
		fc := attrOf(nf, "formatCode")
		tid := m.ensureNumFmt(id, fc)
		srcNumFmtToTarget[id] = tid
	}

	remap := make(map[int]int)
	for _, sIdx := range usedSrcXfs {
		if sIdx < 0 || sIdx >= len(srcCellXfs) {
			remap[sIdx] = 0
			continue
		}
		srcXf := srcCellXfs[sIdx]
		fontId, err := m.ensureIndexed(srcFonts, m.tFonts, &m.tFonts, attrOf(srcXf, "fontId"))
		if err != nil {
			return nil, err
		}
		fillId, err := m.ensureIndexed(srcFills, m.tFills, &m.tFills, attrOf(srcXf, "fillId"))
		if err != nil {
			return nil, err
		}
		borderId, err := m.ensureIndexed(srcBorders, m.tBorders, &m.tBorders, attrOf(srcXf, "borderId"))
		if err != nil {
			return nil, err
		}
		numFmtId := attrOf(srcXf, "numFmtId")
		if tid, ok := srcNumFmtToTarget[numFmtId]; ok {
			numFmtId = tid
		}

		// 构造目标侧等价 xf 文本（按标准属性顺序，保留 alignment）
		newXf := buildXF(numFmtId, itoa(fontId), itoa(fillId), itoa(borderId), srcXf)

		// 在目标 cellXfs 中寻找等价条目
		if idx, ok := findEqualXF(m.tCellXfs, newXf); ok {
			remap[sIdx] = idx
			continue
		}
		idx := len(m.tCellXfs)
		m.tCellXfs = append(m.tCellXfs, newXf)
		remap[sIdx] = idx
	}
	if err := m.mergeStyleXfs(ss, srcFonts, srcFills, srcBorders, srcNumFmtToTarget); err != nil {
		return nil, err
	}
	return remap, nil
}

// mergeStyleXfs 把源的 cellStyleXfs 并入目标，并记录「源命名样式名 -> 目标 xfId」。
//
// 为什么要在这里做（而不是在外层单独处理）：命名样式的 xf 内部同样引用
// fontId/fillId/borderId/numFmtId，这些索引只有本类型才知道正确的目标值 —-
// 手工按"目标表长度"算偏移会与cellXfs 的合并重复累加，导致 fontId 越界。
func (m *StylesMerger) mergeStyleXfs(ss string,
	srcFonts, srcFills, srcBorders []string,
	srcNumFmtToTarget map[string]string) error {
	srcStyleXfs := extractItems(extractBlock(ss, tagName("cellStyleXfs")), "xf")
	if len(srcStyleXfs) <= 1 {
		return nil // 只有默认 xf，无需搬
	}
	// 源命名样式清单：名字 -> 源 cellStyleXfs 下标
	srcNamed := extractItems(extractBlock(ss, tagName("cellStyles")), "cellStyle")
	if len(srcNamed) == 0 {
		return nil
	}
	nameToTarget := map[string]int{}
	for i := 1; i < len(srcStyleXfs); i++ {
		srcXf := srcStyleXfs[i]
		fontId, err := m.ensureIndexed(srcFonts, m.tFonts, &m.tFonts, attrOf(srcXf, "fontId"))
		if err != nil {
			return err
		}
		fillId, err := m.ensureIndexed(srcFills, m.tFills, &m.tFills, attrOf(srcXf, "fillId"))
		if err != nil {
			return err
		}
		borderId, err := m.ensureIndexed(srcBorders, m.tBorders, &m.tBorders, attrOf(srcXf, "borderId"))
		if err != nil {
			return err
		}
		numFmtID := attrOf(srcXf, "numFmtId")
		if tid, ok := srcNumFmtToTarget[numFmtID]; ok {
			numFmtID = tid
		}
		newXf := buildXF(numFmtID, itoa(fontId), itoa(fillId), itoa(borderId), srcXf)
		if idx, ok := findEqualXF(m.tStyleXfs, newXf); ok {
			m.recordNamedStyle(nameToTarget, i, idx)
			continue
		}
		idx := len(m.tStyleXfs)
		m.tStyleXfs = append(m.tStyleXfs, newXf)
		m.recordNamedStyle(nameToTarget, i, idx)
	}
	_ = srcNamed
	return nil
}

// recordNamedStyle 暂存「源 cellStyleXfs 下标 -> 目标下标」的对应关系。
func (m *StylesMerger) recordNamedStyle(_ map[string]int, srcIdx, dstIdx int) {
	if m.styleXfRemap == nil {
		m.styleXfRemap = map[int]int{}
	}
	m.styleXfRemap[srcIdx] = dstIdx
}

// ensureIndexed 确保 srcItems[idx] 对应的元素已存在于目标 items（去重），
// 返回目标侧索引。targetPtr 用于追加时增长目标切片。
func (m *StylesMerger) ensureIndexed(srcItems, _ []string, targetPtr *[]string, srcIdxStr string) (int, error) {
	idx, err := atoiSafe(srcIdxStr)
	if err != nil || idx < 0 || idx >= len(srcItems) {
		return 0, nil // 越界则用默认（索引 0）
	}
	elem := srcItems[idx]
	for i, t := range *targetPtr {
		if eqElem(t, elem) {
			return i, nil
		}
	}
	*targetPtr = append(*targetPtr, elem)
	return len(*targetPtr) - 1, nil
}

// ensureNumFmt 确保自定义 numFmt（id, formatCode）已存在于目标，返回目标侧 id。
// 内置格式（id < 164）直接返回原 id。自定义格式按 formatCode 去重在目标 numFmts 中
// 追加（沿用源 id，保持与目标既有 numFmts 一致；若目标已存在同 formatCode 则复用其 id）。
func (m *StylesMerger) ensureNumFmt(id, formatCode string) string {
	n, err := atoiSafe(id)
	if err == nil && n < 164 {
		return id
	}
	// 按 formatCode 去重
	for _, entry := range m.tNumFmts {
		parts := strings.SplitN(entry, "|", 2)
		if len(parts) == 2 && parts[1] == formatCode {
			return parts[0]
		}
	}
	// 追加（沿用源 id；若存在同 id 则跳过重复添加）
	for _, entry := range m.tNumFmts {
		if strings.HasPrefix(entry, id+"|") {
			return id
		}
	}
	m.tNumFmts = append(m.tNumFmts, id+"|"+formatCode)
	return id
}

// findEqualXF 在目标 cellXfs 列表中查找与目标侧 newXf 语义等价的条目索引。
func findEqualXF(target []string, newXf string) (int, bool) {
	newKey := xfKey(newXf)
	for i, t := range target {
		if xfKey(t) == newKey {
			return i, true
		}
	}
	return 0, false
}

// buildXF 根据反查后的索引与源 xf 原文，构造目标侧等价的 <xf> 元素文本。
// 保留源 xf 的 apply* 属性与 alignment 子元素，仅替换 numFmtId/fontId/fillId/borderId。
func buildXF(numFmtId, fontId, fillId, borderId string, srcXf string) string {
	// 提取原 xf 的 apply* 与 alignment
	applyFont := attrOf(srcXf, "applyFont")
	applyFill := attrOf(srcXf, "applyFill")
	applyBorder := attrOf(srcXf, "applyBorder")
	applyAlign := attrOf(srcXf, "applyAlignment")
	applyNum := attrOf(srcXf, "applyNumberFormat")
	xfId := attrOf(srcXf, "xfId")
	alignRe := regexp.MustCompile(`(?s)<alignment\b.*?/>`)
	align := alignRe.FindString(srcXf)

	// xfId 是 Integer 属性：**缺失时必须整个省略**。
	// 写成 xfId="" 会让消费方转换失败 —— openpyxl 的 CellStyle.xfId 声明为 Int，
	// 遇到空串直接抛 "TypeError: expected <class 'int'>"，整份文件读不出来。
	xfIdAttr := ""
	if xfId != "" {
		xfIdAttr = ` xfId="` + xfId + `"`
	}
	s := `<xf numFmtId="` + numFmtId + `" fontId="` + fontId + `" fillId="` + fillId +
		`" borderId="` + borderId + `"` + xfIdAttr
	if applyFont == "1" {
		s += ` applyFont="1"`
	}
	if applyFill == "1" {
		s += ` applyFill="1"`
	}
	if applyBorder == "1" {
		s += ` applyBorder="1"`
	}
	if applyAlign == "1" {
		s += ` applyAlignment="1"`
	}
	if applyNum == "1" {
		s += ` applyNumberFormat="1"`
	}
	if align != "" {
		s += `>` + align + `</xf>`
	} else {
		s += `/>`
	}
	return s
}

// Build 生成合并后的目标 styles.xml 字节（补回默认命名空间与 XML 头）。
func (m *StylesMerger) Build() ([]byte, error) {
	out := m.target
	// 确保根元素带默认命名空间
	out = ensureStyleSheetNS(out)
	if !strings.HasPrefix(out, `<?xml`) {
		out = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + out
	}
	// 替换五个块
	out = replaceBlockText(out, tagName("numFmts"), func() string {
		s := `<numFmts count="` + itoa(len(m.tNumFmts)) + `">`
		for _, entry := range m.tNumFmts {
			parts := strings.SplitN(entry, "|", 2)
			if len(parts) != 2 {
				continue
			}
			s += `<numFmt numFmtId="` + parts[0] + `" formatCode="` + safeAttr(parts[1]) + `"/>`
		}
		s += `</numFmts>`
		return s
	})
	out = replaceBlockText(out, tagName("fonts"), func() string {
		s := `<fonts count="` + itoa(len(m.tFonts)) + `">`
		for _, f := range m.tFonts {
			s += f
		}
		s += `</fonts>`
		return s
	})
	out = replaceBlockText(out, tagName("fills"), func() string {
		s := `<fills count="` + itoa(len(m.tFills)) + `">`
		for _, f := range m.tFills {
			s += f
		}
		s += `</fills>`
		return s
	})
	out = replaceBlockText(out, tagName("borders"), func() string {
		s := `<borders count="` + itoa(len(m.tBorders)) + `">`
		for _, b := range m.tBorders {
			s += b
		}
		s += `</borders>`
		return s
	})
	out = replaceBlockText(out, tagName("cellStyleXfs"), func() string {
		s := `<cellStyleXfs count="` + itoa(len(m.tStyleXfs)) + `">`
		for _, x := range m.tStyleXfs {
			s += x
		}
		s += `</cellStyleXfs>`
		return s
	})
	out = replaceBlockText(out, tagName("cellXfs"), func() string {
		s := `<cellXfs count="` + itoa(len(m.tCellXfs)) + `">`
		for _, x := range m.tCellXfs {
			s += x
		}
		s += `</cellXfs>`
		return s
	})
	return []byte(out), nil
}

// ensureStyleSheetNS 确保 <styleSheet> 根元素声明默认命名空间。
func ensureStyleSheetNS(s string) string {
	if strings.Contains(s, `xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"`) {
		return s
	}
	rootStart := strings.Index(s, "<styleSheet")
	if rootStart == -1 {
		return s
	}
	rootEnd := strings.Index(s[rootStart:], ">")
	if rootEnd == -1 {
		return s
	}
	rootEnd += rootStart
	root := s[rootStart : rootEnd+1]
	root = strings.Replace(root, "<styleSheet", `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"`, 1)
	return s[:rootStart] + root + s[rootEnd+1:]
}

// replaceBlockText 用 fn 生成的新块替换 out 中名为 tag 的顶层块（保留其余段落）。
func replaceBlockText(out string, tag xmlTagName, fn func() string) string {
	t := regexp.QuoteMeta(string(tag))
	openRe := regexp.MustCompile(`(?s)<` + t + `(\s[^>]*)?>`)
	loc := openRe.FindStringIndex(out)
	if loc == nil {
		return out
	}
	openEnd := loc[1]
	closeRe := regexp.MustCompile(`(?s)</` + t + `>`)
	cls := closeRe.FindStringIndex(out[openEnd:])
	if cls == nil {
		return out
	}
	closeEnd := openEnd + cls[1]
	return out[:loc[0]] + fn() + out[closeEnd:]
}

func atoiSafe(s string) (int, error) {
	if s == "" {
		return 0, fmt.Errorf("empty")
	}
	var n int
	_, err := fmt.Sscanf(s, "%d", &n)
	return n, err
}

func itoa(n int) string {
	return fmt.Sprintf("%d", n)
}

// StyleXfRemap 返回「源 cellStyleXfs 下标 -> 目标下标」的映射。
//
// 供合并命名样式清单（<cellStyle xfId>）时重映射用。必须在 Merge 之后调用。
func (m *StylesMerger) StyleXfRemap() map[int]int {
	out := make(map[int]int, len(m.styleXfRemap))
	for k, v := range m.styleXfRemap {
		out[k] = v
	}
	return out
}
