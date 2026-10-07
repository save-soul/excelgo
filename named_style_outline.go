package excelgo

// 命名样式与大纲折叠。
//
// 两者的共同点：都是"元信息"而非数据 —— 命名样式写在 styles.xml 的
// <cellStyles>/<cellStyleXfs>，折叠状态写在 sheet XML 的行/列属性上。
// 都不改变单元格实际内容，故可独立实现。

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// ---------- 命名样式 ----------

// NamedStyleRef 描述一个命名样式（cellStyle + 其底层 xf）。
type NamedStyleRef struct {
	Name  string
	Index int // 在 cellStyleXfs 中的索引
}

// NewNamedStyle 注册一个命名样式：把 st 应用为一个可复用模板，返回样式名。
//
// 命名样式的作用：多个区域套用同一套格式时，改一处即全改。Excel 的
// "单元格样式"、CSS 的 class 都是这个机制。
//
// 返回的 name 供 SetNamedStyle 使用。
func (s *WorkSheet) NewNamedStyle(name string, st Style) (string, error) {
	if err := validateStyleEnums(st); err != nil {
		return "", err
	}
	if name == "" {
		return "", fmt.Errorf("命名样式名不能为空")
	}
	fm := s.fm()
	styles, ok := fm["xl/styles.xml"]
	if !ok {
		fm["xl/styles.xml"] = []byte(defaultStylesXML())
		fm["[Content_Types].xml"] = insertOverrideInContentTypes(fm["[Content_Types].xml"],
			"/xl/styles.xml",
			"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
		styles = fm["xl/styles.xml"]
	}
	out, _, err := applyNamedStyle(string(styles), name, st)
	if err != nil {
		return "", err
	}
	fm["xl/styles.xml"] = []byte(out)
	s.setWS(string(s.ws())) // 触发 fm 落盘路径一致
	return name, nil
}

// NewNamedStyle 注册命名样式（包级 API）。
func NewNamedStyle(filename, sheetRef, name string, st Style) (string, error) {
	var got string
	err := withSheet(filename, sheetRef, func(s *WorkSheet) error {
		var e error
		got, e = s.NewNamedStyle(name, st)
		return e
	})
	return got, err
}

// SetNamedStyle 把命名样式应用到 ref 区域（等价于套用该模板的全部格式）。
func (s *WorkSheet) SetNamedStyle(ref, name string) error {
	r := newRangeRef(ref)
	if r == "" {
		return fmt.Errorf("无效的样式区域引用: %q", ref)
	}
	fm := s.fm()
	styles, ok := fm["xl/styles.xml"]
	if !ok {
		return fmt.Errorf("工作簿不含 styles.xml，无法应用命名样式")
	}
	xfIdx, err := lookupNamedStyleXF(string(styles), name)
	if err != nil {
		return err
	}
	// 逐格套用样式索引：直接改 s 属性，最小侵入
	ws := string(s.ws())
	cells, err := rectCells(string(r))
	if err != nil {
		return err
	}
	for _, ref := range cells {
		ws = setCellStyleAttrInWS(ws, ref, xfIdx)
	}
	s.setWS(ws)
	return nil
}

// SetNamedStyle 把命名样式应用到 ref 区域（包级 API）。
func SetNamedStyle(filename, sheetRef, ref, name string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.SetNamedStyle(ref, name)
	})
}

// GetNamedStyles 列出工作簿中所有命名样式。
func (s *WorkSheet) GetNamedStyles() ([]NamedStyleRef, error) {
	styles, ok := s.fm()["xl/styles.xml"]
	if !ok {
		return nil, nil
	}
	return listNamedStyles(string(styles))
}

// GetNamedStyles 列出工作簿中所有命名样式（包级 API）。
func GetNamedStyles(filename, sheetRef string) ([]NamedStyleRef, error) {
	b, err := Open(filename)
	if err != nil {
		return nil, err
	}
	s, err := b.Sheet(sheetRef)
	if err != nil {
		return nil, err
	}
	return s.GetNamedStyles()
}

// applyNamedStyle 在 styles.xml 中注册命名样式。
//
// OOXML 的两处结构：
//   - <cellStyleXfs>  —— 命名样式指向的"底层 xf"（不含 applyXxx 标志）
//   - <cellStyles>    —— 命名样式清单，name 属性即用户看到的样式名
//
// 同时要让 <dxfs> 之前保持顺序 —— cellStyles 必须在 dxfs 之前（OOXML 要求）。
func applyNamedStyle(stylesXML, name string, st Style) (string, int, error) {
	// 1) 生成底层 xf 并注册进 cellXfs（复用既有逻辑拿到 s 索引）
	nsBytes, cellXfIdx, err := applyStyleToStylesXML([]byte(stylesXML), st)
	if err != nil {
		return "", 0, err
	}
	// 2) 把该 xf 追加到 cellStyleXfs，并记入 cellStyles
	cellStyleXfsRe := regexp.MustCompile(`(?s)(<cellStyleXfs\b[^>]*>)(.*?)(</cellStyleXfs>)`)
	newStyles := string(nsBytes)
	m := cellStyleXfsRe.FindStringSubmatch(newStyles)
	if m == nil {
		// 没有 cellStyleXfs 就建一个（OOXML 要求该元素存在）
		xf := buildXFForStyle(st, styleIndex("0"), styleIndex("0"),
			styleIndex("0"), styleIndex("0"))
		block := `<cellStyleXfs count="1">` + xf + `</cellStyleXfs>`
		out := insertBeforeElement(newStyles, []string{"<dxfs", "<tableStyles", "</styleSheet>"}, block)
		return registerCellStyle(out, name, 0, cellXfIdx)
	}
	inner := m[2]
	count := strings.Count(inner, "<xf")
	styleXfIdx := count // 新增项的索引 = 现有数量
	// 底层 xf：与 cellXfs 中同位置的 xf 保持一致的 numFmtId/fontId 等
	xf := extractXFAt(newStyles, cellXfIdx)
	inner += xf
	newStyles = cellStyleXfsRe.ReplaceAllString(newStyles,
		`<cellStyleXfs count="`+strconv.Itoa(count+1)+`">`+inner+`</cellStyleXfs>`)
	return registerCellStyle(newStyles, name, styleXfIdx, cellXfIdx)
}

// registerCellStyle 把命名样式写入 <cellStyles> 清单。
func registerCellStyle(stylesXML, name string, styleXfIdx, cellXfIdx int) (string, int, error) {
	entry := `<cellStyle name="` + safeAttr(name) +
		`" xfId="` + strconv.Itoa(styleXfIdx) + `"/>`
	csRe := regexp.MustCompile(`(?s)(<cellStyles\b[^>]*>)(.*?)(</cellStyles>)`)
	if m := csRe.FindStringSubmatch(stylesXML); m != nil {
		inner := m[2]
		count := strings.Count(inner, "<cellStyle ")
		inner += entry
		out := csRe.ReplaceAllString(stylesXML,
			`<cellStyles count="`+strconv.Itoa(count+1)+`">`+inner+`</cellStyles>`)
		return out, cellXfIdx, nil
	}
	// 没有 cellStyles 就新建，必须插在 <dxfs> 之前（OOXML 的元素顺序要求）
	block := `<cellStyles count="1">` + entry + `</cellStyles>`
	out := insertBeforeElement(stylesXML, []string{"<dxfs", "<tableStyles", "</styleSheet>"}, block)
	return out, cellXfIdx, nil
}

// lookupNamedStyleXF 按名字查出命名样式对应的 cellXfs 索引。
func lookupNamedStyleXF(stylesXML, name string) (int, error) {
	entryRe := regexp.MustCompile(`<cellStyle\b[^>]*\bname="([^"]*)"[^>]*\bxfId="(\d+)"`)
	all := entryRe.FindAllStringSubmatch(stylesXML, -1)
	for _, e := range all {
		if e[1] == name {
			xfId, _ := strconv.Atoi(e[2])
			// xfId 指向 cellStyleXfs，需再映射到 cellXfs 的 s 索引：
			// 本库 NewNamedStyle 两者一一对应（同位置追加），故直接沿用。
			if idx, ok := styleXfToCellXf(stylesXML, xfId); ok {
				return idx, nil
			}
			return xfId, nil
		}
	}
	return 0, fmt.Errorf("找不到命名样式 %q（可用 GetNamedStyles 列出已有的）", name)
}

// styleXfToCellXf 求 cellStyleXfs 第 i 个 xf 与 cellXfs 中同位置 xf 的对应关系。
// 本库的 NewNamedStyle 在两处同步追加，故按位置对应。
func styleXfToCellXf(stylesXML string, i int) (int, bool) {
	re := regexp.MustCompile(`(?s)<cellStyleXfs\b[^>]*>(.*?)</cellStyleXfs>`)
	m := re.FindStringSubmatch(stylesXML)
	if m == nil {
		return 0, false
	}
	n := strings.Count(m[1], "<xf")
	if i < 0 || i >= n {
		return 0, false
	}
	return i, true
}

// listNamedStyles 解析 <cellStyles> 清单。
func listNamedStyles(stylesXML string) ([]NamedStyleRef, error) {
	entryRe := regexp.MustCompile(`<cellStyle\b[^>]*\bname="([^"]*)"[^>]*\bxfId="(\d+)"`)
	var out []NamedStyleRef
	for _, e := range entryRe.FindAllStringSubmatch(stylesXML, -1) {
		idx, _ := strconv.Atoi(e[2])
		out = append(out, NamedStyleRef{Name: e[1], Index: idx})
	}
	return out, nil
}

// extractXFAt 取出 cellXfs 中第 i 个 <xf> 元素。
func extractXFAt(stylesXML string, i int) string {
	m := regexp.MustCompile(`(?s)(<cellXfs\b[^>]*>)(.*?)(</cellXfs>)`).
		FindStringSubmatch(stylesXML)
	if m == nil {
		return `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`
	}
	xfRe := regexp.MustCompile(`(?s)<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>`)
	all := xfRe.FindAllString(m[2], -1)
	if i < 0 || i >= len(all) {
		return `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`
	}
	return all[i]
}

// insertBeforeElement 在第一个命中的锚点前插入 block；都没命中则追加到末尾。
func insertBeforeElement(s string, anchors []string, block string) string {
	for _, a := range anchors {
		if i := strings.Index(s, a); i != -1 {
			return s[:i] + block + s[i:]
		}
	}
	return s + block
}

// ---------- 大纲折叠 ----------

// CollapseRows 折叠 [r1, r2] 行分组（隐藏其内容）。
//
// OOXML 里"折叠"由两部分表达：
//   - 隐藏行本身：<row r="N" hidden="1">
//   - 在分组起点的**上一行**加 <outlineLevel> 与 collapsed="1"（Excel 的折叠按钮）
//
// 只做前者会让数据消失但看不到折叠控件；只做后者则 Excel 认为该组已折叠
// 但数据仍显示。两者都要设。
func (s *WorkSheet) CollapseRows(r1, r2 int) error {
	if r1 < 1 || r2 < r1 {
		return fmt.Errorf("折叠行区间非法: %d-%d", r1, r2)
	}
	ws := string(s.ws())
	// 1) 隐藏区间内的行
	for r := r1; r <= r2; r++ {
		ws = modifyRowTag(ws, r, func(head string) string {
			head = regexp.MustCompile(`\s+hidden="[^"]*"`).ReplaceAllString(head, "")
			return strings.TrimRight(head, " ") + ` hidden="1"`
		})
	}
	// 2) 上一行标记为折叠态（该行本身不隐藏）
	if r1 > 1 {
		prev := r1 - 1
		ws = modifyRowTag(ws, prev, func(head string) string {
			head = regexp.MustCompile(`\s+collapsed="[^"]*"`).ReplaceAllString(head, "")
			return strings.TrimRight(head, " ") + ` collapsed="1"`
		})
	}
	s.setWS(ws)
	return nil
}

// ExpandRows 展开 [r1, r2] 行分组（取消隐藏并清除折叠标记）。
func (s *WorkSheet) ExpandRows(r1, r2 int) error {
	if r1 < 1 || r2 < r1 {
		return fmt.Errorf("展开行区间非法: %d-%d", r1, r2)
	}
	ws := string(s.ws())
	for r := r1; r <= r2; r++ {
		ws = modifyRowTag(ws, r, func(head string) string {
			return regexp.MustCompile(`\s+hidden="[^"]*"`).ReplaceAllString(head, "")
		})
	}
	if r1 > 1 {
		prev := r1 - 1
		ws = modifyRowTag(ws, prev, func(head string) string {
			return regexp.MustCompile(`\s+collapsed="[^"]*"`).ReplaceAllString(head, "")
		})
	}
	s.setWS(ws)
	return nil
}

// CollapseRows 折叠 [r1, r2] 行分组（包级 API）。
func CollapseRows(filename, sheetRef string, r1, r2 int) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.CollapseRows(r1, r2)
	})
}

// ExpandRows 展开 [r1, r2] 行分组（包级 API）。
func ExpandRows(filename, sheetRef string, r1, r2 int) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.ExpandRows(r1, r2)
	})
}

// SetOutlineSummary 设定摘要行的位置：true 表示摘要在下方（默认），false 表示在上方。
//
// summaryBelow / summaryRight 属于 <sheetPr> 里的 <outlinePr> 子元素，
// **不是 sheetPr 的属性** —— 写错位置会让消费方解析失败
// （openpyxl 报 WorksheetProperties.__init__() got an unexpected keyword）。
func (s *WorkSheet) SetOutlineSummary(below bool) error {
	ws := string(s.ws())
	v := "0"
	if below {
		v = "1"
	}
	attr := ` summaryBelow="` + v + `"`
	// 已有 outlinePr：改写其 summaryBelow
	if re := regexp.MustCompile(`<outlinePr[^>]*?/?>`); re.MatchString(ws) {
		m := re.FindString(ws)
		newTag := m
		if strings.Contains(m, ` summaryBelow="`) {
			newTag = regexp.MustCompile(` summaryBelow="[^"]*"`).
				ReplaceAllString(m, attr)
		} else {
			// 自闭合 -> 需展开成带子元素形态；带内容 -> 直接加属性
			if strings.HasSuffix(m, "/>") {
				newTag = strings.TrimSuffix(m, "/>") + attr + `/>`
			} else {
				newTag = strings.TrimSuffix(m, ">") + attr + `>`
			}
		}
		return s.replaceWSIn(ws, m, newTag)
	}
	// 没有 outlinePr：在 sheetPr 内补一个（sheetPr 必须是 worksheet 首个子元素）
	tag := `<outlinePr` + attr + `/>`
	if strings.Contains(ws, "<sheetPr") {
		// sheetPr 是自闭合的：展开后放入 outlinePr
		re := regexp.MustCompile(`<sheetPr[^>]*?/>`)
		if m := re.FindString(ws); m != "" {
			head := strings.TrimSuffix(m, "/>")
			newTag := head + `>` + tag + `</sheetPr>`
			return s.replaceWSIn(ws, m, newTag)
		}
		// sheetPr 带内容：插在它的开标签之后
		re2 := regexp.MustCompile(`<sheetPr[^>]*?>`)
		if m := re2.FindString(ws); m != "" {
			return s.replaceWSIn(ws, m, m+tag)
		}
	}
	// worksheet 根元素后紧跟 sheetPr（OOXML 要求它是第一个子元素）
	if i := strings.Index(ws, "<worksheet"); i != -1 {
		if j := strings.Index(ws[i:], ">"); j != -1 {
			at := i + j + 1
			block := `<sheetPr>` + tag + `</sheetPr>`
			s.setWS(ws[:at] + block + ws[at:])
			return nil
		}
	}
	s.setWS(ws)
	return nil
}

// replaceWSIn 把 sheet XML 中 first 替换为 second（仅替换一次），并写回。
func (s *WorkSheet) replaceWSIn(ws, first, second string) error {
	s.setWS(strings.Replace(ws, first, second, 1))
	return nil
}

// SetOutlineSummary 设定摘要行位置（包级 API）。
func SetOutlineSummary(filename, sheetRef string, below bool) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.SetOutlineSummary(below)
	})
}

// IsRowHidden 读回某行是否被隐藏。
func (s *WorkSheet) IsRowHidden(row int) bool {
	ws := string(s.ws())
	re := regexp.MustCompile(`(?s)<row\b[^>]*?\br="` + strconv.Itoa(row) + `"[^>]*?(?:/>|>)`)
	m := re.FindString(ws)
	return m != "" && strings.Contains(m, ` hidden="1"`) ||
		strings.Contains(m, ` hidden="true"`)
}

// IsRowHidden 读回某行是否隐藏（包级 API）。
func IsRowHidden(filename, sheetRef string, row int) (bool, error) {
	b, err := Open(filename)
	if err != nil {
		return false, err
	}
	s, err := b.Sheet(sheetRef)
	if err != nil {
		return false, err
	}
	return s.IsRowHidden(row), nil
}

// rectCells 枚举区域内的所有单元格坐标（0 基 rect）。
func rectCells(ref string) ([]string, error) {
	rc, err := newRangeRect(ref)
	if err != nil {
		return nil, err
	}
	var out []string
	for r := rc.top; r < rc.top+rc.h; r++ {
		for c := rc.col; c < rc.col+rc.w; c++ {
			out = append(out, refOf(r, c))
		}
	}
	return out, nil
}

// setCellStyleAttrInWS 给指定坐标的 <c> 设置 s 属性（不存在则补空壳）。
func setCellStyleAttrInWS(ws, ref string, styleIdx int) string {
	re := regexp.MustCompile(`<c r="` + regexp.QuoteMeta(ref) + `"([^>]*?)(?:/>|>)`)
	if m := re.FindStringSubmatch(ws); m != nil {
		attrs := regexp.MustCompile(`\s+s="[^"]*"`).ReplaceAllString(m[1], "")
		openTag := `<c r="` + ref + `" s="` + strconv.Itoa(styleIdx) + `"` + attrs + `/>`
		if !strings.HasSuffix(m[0], "/>") {
			// 原标签带内容：需保留到 </c> 的部分
			openTag = strings.TrimSuffix(openTag, "/>") + ">"
		}
		return ws[:re.FindStringIndex(ws)[0]] + openTag + ws[re.FindStringIndex(ws)[1]:]
	}
	// 单元格不存在：建空壳并按列序插入
	col, row, ok := refCoords(ref)
	if !ok {
		return ws
	}
	return setCellContents(ws, []cellOp{{row: row, col: col,
		xml: `<c r="` + ref + `" s="` + strconv.Itoa(styleIdx) + `"/>`}})
}
