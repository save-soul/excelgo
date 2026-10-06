package excelgo

// sheet_extra.go 提供除上述基础读写/行列/样式/区域之外的常用工作表功能：
// 合并单元格、列宽行高与隐藏、冻结窗格、超链接、自动筛选、数据验证、条件格式、
// 表格(ListObject)、批注、工作表属性(标签色/网格线/缩放)、工作表保护、行列分组大纲、
// 查找替换、文档属性。
//
// 全部沿用「内存 fileMap 字符串/正则精确定位」策略，仅新增/改写必要部件，不破坏其他
// 样式/布局/数据/图片。每个方法都在内存 fileMap 上操作，由 Book.Save 统一落盘。

import (
	"fmt"
	"path"
	"regexp"
	"strconv"
	"strings"
)

// ---------- 合并单元格 ----------

// MergeCells 把 ref 区域（如 "A1:C3"）合并为一块合并单元格。
func (s *WorkSheet) MergeCells(ref string) error {
	s.setWS(mergeCellsInWS(string(s.ws()), ref))
	return nil
}

// UnmergeCells 取消 ref 区域内的合并（移除匹配的 <mergeCell>）。
func (s *WorkSheet) UnmergeCells(ref string) error {
	s.setWS(unmergeCellsInWS(string(s.ws()), ref))
	return nil
}

func mergeCellsInWS(ws, ref string) string {
	ref = strings.TrimSpace(ref)
	re := regexp.MustCompile(`(?s)<mergeCells\b[^>]*>(.*?)</mergeCells>`)
	if m := re.FindStringSubmatch(ws); m != nil {
		inner := m[1]
		if strings.Contains(inner, `ref="`+regexp.QuoteMeta(ref)+`"`) {
			return ws // 已存在
		}
		body := `<mergeCells count="%d">` + inner + `<mergeCell ref="` + ref + `"/>` + `</mergeCells>`
		count := strings.Count(inner, "<mergeCell")
		body = fmt.Sprintf(body, count+1)
		return re.ReplaceAllString(ws, body)
	}
	block := `<mergeCells count="1"><mergeCell ref="` + ref + `"/>` + `</mergeCells>`
	return insertAfterSheetData(ws, frag(block))
}

func unmergeCellsInWS(ws, ref string) string {
	ref = strings.TrimSpace(ref)
	re := regexp.MustCompile(`(?s)<mergeCells\b[^>]*>(.*?)</mergeCells>`)
	m := re.FindStringSubmatch(ws)
	if m == nil {
		return ws
	}
	inner := m[1]
	newInner := regexp.MustCompile(`<mergeCell\b[^>]*\bref="`+regexp.QuoteMeta(ref)+`"[^>]*/?>`).ReplaceAllString(inner, "")
	count := strings.Count(newInner, "<mergeCell")
	if count == 0 {
		return re.ReplaceAllString(ws, "")
	}
	body := fmt.Sprintf(`<mergeCells count="%d">%s</mergeCells>`, count, newInner)
	return re.ReplaceAllString(ws, body)
}

// ---------- 列宽 / 行高 / 隐藏 ----------

// SetColWidth 设置单列（1 基列号）宽度（字符宽）。
func (s *WorkSheet) SetColWidth(col int, width float64) error {
	s.setWS(setColWidthInWS(string(s.ws()), col, col, width))
	return nil
}

// SetColWidthRange 设置 [min,max] 连续列的宽度。
func (s *WorkSheet) SetColWidthRange(min, max int, width float64) error {
	s.setWS(setColWidthInWS(string(s.ws()), min, max, width))
	return nil
}

// GetColWidth 读取覆盖 col 列的自定义宽度（未设置返回 0）。
func (s *WorkSheet) GetColWidth(col int) float64 {
	return getColWidthInWS(string(s.ws()), col)
}

// SetRowHeight 设置行高（点）。
func (s *WorkSheet) SetRowHeight(row int, height float64) error {
	s.setWS(setRowHeightInWS(string(s.ws()), row, height, true))
	return nil
}

// SetRowVisible 设置行是否可见（hidden=false 隐藏）。
func (s *WorkSheet) SetRowVisible(row int, visible bool) error {
	s.setWS(setRowHiddenInWS(string(s.ws()), row, !visible))
	return nil
}

// SetColVisible 设置列是否可见（hidden=false 隐藏）。
func (s *WorkSheet) SetColVisible(col int, visible bool) error {
	s.setWS(setColHiddenInWS(string(s.ws()), col, !visible))
	return nil
}

// setColWidthInWS 设置 [min, max] 区间的列宽。
//
// 关键实现选择：**逐列写 <col min="N" max="N"/>，不合并为区间**。
// 合并成 `<col min="4" max="6">` 在 OOXML 规范上合法，但许多消费方
// （实测 openpyxl 即如此）把 column_dimensions 当稀疏字典处理，只为显式
// 出现的列建条目，区间内的其它列会被读成"未设置"。
// openpyxl 自身即使宽度相同也逐列展开（min=max），说明这是被广泛验证的
// 实现所选择的保守形式 —— 牺牲一点压缩换最大互操作性。
func setColWidthInWS(ws string, min, max int, width float64) string {
	if min > max {
		min, max = max, min
	}
	w := strconv.FormatFloat(width, 'f', -1, 64)
	var sb strings.Builder
	for c := min; c <= max; c++ {
		sb.WriteString(fmt.Sprintf(`<col min="%d" max="%d" width="%s" customWidth="1"/>`, c, c, w))
	}
	newCols := sb.String()
	re := regexp.MustCompile(`(?s)<cols\b[^>]*>(.*?)</cols>`)
	if m := re.FindStringSubmatch(ws); m != nil {
		inner := m[1]
		// 移除与被设区间重叠的现有 <col>
		inner = regexp.MustCompile(`<col\b[^>]*\bmin="(\d+)"[^>]*\bmax="(\d+)"[^>]*/?>`).ReplaceAllStringFunc(inner, func(c string) string {
			sm := regexp.MustCompile(`\bmin="(\d+)"`).FindStringSubmatch(c)
			em := regexp.MustCompile(`\bmax="(\d+)"`).FindStringSubmatch(c)
			if sm == nil || em == nil {
				return c
			}
			cmin, _ := strconv.Atoi(sm[1])
			cmax, _ := strconv.Atoi(em[1])
			if cmin > max || cmax < min {
				return c
			}
			return ""
		})
		inner = inner + newCols
		body := `<cols>` + inner + `</cols>`
		return re.ReplaceAllString(ws, body)
	}
	block := `<cols>` + newCols + `</cols>`
	// 放在 <sheetData> 之前（合法顺序）
	if i := strings.Index(ws, "<sheetData"); i != -1 {
		return ws[:i] + string(block) + ws[i:]
	}
	return insertAfterSheetData(ws, frag(block))
}

func getColWidthInWS(ws string, col int) float64 {
	re := regexp.MustCompile(`(?s)<cols\b[^>]*>(.*?)</cols>`)
	m := re.FindStringSubmatch(ws)
	if m == nil {
		return 0
	}
	colRe := regexp.MustCompile(`<col\b[^>]*\bmin="(\d+)"[^>]*\bmax="(\d+)"[^>]*?(?:/>|>(.*?)</col>)`)
	for _, cm := range colRe.FindAllStringSubmatch(m[1], -1) {
		cmin, _ := strconv.Atoi(cm[1])
		cmax, _ := strconv.Atoi(cm[2])
		if col >= cmin && col <= cmax {
			if w := attrOf(cm[0], "width"); w != "" {
				if f, err := strconv.ParseFloat(w, 64); err == nil {
					return f
				}
			}
		}
	}
	return 0
}

// modifyRowTag 定位 <row r="N"> 或 <row r="N"/> 的起始标签，对标签头（不含结束符）
// 应用 fn 改造，并正确保留「自闭合」或「带内容」两种形态。
//
// 关键修正：此前用 (<?/>) 捕获结束符，但 [^>]* 是贪婪的、且 / 不属于 >，
// 会把自闭合斜杠 / 当作属性字符吞掉，导致 m[2] 误判为 ">" 而非 "/>"，
// 进而把 hidden 等属性插到 /> 之后，生成非法 XML（如 <row ...1"/ hidden="1">）。
// 此处改用 (?s)...(?:/>|>) 精确匹配完整结束符，并据末尾字符判定形态。
// fn 接收去掉结束符的标签头（如 `<row r="3" ht="30" customHeight="1"`），返回改造后的标签头。
func modifyRowTag(ws string, row int, fn func(head string) string) string {
	rowRe := regexp.MustCompile(`(?s)<row\b[^>]*?\br="` + strconv.Itoa(row) + `"[^>]*?(?:/>|>)`)
	idx := rowRe.FindStringIndex(ws)
	if idx == nil {
		return ws
	}
	start, end := idx[0], idx[1]
	openTag := ws[start:end]
	if strings.HasSuffix(openTag, "/>") {
		// 自闭合空行：改写标签头后仍以 /> 收尾
		head := strings.TrimSuffix(openTag, "/")
		head = fn(head)
		return ws[:start] + head + `/>` + ws[end:]
	}
	// 带内容的行：定位配对的 </row>
	closeIdx := strings.Index(ws[end:], "</row>")
	if closeIdx == -1 {
		// 异常：无闭合标签，仅在起始标签内追加属性
		head := strings.TrimSuffix(openTag, ">")
		head = fn(head)
		return ws[:start] + head + `>` + ws[end:]
	}
	closeStart := end + closeIdx
	content := ws[end:closeStart]
	head := strings.TrimSuffix(openTag, ">")
	head = fn(head)
	return ws[:start] + head + `>` + content + `</row>` + ws[closeStart+len("</row>"):]
}

// rowExists 判断工作表中是否已存在指定行号的元素。
func rowExists(ws string, row int) bool {
	rowRe := regexp.MustCompile(`(?s)<row\b[^>]*?\br="` + strconv.Itoa(row) + `"[^>]*?(?:/>|>)`)
	return rowRe.MatchString(ws)
}

func setRowHeightInWS(ws string, row int, height float64, custom bool) string {
	h := strconv.FormatFloat(height, 'f', -1, 64)
	if !rowExists(ws, row) {
		// 行不存在：插入带高度的空行
		newRow := `<row r="` + strconv.Itoa(row) + `" ht="` + h + `" customHeight="1"/>`
		return insertRowElement(ws, row, frag(newRow))
	}
	return modifyRowTag(ws, row, func(head string) string {
		head = regexp.MustCompile(`\s+ht="[^"]*"`).ReplaceAllString(head, "")
		head = regexp.MustCompile(`\s+customHeight="\d"`).ReplaceAllString(head, "")
		head = strings.TrimRight(head, " ") + ` ht="` + h + `" customHeight="1"`
		return head
	})
}

func setRowHiddenInWS(ws string, row int, hidden bool) string {
	hv := ""
	if hidden {
		hv = ` hidden="1"`
	}
	if !rowExists(ws, row) {
		newRow := `<row r="` + strconv.Itoa(row) + `"` + hv + `/>`
		return insertRowElement(ws, row, frag(newRow))
	}
	return modifyRowTag(ws, row, func(head string) string {
		head = regexp.MustCompile(`\s+hidden="\d"`).ReplaceAllString(head, "")
		if hidden {
			head = strings.TrimRight(head, " ") + hv
		}
		return head
	})
}

func setColHiddenInWS(ws string, col int, hidden bool) string {
	updated := false
	re := regexp.MustCompile(`(?s)<cols\b[^>]*>(.*?)</cols>`)
	if m := re.FindStringSubmatch(ws); m != nil {
		inner := m[1]
		// 匹配 <col .../> 或 <col ...>（列元素无内容，统一重建为自闭合标签，
		// 确保 hidden 属性落在标签内部、绝不出现在自闭合 /> 之后导致 XML 非法）。
		colRe := regexp.MustCompile(`<col\b[^>]*\bmin="` + strconv.Itoa(col) + `"[^>]*\bmax="` + strconv.Itoa(col) + `"[^>]*/?>`)
		newInner := colRe.ReplaceAllStringFunc(inner, func(c string) string {
			// 去掉已有 hidden 属性，并裁掉结尾的 /> 或 >，统一重建成自闭合标签
			c = regexp.MustCompile(`\s+hidden="\d"`).ReplaceAllString(c, "")
			c = strings.TrimRight(c, " />")
			if hidden {
				c = c + ` hidden="1"`
			}
			updated = true
			return c + `/>`
		})
		if updated {
			return re.ReplaceAllString(ws, `<cols>`+newInner+`</cols>`)
		}
		// 该列无 <col> 定义：追加一个仅带隐藏的 <col>
		inner = inner + fmt.Sprintf(`<col min="%d" max="%d" hidden="1"/>`, col, col)
		return re.ReplaceAllString(ws, `<cols>`+inner+`</cols>`)
	}
	block := fmt.Sprintf(`<cols><col min="%d" max="%d" hidden="1"/></cols>`, col, col)
	if i := strings.Index(ws, "<sheetData"); i != -1 {
		return ws[:i] + string(block) + ws[i:]
	}
	return insertAfterSheetData(ws, frag(block))
}

// ---------- 冻结窗格 ----------

// FreezePanes 在 ref 单元格处冻结窗格（如 "A2" 冻结首行，"B1" 冻结首列，"B2" 双冻结）。
func (s *WorkSheet) FreezePanes(ref string) error {
	col, row, err := parseCellRef(ref)
	if err != nil {
		return err
	}
	s.setWS(freezePanesInWS(string(s.ws()), col, row+1))
	return nil
}

func freezePanesInWS(ws string, col0, row1 int) string {
	xSplit := col0 // 0 基列号 = 冻结列数
	ySplit := row1 - 1
	ws = ensureWorksheetNamespaces(ws)
	// 构建 <pane>
	var pane strings.Builder
	pane.WriteString("<pane")
	var activePane string
	if xSplit > 0 && ySplit > 0 {
		pane.WriteString(fmt.Sprintf(` xSplit="%d" ySplit="%d"`, xSplit, ySplit))
		pane.WriteString(` topLeftCell="` + colNumToLetters(xSplit) + strconv.Itoa(ySplit+1) + `"`)
		activePane = "bottomRight"
	} else if ySplit > 0 {
		pane.WriteString(fmt.Sprintf(` ySplit="%d"`, ySplit))
		pane.WriteString(` topLeftCell="A` + strconv.Itoa(ySplit+1) + `"`)
		activePane = "bottomLeft"
	} else if xSplit > 0 {
		pane.WriteString(fmt.Sprintf(` xSplit="%d"`, xSplit))
		pane.WriteString(` topLeftCell="` + colNumToLetters(xSplit) + `1"`)
		activePane = "topRight"
	} else {
		// 无冻结
		return removePaneInWS(ws)
	}
	pane.WriteString(` activePane="` + activePane + `" state="frozen"/>`)
	// 插入到 <sheetView> 内（首子元素）
	svRe := regexp.MustCompile(`(?s)(<sheetViews?>)(.*?)(</sheetViews>)`)
	if m := svRe.FindStringSubmatch(ws); m != nil {
		inner := m[2]
		// 仅保留第一个 <sheetView>
		svOpenRe := regexp.MustCompile(`(?s)(<sheetView\b[^>]*>)`)
		if sv := svOpenRe.FindStringSubmatch(inner); sv != nil {
			newInner := strings.Replace(inner, sv[1], sv[1]+pane.String(), 1)
			return svRe.ReplaceAllString(ws, m[1]+newInner+m[3])
		}
		// 没有 <sheetView>：补一个
		newInner := `<sheetView>` + pane.String() + `</sheetView>`
		return svRe.ReplaceAllString(ws, m[1]+newInner+m[3])
	}
	// 没有 <sheetViews>：创建
	block := `<sheetViews><sheetView>` + pane.String() + `</sheetView></sheetViews>`
	if i := strings.Index(ws, "<sheetData"); i != -1 {
		return ws[:i] + string(block) + ws[i:]
	}
	return insertAfterSheetData(ws, frag(block))
}

func removePaneInWS(ws string) string {
	ws = regexp.MustCompile(`(?s)<pane\b[^>]*/>`).ReplaceAllString(ws, "")
	return ws
}

// ---------- 超链接 ----------

// AddHyperlink 在 cell 单元格添加指向 url 的外部超链接（可选 displayText 写入单元格）。
func (s *WorkSheet) AddHyperlink(cell, url, displayText string) error {
	// 1) 可选显示文本（先写，避免覆盖后续写入的 <hyperlinks> 块）
	if displayText != "" {
		if err := s.SetCellStr(cell, displayText); err != nil {
			return err
		}
	}
	// 2) worksheet rels 添加超链接关系
	sheetRelsFile := "xl/worksheets/_rels/" + path.Base(s.file) + ".rels"
	rels := loadRels(s.fm(), sheetRelsFile)
	if rels == nil {
		rels = &Relationships{XMLName: xmlNameRelationships()}
	}
	rid := fmt.Sprintf("rId%d", getMaxRId(rels)+1)
	rels.Relationship = append(rels.Relationship, Relationship{
		ID:         rid,
		Type:       "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
		Target:     url,
		TargetMode: "External",
	})
	s.fm()[sheetRelsFile] = serializeRelationships(rels)
	// 3) worksheet 加 <hyperlinks>
	ws := string(s.ws())
	ws = ensureWorksheetNamespaces(ws)
	// cell 规范化并校验：非法坐标直接拒绝，避免把未校验的用户串写入属性
	ref := newCellRef(cell)
	if ref == "" {
		return fmt.Errorf("无效的单元格引用: %q", cell)
	}
	hyperlinkTag := frag(`<hyperlink ref="` + string(ref) + `" r:id="` + safeAttr(rid) + `"/>`)
	if strings.Contains(ws, "<hyperlinks") {
		ws = regexp.MustCompile(`(?s)(<hyperlinks\b[^>]*>)(.*?)(</hyperlinks>)`).ReplaceAllString(ws,
			`$1$2`+string(hyperlinkTag)+`$3`)
	} else {
		ws = insertAfterSheetData(ws, frag(`<hyperlinks>`+string(hyperlinkTag)+`</hyperlinks>`))
	}
	s.setWS(ws)
	return nil
}

// ---------- 自动筛选 ----------

// AutoFilter 在 ref 区域添加自动筛选器。
func (s *WorkSheet) AutoFilter(ref string) error {
	// 区域引用先经校验规范化，未校验的坐标不写入 XML
	r := newRangeRef(ref)
	if r == "" {
		return fmt.Errorf("无效的区域引用: %q", ref)
	}
	ws := string(s.ws())
	tag := `<autoFilter ref="` + string(r) + `"/>`
	if strings.Contains(ws, "<autoFilter") {
		ws = regexp.MustCompile(`(?s)<autoFilter\b[^>]*/>`).ReplaceAllString(ws, tag)
	} else {
		ws = insertAfterSheetData(ws, frag(tag))
	}
	s.setWS(ws)
	return nil
}

// ---------- 数据验证 ----------

// validDVTypes 是 OOXML 规范（CT_DataValidation/@type）认可的类型。
// 与条件格式同理：这些值原样写入 XML，写错会让消费方拒绝加载整个工作簿。
var validDVTypes = map[string]bool{
	"none": true, "whole": true, "decimal": true, "list": true,
	"date": true, "time": true, "textLength": true, "custom": true,
}

// validDVOperators 是 OOXML 规范（ST_DataValidationOperator）认可的比较运算符。
var validDVOperators = map[string]bool{
	"between": true, "notBetween": true, "equal": true, "notEqual": true,
	"greaterThan": true, "lessThan": true, "greaterThanOrEqual": true,
	"lessThanOrEqual": true,
}

// AddDataValidation 为 ref 区域添加数据验证。typ 为 "list"/"whole"/"decimal"/"date" 等；
// formula1/formula2 为公式（列表用 "a,b,c" 形式时调用方应自行包裹引号，或用 AddDataValidationList）。
//
// typ 与 op 必须是 OOXML 规范取值（见 validDVTypes / validDVOperators）——
// 它们原样写入 XML 属性，非法取值会导致 openpyxl/Excel 拒绝加载工作簿，故提前拦截。
func (s *WorkSheet) AddDataValidation(ref, typ, op, formula1, formula2 string, allowBlank bool) error {
	if !validDVTypes[typ] {
		return fmt.Errorf("无效的数据验证类型 %q，合法的 OOXML 取值如："+
			"list / whole / decimal / date / time / textLength / custom", typ)
	}
	if op != "" && !validDVOperators[op] {
		return fmt.Errorf("无效的数据验证比较运算符 %q，合法的 OOXML 取值如："+
			"between / equal / notEqual / greaterThan / lessThan / "+
			"greaterThanOrEqual / lessThanOrEqual", op)
	}
	// 区域引用校验：未校验的坐标会直接进 XML 属性
	r := newRangeRef(ref)
	if r == "" {
		return fmt.Errorf("无效的数据验证区域引用: %q", ref)
	}
	ref = string(r)
	ws := string(s.ws())
	blank := "0"
	if allowBlank {
		blank = "1"
	}
	var bld strings.Builder
	bld.WriteString(`<dataValidation type="` + safeAttr(typ) + `" allowBlank="` + blank + `"`)
	if op != "" {
		bld.WriteString(` operator="` + safeAttr(op) + `"`)
	}
	bld.WriteString(` sqref="` + safeAttr(ref) + `">`)
	if formula1 != "" {
		bld.WriteString(`<formula1>` + safeText(formula1) + `</formula1>`)
	}
	if formula2 != "" {
		bld.WriteString(`<formula2>` + safeText(formula2) + `</formula2>`)
	}
	bld.WriteString(`</dataValidation>`)
	if strings.Contains(ws, "<dataValidations") {
		ws = regexp.MustCompile(`(?s)(<dataValidations\b[^>]*>)(.*?)(</dataValidations>)`).ReplaceAllStringFunc(ws, func(m string) string {
			mm := regexp.MustCompile(`(?s)(<dataValidations\b[^>]*>)(.*?)(</dataValidations>)`).FindStringSubmatch(m)
			count := strings.Count(mm[2], "<dataValidation") + 1
			open := regexp.MustCompile(`<dataValidations\b[^>]*>`).ReplaceAllString(mm[1], fmt.Sprintf(`<dataValidations count="%d">`, count))
			return open + mm[2] + bld.String() + mm[3]
		})
	} else {
		ws = insertAfterSheetData(ws, frag(`<dataValidations count="1">`+bld.String()+`</dataValidations>`))
	}
	s.setWS(ws)
	return nil
}

// AddDataValidationList 为 ref 区域添加「下拉列表」数据验证（values 为可选项）。
func (s *WorkSheet) AddDataValidationList(ref string, values []string, allowBlank bool) error {
	list := `"` + strings.Join(values, ",") + `"`
	return s.AddDataValidation(ref, "list", "", list, "", allowBlank)
}

// ---------- 条件格式 ----------

// validCFTypes 是 OOXML 规范（CT_CfRule/@type）认可的条件格式类型。
//
// 这些值**不能随意写**：写错会让 openpyxl 直接拒绝加载整个工作簿
// （实测 type="cell" 即导致 "Unable to read workbook"），Excel 同样会判损坏。
// 常见的误用是把 "cellIs" 写成 "cell"（直觉命名），故在此显式校验并给出映射提示。
var validCFTypes = map[string]bool{
	"expression": true, "cellIs": true, "duplicateValues": true,
	"uniqueValues": true, "top10": true, "aboveAverage": true,
	"containsText": true, "notContainsText": true, "beginsWith": true,
	"endsWith": true, "containsBlanks": true, "notContainsBlanks": true,
	"containsErrors": true, "notContainsErrors": true,
	"colorScale": true, "dataBar": true, "iconSet": true,
	"timePeriod": true,
}

// cfTypeAliases 常见误写到正确取值的映射，用于给出更友好的错误提示。
var cfTypeAliases = map[string]string{
	"cell":         "cellIs",
	"cellis":       "cellIs",
	"formula":      "expression",
	"duplicate":    "duplicateValues",
	"unique":       "uniqueValues",
	"textcontains": "containsText",
	"icon":         "iconSet",
	"databar":      "dataBar",
	"colorscale":   "colorScale",
}

// SetConditionalFormat 为 ref 区域添加一条条件格式。
//
// cfType 必须是 OOXML 规范取值（见 validCFTypes），如 "expression"/"cellIs"/
// "duplicateValues"；formula 为条件公式（如 "A1>100"）；st 为命中时应用的样式
// （写入 styles.xml 的 dxfs 差分格式）。
//
// cfType 会**原样写入 XML**，因此非法取值不能靠转义兜底 —— 写错会让消费方
// 拒绝加载整个工作簿。这里提前校验并对常见误写给出正确取值的提示。
func (s *WorkSheet) SetConditionalFormat(ref, cfType, formula string, priority int, st Style) error {
	if !validCFTypes[cfType] {
		if fixed, ok := cfTypeAliases[strings.ToLower(cfType)]; ok {
			return fmt.Errorf("无效的条件格式类型 %q，应为 %q（OOXML 规范要求；"+
				"写成 %q 会导致 openpyxl/Excel 拒绝加载整个工作簿）",
				cfType, fixed, cfType)
		}
		return fmt.Errorf("无效的条件格式类型 %q，合法的 OOXML 取值如："+
			"expression / cellIs / duplicateValues / containsText / colorScale / dataBar / iconSet",
			cfType)
	}
	// 区域引用也需校验：未校验的坐标会直接进 XML 属性
	r := newRangeRef(ref)
	if r == "" {
		return fmt.Errorf("无效的条件格式区域引用: %q", ref)
	}
	ws := string(s.ws())
	fm := s.fm()
	dxfId, err := addDxfToStyles(fm, st)
	if err != nil {
		return err
	}
	rule := fmt.Sprintf(`<cfRule type="%s" dxfId="%d" priority="%d"><formula>%s</formula></cfRule>`,
		safeAttr(cfType), dxfId, priority, safeText(formula))
	ws = insertAfterSheetData(ws, frag(`<conditionalFormatting sqref="`+string(r)+`">`+rule+`</conditionalFormatting>`))
	s.setWS(ws)
	return nil
}

// addDxfToStyles 在 styles.xml 的 <dxfs> 中追加一条差分格式，返回其索引（不去重，便于多规则各自独立）。
func addDxfToStyles(fileMap map[string][]byte, st Style) (int, error) {
	stylesPath := "xl/styles.xml"
	if _, ok := fileMap[stylesPath]; !ok {
		fileMap[stylesPath] = []byte(defaultStylesXML())
		fileMap["[Content_Types].xml"] = insertOverrideInContentTypes(fileMap["[Content_Types].xml"],
			"/"+stylesPath, "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
	}
	xml := string(fileMap[stylesPath])
	// 条件格式的 dxf 里也会写填充图案，同样需要先校验
	// 条件格式的 dxf 同样走 buildFont/buildBorder/buildFill，枚举值同样要校验
	if err := validateStyleEnums(st); err != nil {
		return 0, err
	}
	if st.Fill != nil && st.Fill.Color != "" {
		if _, err := validatePatternType(st.Fill.PatternType); err != nil {
			return 0, err
		}
	}
	dxf := buildDXF(st)
	re := regexp.MustCompile(`(?s)<dxfs\b[^>]*>(.*?)</dxfs>`)
	if m := re.FindStringSubmatch(xml); m != nil {
		inner := m[1]
		count := strings.Count(inner, "<dxf")
		newInner := inner + dxf
		body := fmt.Sprintf(`<dxfs count="%d">%s</dxfs>`, count+1, newInner)
		xml = re.ReplaceAllString(xml, body)
		fileMap[stylesPath] = []byte(xml)
		return count, nil
	}
	// 缺失 <dxfs>：插入（放在 cellXfs 之后）
	xml = upsertBlock(xml, "dxfs", []string{dxf}, "dxf", func(n int) string {
		return fmt.Sprintf(`<dxfs count="%d">`, n)
	})
	fileMap[stylesPath] = []byte(xml)
	return 0, nil
}

// buildDXF 由 Style 构造 <dxf> 差分格式内容。
func buildDXF(st Style) string {
	var b strings.Builder
	b.WriteString("<dxf>")
	if st.Font != nil && !st.Font.isEmpty() {
		b.WriteString(buildFont(st.Font))
	}
	if st.Fill != nil && st.Fill.Color != "" {
		b.WriteString(buildFill(st.Fill))
	}
	if st.Border != nil && !st.Border.isEmpty() {
		b.WriteString(buildBorder(st.Border))
	}
	b.WriteString("</dxf>")
	return b.String()
}

// ---------- 表格(ListObject) ----------

// AddTable 把 ref 区域创建为一张结构化表格（ListObject），name 为表名（如 "Table1"）。
// 表头取自区域首行的单元格值（为空则用 Column1..）。
func (s *WorkSheet) AddTable(ref, name string) error {
	ws := string(s.ws())
	fm := s.fm()
	c1, r1, c2, _, err := parseRangeRef(ref)
	if err != nil {
		return err
	}
	// 表头名
	ncols := c2 - c1 + 1
	headers := make([]string, ncols)
	for j := 0; j < ncols; j++ {
		cellRef := colNumToLetters(c1+j) + strconv.Itoa(r1+1)
		v := readCellValue(fm, ws, cellRef)
		if v == "" {
			v = "Column" + strconv.Itoa(j+1)
		}
		headers[j] = v
	}
	// 编号与关系
	num := nextFreeNumber(fm, "xl/tables/table", ".xml")
	tableFile := fmt.Sprintf("xl/tables/table%d.xml", num)
	// worksheet rels
	sheetRelsFile := "xl/worksheets/_rels/" + path.Base(s.file) + ".rels"
	rels := loadRels(fm, sheetRelsFile)
	if rels == nil {
		rels = &Relationships{XMLName: xmlNameRelationships()}
	}
	rid := fmt.Sprintf("rId%d", getMaxRId(rels)+1)
	rels.Relationship = append(rels.Relationship, Relationship{
		ID:     rid,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
		Target: "../tables/" + path.Base(tableFile),
	})
	fm[sheetRelsFile] = serializeRelationships(rels)
	// worksheet 加 <tableParts>
	ws = ensureWorksheetNamespaces(ws)
	if strings.Contains(ws, "<tableParts") {
		ws = regexp.MustCompile(`(?s)(<tableParts\b[^>]*>)(.*?)(</tableParts>)`).ReplaceAllString(ws,
			`$1$2<tablePart r:id="`+rid+`"/>$3`)
	} else {
		ws = insertAfterSheetData(ws, frag(`<tableParts count="1"><tablePart r:id="`+rid+`"/></tableParts>`))
	}
	s.setWS(ws)
	// table xml
	fm[tableFile] = []byte(buildTableXML(fm, name, ref, headers))
	fm["[Content_Types].xml"] = insertOverrideInContentTypes(fm["[Content_Types].xml"], "/"+tableFile,
		"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")
	return nil
}

func buildTableXML(fm map[string][]byte, name, ref string, headers []string) string {
	var cols strings.Builder
	cols.WriteString(fmt.Sprintf(`<tableColumns count="%d">`, len(headers)))
	for i, h := range headers {
		cols.WriteString(fmt.Sprintf(`<tableColumn id="%d" name="%s"/>`, i+1, safeAttr(h)))
	}
	cols.WriteString(`</tableColumns>`)
	// 注意：table 的 id 属性是**表内容的一部分**，不是文件名序号。
	// 复制工作表时表内容被整体搬走，若沿用源表的 id 会出现重复 ——
	// 而 table id 必须全局唯一，否则 openpyxl/Excel 会拒绝加载工作簿。
	// 故由调用方按当前工作簿内已有 table 的最大 id +1 重新分配。
	id := nextTableContentID(fm)
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="` +
		strconv.Itoa(id) + `" name="` + safeAttr(name) + `" displayName="` + safeAttr(name) +
		`" ref="` + safeAttr(ref) + `" totalsRowShown="0">` +
		`<autoFilter ref="` + safeAttr(ref) + `"/>` +
		cols.String() +
		`<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>` +
		`</table>`
}

// nextTableContentID 返回当前工作簿内可用的 table id（已用最大值 +1）。
// 扫描所有 xl/tables/table*.xml 的 id 属性取最大值。
func nextTableContentID(fm map[string][]byte) int {
	max := 0
	re := regexp.MustCompile(`<table\b[^>]*\bid="(\d+)"`)
	for name, data := range fm {
		if !strings.HasPrefix(name, "xl/tables/table") || !strings.HasSuffix(name, ".xml") {
			continue
		}
		for _, m := range re.FindAllStringSubmatch(string(data), -1) {
			if n, err := strconv.Atoi(m[1]); err == nil && n > max {
				max = n
			}
		}
	}
	return max + 1
}

// ---------- 批注 ----------

// AddComment 在 cell 单元格添加批注 text（author 为作者，空则用 "Author"）。
func (s *WorkSheet) AddComment(cell, text, author string) error {
	ws := string(s.ws())
	fm := s.fm()
	if author == "" {
		author = "Author"
	}
	// 定位/创建 comments 部件
	commentsFile, rid, err := ensureCommentsPart(fm, s.file)
	if err != nil {
		return err
	}
	cxml := string(fm[commentsFile])
	// 解析 authors / commentList
	var authors []string
	authorRe := regexp.MustCompile(`(?s)<authors>(.*?)</authors>`)
	if m := authorRe.FindStringSubmatch(cxml); m != nil {
		for _, a := range regexp.MustCompile(`<author>(.*?)</author>`).FindAllStringSubmatch(m[1], -1) {
			authors = append(authors, unescapeXML(a[1]))
		}
	}
	authorId := -1
	for i, a := range authors {
		if a == author {
			authorId = i
			break
		}
	}
	if authorId == -1 {
		authorId = len(authors)
		authors = append(authors, author)
	}
	// 删除该 cell 已有批注
	cxml = regexp.MustCompile(`(?s)<comment\b[^>]*\bref="`+regexp.QuoteMeta(cell)+`"[^>]*>.*?</comment>`).ReplaceAllString(cxml, "")
	// 重建 authors
	var ab strings.Builder
	ab.WriteString("<authors>")
	for _, a := range authors {
		ab.WriteString(`<author>` + safeText(a) + `</author>`)
	}
	ab.WriteString("</authors>")
	if strings.Contains(cxml, "<authors>") {
		cxml = authorRe.ReplaceAllString(cxml, ab.String())
	} else {
		cxml = regexp.MustCompile(`(?s)<commentList>`).ReplaceAllStringFunc(cxml, func(m string) string {
			return ab.String() + m
		})
		if !strings.Contains(cxml, "<commentList>") {
			cxml = strings.Replace(cxml, "</comments>", "<commentList></commentList></comments>", 1)
		}
	}
	// 追加 comment
	// cell 规范化并校验：非法坐标直接拒绝
	cref := newCellRef(cell)
	if cref == "" {
		return fmt.Errorf("无效的单元格引用: %q", cell)
	}
	comment := frag(`<comment ref="` + string(cref) + `" authorId="` + strconv.Itoa(authorId) + `"><text><t xml:space="preserve">` +
		safeText(text) + `</t></text></comment>`)
	if strings.Contains(cxml, "<commentList>") {
		cxml = regexp.MustCompile(`(?s)(<commentList>)`).ReplaceAllString(cxml, `$1`+string(comment))
	} else {
		cxml = strings.Replace(cxml, "</comments>", "<commentList>"+string(comment)+"</commentList></comments>", 1)
	}
	fm[commentsFile] = []byte(cxml)
	// worksheet 加 <comments r:id>
	ws = ensureWorksheetNamespaces(ws)
	if !strings.Contains(ws, "<comments") {
		ws = insertAfterSheetData(ws, frag(`<comments r:id="`+rid+`"/>`))
	}
	s.setWS(ws)
	return nil
}

// GetComment 读取 cell 单元格的批注文本（无则返回空）。
func (s *WorkSheet) GetComment(cell string) string {
	fm := s.fm()
	commentsFile, _, err := locateCommentsPart(fm, s.file)
	if err != nil {
		return ""
	}
	cxml := string(fm[commentsFile])
	m := regexp.MustCompile(`(?s)<comment\b[^>]*\bref="` + regexp.QuoteMeta(cell) + `"[^>]*>(.*?)</comment>`).FindStringSubmatch(cxml)
	if m == nil {
		return ""
	}
	t := regexp.MustCompile(`(?s)<t\b[^>]*>(.*?)</t>`).FindStringSubmatch(m[1])
	if t == nil {
		return ""
	}
	return unescapeXML(t[1])
}

func ensureCommentsPart(fm map[string][]byte, sheetFile string) (file string, rid string, err error) {
	if f, r, e := locateCommentsPart(fm, sheetFile); e == nil {
		return f, r, nil
	}
	// 新建 xl/comments/commentsN.xml（与 rel Target "../comments/..." 对应）
	num := nextFreeNumber(fm, "xl/comments/comments", ".xml")
	commentsFile := fmt.Sprintf("xl/comments/comments%d.xml", num)
	if _, ok := fm["xl/comments"]; !ok {
		// 目录无需显式创建（zip 按完整路径存），但确保 Content_Types 对目录无要求
	}
	fm[commentsFile] = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">` +
		`<authors><author>Author</author></authors><commentList></commentList></comments>`)
	fm["[Content_Types].xml"] = insertOverrideInContentTypes(fm["[Content_Types].xml"], "/"+commentsFile,
		"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
	// worksheet rels
	sheetRelsFile := "xl/worksheets/_rels/" + path.Base(sheetFile) + ".rels"
	rels := loadRels(fm, sheetRelsFile)
	if rels == nil {
		rels = &Relationships{XMLName: xmlNameRelationships()}
	}
	rid = fmt.Sprintf("rId%d", getMaxRId(rels)+1)
	rels.Relationship = append(rels.Relationship, Relationship{
		ID:     rid,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
		Target: "../comments/" + path.Base(commentsFile),
	})
	fm[sheetRelsFile] = serializeRelationships(rels)
	return commentsFile, rid, nil
}

func locateCommentsPart(fm map[string][]byte, sheetFile string) (file string, rid string, err error) {
	sheetRelsFile := "xl/worksheets/_rels/" + path.Base(sheetFile) + ".rels"
	rels := loadRels(fm, sheetRelsFile)
	if rels == nil {
		return "", "", fmt.Errorf("无工作表关系")
	}
	for _, r := range rels.Relationship {
		if r.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" {
			return resolveTarget("xl/worksheets", r.Target), r.ID, nil
		}
	}
	return "", "", fmt.Errorf("无批注部件")
}

// ---------- 工作表属性（标签色/网格线/缩放） ----------

// SheetProps 工作表展示属性。
type SheetProps struct {
	TabColor      string // ARGB 十六进制（如 "FFFF0000"），空=不改变
	ShowGridLines *bool  // 是否显示网格线
	Zoom          int    // 缩放百分比（如 120），0=不改变
	// 打印 / 页面设置（见 print.go）
	PageMargins    *PageMargins  // 页边距
	PageSetup      *PageSetup    // 页面设置
	PrintArea      string        // 打印区域，如 "A1:D10"
	HeaderFooter   *HeaderFooter // 页眉页脚
	PrintTitleRows string        // 重复打印行，如 "1:1"
}

// SetProps 设置工作表展示属性。
func (s *WorkSheet) SetProps(opts SheetProps) error {
	ws := string(s.ws())
	if opts.ShowGridLines != nil || opts.Zoom != 0 {
		ws = ensureWorksheetNamespaces(ws)
		ws = ensureSheetViewWrap(ws)
		grid := ""
		if opts.ShowGridLines != nil {
			if *opts.ShowGridLines {
				grid = ` showGridLines="1"`
			} else {
				grid = ` showGridLines="0"`
			}
		}
		zoom := ""
		if opts.Zoom != 0 {
			zoom = fmt.Sprintf(` zoomScale="%d"`, opts.Zoom)
		}
		ws = regexp.MustCompile(`(?s)(<sheetView\b)`).ReplaceAllString(ws, `${1}`+grid+zoom)
	}

	// 打印属性（worksheet 级）
	if opts.PageMargins != nil || opts.PageSetup != nil || opts.HeaderFooter != nil {
		ws = ensureWorksheetNamespaces(ws)
		ws = setPageMarginsInWS(ws, opts.PageMargins)
		ws = setPageSetupInWS(ws, opts.PageSetup)
		ws = setHeaderFooterInWS(ws, opts.HeaderFooter)
	}

	// 标签色（OOXML 规范位置：worksheet XML 的 <sheetPr><tabColor>）
	if opts.TabColor != "" {
		ws = setTabColorInWS(ws, opts.TabColor)
	}
	s.setWS(ws)

	// 打印区域 / 重复行（workbook 级 definedName）
	if opts.PrintArea != "" || opts.PrintTitleRows != "" {
		idx := s.sheetLocalIndex()
		if idx < 0 {
			return fmt.Errorf("找不到工作表 %q 的序号", s.Name())
		}
		wbXML := s.fm()["xl/workbook.xml"]
		if opts.PrintArea != "" {
			wbXML = setOrReplaceDefinedName(wbXML, "_xlnm.Print_Area", idx, fmt.Sprintf("%s!%s", s.Name(), absRange(opts.PrintArea)))
		}
		if opts.PrintTitleRows != "" {
			wbXML = setOrReplaceDefinedName(wbXML, "_xlnm.Print_Titles", idx, fmt.Sprintf("%s!%s", s.Name(), absRangeRows(opts.PrintTitleRows)))
		}
		s.fm()["xl/workbook.xml"] = wbXML
	}
	return nil
}

// ensureSheetViewWrap 确保存在 <sheetViews><sheetView> 包裹（无则创建并包裹 sheetData 之前）。
func ensureSheetViewWrap(ws string) string {
	if strings.Contains(ws, "<sheetViews") {
		return ws
	}
	block := `<sheetViews><sheetView></sheetView></sheetViews>`
	if i := strings.Index(ws, "<sheetData"); i != -1 {
		return ws[:i] + string(block) + ws[i:]
	}
	return insertAfterSheetData(ws, frag(block))
}

// ---------- 工作表保护 ----------

// Protect 开启工作表保护。
//
// password 语义（**注意与 ProtectWorkbook 的差异**）：这里要求传**已哈希**的
// Excel password verifier 值（如明文 "1234" 的哈希是 "CC3D"），本函数不做哈希。
// 这是本库的既有约定，为保持兼容而保留。
//
// 新代码建议用 ProtectWithPassword —— 它收**明文**并在内部哈希，
// 与 ProtectWorkbook 语义一致，不会搞混。为空则不设密码（仅锁结构）。
func (s *WorkSheet) Protect(password string) error {
	return s.protectWithHash(password)
}

// ProtectWithPassword 开启工作表保护，password 传**明文**，内部按 Excel 标准
// password verifier 算法哈希后写入。
//
// 推荐入口：与 ProtectWorkbook 语义一致（都收明文），避免"工作表要哈希、
// 工作簿不用"这种容易踩的坑 —— 传错会导致用户以为设了密码、实际解不开。
func (s *WorkSheet) ProtectWithPassword(plaintext string) error {
	if plaintext == "" {
		return s.protectWithHash("")
	}
	return s.protectWithHash(hashExcelPassword(plaintext))
}

// protectWithHash 写入 <sheetProtection>，passwordHash 为已哈希值（可空）。
func (s *WorkSheet) protectWithHash(passwordHash string) error {
	ws := string(s.ws())
	ws = ensureWorksheetNamespaces(ws)
	tag := `<sheetProtection sheet="1" selectLockedCells="1" selectUnlockedCells="1"`
	if passwordHash != "" {
		tag += ` password="` + safeAttr(passwordHash) + `"`
	}
	tag += `/>`
	if strings.Contains(ws, "<sheetProtection") {
		ws = regexp.MustCompile(`(?s)<sheetProtection\b[^>]*/>`).ReplaceAllString(ws, tag)
	} else {
		ws = insertAfterSheetData(ws, frag(tag))
	}
	s.setWS(ws)
	return nil
}

// ---------- 行列分组大纲 ----------

// GroupRows 把 [r1,r2] 行设为分组大纲级别 level（1 基）。
func (s *WorkSheet) GroupRows(r1, r2, level int) error {
	s.setWS(groupRowsInWS(string(s.ws()), r1, r2, level))
	return nil
}

// GroupCols 把 [c1,c2] 列设为分组大纲级别 level（1 基）。
func (s *WorkSheet) GroupCols(c1, c2, level int) error {
	s.setWS(setColWidthInWSGroup(string(s.ws()), c1, c2, level))
	return nil
}

func groupRowsInWS(ws string, r1, r2, level int) string {
	// 逐行定位并改写起始标签，复用 modifyRowTag 保证自闭合/带内容两种形态都正确，
	// 规避 (<?/>) 捕获把 / 误吞导致 m[2] 误判的同类缺陷。
	for row := r1; row <= r2; row++ {
		ws = modifyRowTag(ws, row, func(head string) string {
			head = regexp.MustCompile(`\s+outlineLevel="\d"`).ReplaceAllString(head, "")
			head = strings.TrimRight(head, " ") + fmt.Sprintf(` outlineLevel="%d"`, level)
			return head
		})
	}
	return ws
}

func setColWidthInWSGroup(ws string, min, max, level int) string {
	re := regexp.MustCompile(`(?s)<cols\b[^>]*>(.*?)</cols>`)
	if m := re.FindStringSubmatch(ws); m != nil {
		inner := m[1]
		// 收集所有现存的 <col>，按「重叠拆分」原则重建：
		// 与 [min,max] 重叠的 col 拆成「左段(原属性,无分组) + 重叠段(带 outlineLevel) + 右段(原属性,无分组)」，
		// 以保留原有 width 等属性，同时让分组级别覆盖目标区间生效（避免被前面的宽定义遮蔽）。
		colRe := regexp.MustCompile(`<col\b[^>]*\bmin="(\d+)"[^>]*\bmax="(\d+)"[^>]*/?>`)
		var rebuilt []string
		for _, cm := range colRe.FindAllStringSubmatch(inner, -1) {
			cmin, _ := strconv.Atoi(cm[1])
			cmax, _ := strconv.Atoi(cm[2])
			if cmax < min || cmin > max {
				rebuilt = append(rebuilt, cm[0]) // 完全不重叠，原样保留
				continue
			}
			if cmin < min {
				rebuilt = append(rebuilt, rebuildColWith(cm[0], cmin, min-1, 0))
			}
			lo := maxInt(cmin, min)
			hi := minInt(cmax, max)
			rebuilt = append(rebuilt, rebuildColWith(cm[0], lo, hi, level))
			if cmax > max {
				rebuilt = append(rebuilt, rebuildColWith(cm[0], max+1, cmax, 0))
			}
		}
		// 注：与 [min,max] 重叠的 col 已在上面生成覆盖该区间的带 outlineLevel 段，
		// 因此无需再额外追加精确区间，避免产生重复 <col> 定义。
		newInner := strings.Join(rebuilt, "")
		return re.ReplaceAllString(ws, `<cols>`+newInner+`</cols>`)
	}
	block := fmt.Sprintf(`<cols><col min="%d" max="%d" outlineLevel="%d"/></cols>`, min, max, level)
	if i := strings.Index(ws, "<sheetData"); i != -1 {
		return ws[:i] + string(block) + ws[i:]
	}
	return insertAfterSheetData(ws, frag(block))
}

// rebuildColWith 基于已有 col 原始串，重设 min/max 并视 level>0 追加 outlineLevel，
// 保留其余属性（如 width/customWidth），避免丢失原有列宽。
func rebuildColWith(raw string, newMin, newMax, level int) string {
	s := raw
	s = regexp.MustCompile(`\bmin="\d+"`).ReplaceAllString(s, fmt.Sprintf(`min="%d"`, newMin))
	s = regexp.MustCompile(`\bmax="\d+"`).ReplaceAllString(s, fmt.Sprintf(`max="%d"`, newMax))
	s = regexp.MustCompile(`\s+outlineLevel="\d"`).ReplaceAllString(s, "")
	if level > 0 {
		s = regexp.MustCompile(`(\bmax="\d+")`).ReplaceAllString(s, `${1} outlineLevel="`+strconv.Itoa(level)+`"`)
	}
	return s
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}
func minInt(a, b int) int {
	if a < b {
		return a
	}
	return b
}

// ---------- 查找 / 替换 ----------

// Replace 在工作表内把所有单元格文本中的 old 替换为 new，返回替换次数。
// 覆盖共享字符串、内联字符串与公式单元格的结果文本。
func (s *WorkSheet) Replace(old, new string) (int, error) {
	fm := s.fm()
	ws := string(s.ws())
	count := 0
	// 1) 共享字符串
	if ss, ok := fm["xl/sharedStrings.xml"]; ok {
		out, n := replaceInSharedStrings(string(ss), old, new)
		if n > 0 {
			fm["xl/sharedStrings.xml"] = []byte(out)
			count += n
		}
	}
	// 2) 工作表内联 <is><t> 与 <v>（非公式）文本
	re := regexp.MustCompile(`(?s)(<is>.*?</is>)|(<v>.*?</v>)`)
	ws = re.ReplaceAllStringFunc(ws, func(m string) string {
		if strings.Contains(m, "<f>") {
			return m
		}
		if strings.Contains(m, old) {
			count++
			return strings.Replace(m, old, new, -1)
		}
		return m
	})
	s.setWS(ws)
	return count, nil
}

func replaceInSharedStrings(ss, old, new string) (string, int) {
	count := 0
	re := regexp.MustCompile(`(?s)(<t\b[^>]*>)(.*?)(</t>)`)
	out := re.ReplaceAllStringFunc(ss, func(m string) string {
		sm := re.FindStringSubmatch(m)
		if sm == nil {
			return m
		}
		if strings.Contains(sm[2], old) {
			count++
			return sm[1] + strings.Replace(sm[2], old, new, -1) + sm[3]
		}
		return m
	})
	return out, count
}

// ---------- 文档属性 ----------

// SetDocProps 设置核心文档属性（core.xml）。props 键支持 title/subject/creator/keywords/description。
func (b *Book) SetDocProps(props map[string]string) error {
	const ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" ` +
		`xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" ` +
		`xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"></cp:coreProperties>`
	path := "docProps/core.xml"
	xml := ct
	if v, ok := b.fileMap[path]; ok {
		xml = string(v)
	}
	for k, v := range props {
		tag := ""
		switch k {
		case "title":
			tag = "dc:title"
		case "subject":
			tag = "dc:subject"
		case "creator":
			tag = "dc:creator"
		case "keywords":
			tag = "cp:keywords"
		case "description":
			tag = "dc:description"
		case "lastModifiedBy":
			tag = "cp:lastModifiedBy"
		default:
			continue
		}
		re := regexp.MustCompile(`(?s)<` + tag + `\b[^>]*>.*?</` + tag + `>`)
		elem := `<` + tag + `>` + safeText(v) + `</` + tag + `>`
		if re.MatchString(xml) {
			xml = re.ReplaceAllString(xml, elem)
		} else {
			xml = strings.Replace(xml, "</cp:coreProperties>", elem+"</cp:coreProperties>", 1)
		}
	}
	b.fileMap[path] = []byte(xml)
	if _, ok := b.fileMap[path]; ok {
		b.fileMap["[Content_Types].xml"] = insertOverrideInContentTypes(b.fileMap["[Content_Types].xml"],
			"/docProps/core.xml", "application/vnd.openxmlformats-package.core-properties+xml")
	}
	// 根 .rels 确保 core.xml 关系
	rootRels := "_rels/.rels"
	if rels := loadRels(b.fileMap, rootRels); rels != nil {
		has := false
		for _, r := range rels.Relationship {
			if r.Type == "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" {
				has = true
			}
		}
		if !has {
			rid := fmt.Sprintf("rId%d", getMaxRId(rels)+1)
			rels.Relationship = append(rels.Relationship, Relationship{
				ID:     rid,
				Type:   "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
				Target: "docProps/core.xml",
			})
			b.fileMap[rootRels] = serializeRelationships(rels)
		}
	}
	return nil
}

// ---------- 通用插入辅助 ----------

// insertRowElement 在 sheetData 内插入一个完整 <row ...>...</row> 元素（rowNum 行不存在时）。
// rowXML 为已清洗的 <row> 片段（调用方保证其内容已过 safeText/safeAttr）。
func insertRowElement(ws string, rowNum int, rowXML xmlFrag) string {
	rowRe := regexp.MustCompile(`(?s)<row\b[^>]*\br="` + strconv.Itoa(rowNum) + `"[^>]*>.*?</row>`)
	if rowRe.MatchString(ws) {
		return ws
	}
	sdOpen := strings.Index(ws, "<sheetData")
	sdClose := strings.LastIndex(ws, "</sheetData>")
	if sdOpen == -1 || sdClose == -1 {
		wi := strings.LastIndex(ws, "</worksheet>")
		if wi == -1 {
			return ws + string(rowXML)
		}
		return ws[:wi] + `<sheetData>` + string(rowXML) + `</sheetData>` + ws[wi:]
	}
	inner := ws[sdOpen:sdClose]
	rowOpenRe := regexp.MustCompile(`<row\b[^>]*\br="(\d+)"`)
	insertAt := sdClose
	for _, rm := range rowOpenRe.FindAllStringSubmatchIndex(inner, -1) {
		curNum, _ := strconv.Atoi(inner[rm[2]:rm[3]])
		if curNum > rowNum {
			insertAt = sdOpen + rm[0]
			break
		}
	}
	return ws[:insertAt] + string(rowXML) + ws[insertAt:]
}

// insertAfterSheetData 把 block 插入到 </sheetData> 之后（若没有则插在 </worksheet> 前）。
// 用于 mergeCells/autoFilter/conditionalFormatting/dataValidations/hyperlinks/tableParts/comments
// 等应位于 sheetData 之后的元素。
func insertAfterSheetData(ws string, block xmlFrag) string {
	// 优先插入到 </sheetData> 之后（标准顺序：sheetData 之后才是 mergeCells/filter 等）
	if i := strings.LastIndex(ws, "</sheetData>"); i != -1 {
		return ws[:i+len("</sheetData>")] + string(block) + ws[i+len("</sheetData>"):]
	}
	// 兜底：自闭合 <sheetData/> 形式，替换为显式闭合再加 block
	if i := strings.Index(ws, "<sheetData/>"); i != -1 {
		return ws[:i] + "<sheetData></sheetData>" + string(block) + ws[i+len("<sheetData/>"):]
	}
	if i := strings.LastIndex(ws, "</worksheet>"); i != -1 {
		return ws[:i] + string(block) + ws[i:]
	}
	return ws + string(block)
}
