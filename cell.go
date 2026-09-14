package excelgo

// cell.go 提供单元格读写能力。
//
// 实现方式说明：worksheet XML 结构复杂（含 drawing 引用、WPS 内嵌图片块、数据验证等），
// 用 encoding/xml 整体反序列化容易丢字段或破坏命名空间。因此这里采用「字符串/正则定位
// 目标单元格 → 精确保留其余内容 → 仅改写该 <c> 元素」的策略，最大程度不破坏原布局与样式。
//
// 单元格 <c> 的关键属性：
//   - r：引用（如 "A1"）
//   - s：样式索引（保留原值，新写单元格默认 0 即常规样式）
//   - t：类型（"s" 共享字符串 / "str" 公式字符串 / "b" 布尔 / "e" 错误 / 缺省为数值或 inline）
// 子元素：<f>公式</f> <v>值</v> <is><t>内联字符串</t></is>

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// CellType 单元格写入类型（参考 excelize 的细分 API）。
type CellType int

const (
	// CellTypeDefault 自动推断：数值/布尔/字符串按内容选最合适类型。
	CellTypeDefault CellType = iota
	// CellTypeNumeric 数值（默认样式 n，写 <v>）。
	CellTypeNumeric
	// CellTypeString 共享字符串（t="s"，写入 sharedStrings.xml）。
	CellTypeString
	// CellTypeInline 内联字符串（t="inlineStr"，写 <is><t>）。
	CellTypeInline
	// CellTypeBool 布尔（t="b"，写 <v>0|1</v>）。
	CellTypeBool
	// CellTypeFormula 公式（写 <f>公式</f><v>可选结果</v>），类型按结果推断。
	CellTypeFormula
)

// GetCell 读取 sheetRef 工作表中 cell（如 "A1"）的显示值（字符串）。
// 返回空字符串表示单元格为空。共享字符串（t="s"）会解析为真实文本；公式（含 <f>）返回
// 公式结果 <v>，若无可读结果则返回公式文本（以 "=" 开头）。
func GetCell(filename, sheetRef, cell string) (string, error) {
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return "", err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return "", err
	}
	file, _, _, _, err := locateSheetInMap(fileMap, wb, sheetRef)
	if err != nil {
		return "", err
	}
	ws := string(fileMap[file])
	return readCellValue(fileMap, ws, cell), nil
}

// SetCellValue 按 value 的类型自动选择写入格式（参考 excelize.SetCellValue）：
//   - string  → 共享字符串（t="s"）
//   - int/int64/float  → 数值（默认类型）
//   - bool    → 布尔（t="b"）
//   - time.Time → 序列号数值（TODO：暂按 float64 处理，需要时再补日期格式）
//
// 保留原单元格 s 样式索引；若单元格不存在则新建（s 默认 0）。
func SetCellValue(filename, sheetRef, cell string, value interface{}) error {
	switch v := value.(type) {
	case string:
		return SetCellStr(filename, sheetRef, cell, v)
	case bool:
		return SetCellBool(filename, sheetRef, cell, v)
	case int:
		return SetCellInt(filename, sheetRef, cell, v)
	case int64:
		return SetCellInt(filename, sheetRef, cell, int(v))
	case float64:
		return SetCellNumeric(filename, sheetRef, cell, v)
	case float32:
		return SetCellNumeric(filename, sheetRef, cell, float64(v))
	default:
		return fmt.Errorf("不支持的值类型: %T", value)
	}
}

// SetCellDefault 以最通用方式写入：数字按数值，其余按共享字符串。等价于 SetCellValue。
func SetCellDefault(filename, sheetRef, cell string, value interface{}) error {
	return SetCellValue(filename, sheetRef, cell, value)
}

// SetCellStr 写入共享字符串（t="s"）。需要先确保 sharedStrings.xml 存在并追加条目。
func SetCellStr(filename, sheetRef, cell, value string) error {
	return writeCell(filename, sheetRef, cell, CellTypeString, value, "")
}

// SetCellInline 写入内联字符串（t="inlineStr"，不依赖 sharedStrings.xml）。
func SetCellInline(filename, sheetRef, cell, value string) error {
	return writeCell(filename, sheetRef, cell, CellTypeInline, value, "")
}

// SetCellInt 写入整数数值。
func SetCellInt(filename, sheetRef, cell string, value int) error {
	return writeCell(filename, sheetRef, cell, CellTypeNumeric, strconv.Itoa(value), "")
}

// SetCellNumeric 写入浮点数值。
func SetCellNumeric(filename, sheetRef, cell string, value float64) error {
	return writeCell(filename, sheetRef, cell, CellTypeNumeric, strconv.FormatFloat(value, 'f', -1, 64), "")
}

// SetCellBool 写入布尔值（t="b"，<v>1|0</v>）。
func SetCellBool(filename, sheetRef, cell string, value bool) error {
	b := "0"
	if value {
		b = "1"
	}
	return writeCell(filename, sheetRef, cell, CellTypeBool, b, "")
}

// SetCellFormula 写入公式（写 <f>公式</f>，可选 result 作为预计算值写入 <v>）。
// 例：SetCellFormula(f, "Sheet1", "B1", "A1+A2", "3")。result 为空则只写公式。
func SetCellFormula(filename, sheetRef, cell, formula, result string) error {
	return writeCell(filename, sheetRef, cell, CellTypeFormula, formula, result)
}

// ---------- 内部实现 ----------

// writeCell 是统一的单元格写入入口（按文件读写）。
func writeCell(filename, sheetRef, cell string, ct CellType, value, result string) error {
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return err
	}
	file, _, _, _, err := locateSheetInMap(fileMap, wb, sheetRef)
	if err != nil {
		return err
	}
	if err := setCellInMap(fileMap, file, cell, ct, value, result); err != nil {
		return err
	}
	return writeMapToZip(filename, fileMap)
}

// setCellInMap 在已读入内存的 fileMap 上写入单个单元格（不读写磁盘）。
// 供 writeCell 与区域批量写入 SetRange 复用，避免每个单元格重复读盘。
// file 为 worksheet XML 在 fileMap 中的路径；ct/value/result 语义同 writeCell。
func setCellInMap(fileMap map[string][]byte, file, cell string, ct CellType, value, result string) error {
	ws := string(fileMap[file])

	// 共享字符串：转为索引
	finalValue := value
	tAttr := ""
	switch ct {
	case CellTypeString:
		idx, err := ensureSharedString(fileMap, value)
		if err != nil {
			return err
		}
		finalValue = strconv.Itoa(idx)
		tAttr = ` t="s"`
	case CellTypeInline:
		tAttr = ` t="inlineStr"`
	case CellTypeBool:
		tAttr = ` t="b"`
	case CellTypeFormula:
		tAttr = ` t="str"`
		// 公式文本不应包含前导 "="（OOXML <f> 内不含 "="，由应用程序补）；
		// 去掉调用方可能误带的前导 "="，避免 Excel/openpyxl 解析为 "=="。
		finalValue = strings.TrimPrefix(value, "=")
	}

	newCellXML := buildCellXML(cell, tAttr, ct, finalValue, result)
	ws = replaceOrInsertCell(ws, cell, newCellXML)
	fileMap[file] = []byte(ws)
	return nil
}

// ensureSharedString 确保 value 存在于 sharedStrings.xml，返回其索引（不存在则追加）。
func ensureSharedString(fileMap map[string][]byte, value string) (int, error) {
	const ssPath = "xl/sharedStrings.xml"
	// 读取或初始化
	var items []string
	count := 0
	if data, ok := fileMap[ssPath]; ok {
		items = extractSharedStrings(string(data))
		count = len(items)
	}
	// 去重：已存在则复用索引
	for i, s := range items {
		if s == value {
			return i, nil
		}
	}
	// 追加
	items = append(items, value)
	// 重建 sharedStrings.xml
	fileMap[ssPath] = []byte(buildSharedStrings(items))
	// 确保 Content_Types 有 Override
	fileMap["[Content_Types].xml"] = insertOverrideInContentTypes(
		fileMap["[Content_Types].xml"],
		"/xl/sharedStrings.xml",
		"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml",
	)
	return count, nil
}

// extractSharedStrings 从 sharedStrings.xml 提取所有 <t> 文本（顺序即索引）。
func extractSharedStrings(ssXML string) []string {
	// 匹配 <si>...<t>文本</t>...</si>
	siRe := regexp.MustCompile(`(?s)<si\b[^>]*>(.*?)</si>`)
	tRe := regexp.MustCompile(`(?s)<t\b[^>]*>(.*?)</t>`)
	var out []string
	for _, si := range siRe.FindAllStringSubmatch(ssXML, -1) {
		inner := si[1]
		tm := tRe.FindStringSubmatch(inner)
		if tm != nil {
			out = append(out, unescapeXML(tm[1]))
		} else {
			out = append(out, "")
		}
	}
	return out
}

// buildSharedStrings 重建整个 sharedStrings.xml（最小但合法）。
func buildSharedStrings(items []string) string {
	var b strings.Builder
	b.WriteString(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`)
	b.WriteString(`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="`)
	b.WriteString(strconv.Itoa(len(items)))
	b.WriteString(`" uniqueCount="`)
	b.WriteString(strconv.Itoa(len(items)))
	b.WriteString(`">`)
	for _, s := range items {
		b.WriteString(`<si><t xml:space="preserve">`)
		b.WriteString(escapeXML(s))
		b.WriteString(`</t></si>`)
	}
	b.WriteString(`</sst>`)
	return b.String()
}

// buildCellXML 构造 <c r="cell" [t="..."] [s="N"]> 元素。
// 对已有单元格，s 在 replaceOrInsertCell 中保留；此处 s 缺省（新建时默认 0 由调用方决定，
// 但为简洁：新建单元格不带 s，使用默认样式 0；若需保留调用方应传入 s）。
func buildCellXML(cell, tAttr string, ct CellType, value, result string) string {
	// 读取现有 s（若 ws 中已有该单元格，在 replaceOrInsertCell 处理；此处仅构造值部分）
	switch ct {
	case CellTypeString, CellTypeBool, CellTypeNumeric, CellTypeDefault:
		if value == "" {
			return `<c r="` + cell + `"` + tAttr + `/>`
		}
		return `<c r="` + cell + `"` + tAttr + `><v>` + escapeXML(value) + `</v></c>`
	case CellTypeInline:
		return `<c r="` + cell + `"` + tAttr + `><is><t xml:space="preserve">` + escapeXML(value) + `</t></is></c>`
	case CellTypeFormula:
		inner := `<f>` + escapeXML(value) + `</f>`
		if result != "" {
			inner += `<v>` + escapeXML(result) + `</v>`
		}
		return `<c r="` + cell + `"` + tAttr + `>` + inner + `</c>`
	}
	return `<c r="` + cell + `"` + tAttr + `/>`
}

// replaceOrInsertCell 把 worksheet 中 cell 对应的 <c> 替换为 newXML；若不存在则插入。
// 规则：保留原单元格的 s 样式索引；插入时严格放进对应的 <row r="行号"> 内（OOXML 要求
// <c> 必须位于 <row> 中），不破坏任何其他行/单元格/布局。
func replaceOrInsertCell(ws, cell, newXML string) string {
	// 查找现有单元格
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + regexp.QuoteMeta(cell) + `"[^>]*/?>(?:.*?</c>)?`)
	loc := re.FindStringIndex(ws)
	if loc != nil {
		oldCell := ws[loc[0]:loc[1]]
		sAttr := extractSAttr(oldCell)
		newXML = injectSAttr(newXML, sAttr)
		return ws[:loc[0]] + newXML + ws[loc[1]:]
	}
	// 不存在：注入默认 s（0）
	newXML = injectSAttr(newXML, ` s="0"`)
	_, rowNum, err := parseCellRef(cell)
	if err != nil {
		rowNum = 0
	}
	return insertCellIntoRow(ws, rowNum+1, newXML)
}

// extractSAttr 从 <c> 标签提取 s="N" 属性（含前导空格），无则返回 ""。
func extractSAttr(c string) string {
	re := regexp.MustCompile(`\s+s="\d+"`)
	m := re.FindString(c)
	if m == "" {
		return ` s="0"`
	}
	return m
}

// injectSAttr 把一个 s 属性注入到 newXML 的 <c ...> 中（在 r 属性之后、其他属性之前）。
// 若 newXML 已含 s 则替换；若不含则插入到 r="..." 之后。
func injectSAttr(newXML, sAttr string) string {
	if strings.Contains(newXML, ` s="`) {
		// 替换已有 s
		re := regexp.MustCompile(`\s+s="\d+"`)
		return re.ReplaceAllString(newXML, sAttr)
	}
	// 插入到 r="..." 之后
	re := regexp.MustCompile(`(<c\b[^>]*\br="[^"]*")`)
	return re.ReplaceAllString(newXML, `${1}`+sAttr)
}

// insertCellIntoRow 把 newCell 插入 worksheet 中行号 rowNum 对应的 <row> 内。
// 若该 <row> 已存在，插入到其 </row> 前；若不存在，新建 <row r="rowNum"> 并插入到
// sheetData 内、按行序位于第一个更大行号之前（保持行序合法），无更大行则置于末尾。
func insertCellIntoRow(ws string, rowNum int, newCell string) string {
	// 查找 <row r="rowNum" ...> ... </row>
	rowRe := regexp.MustCompile(`(?s)<row\b[^>]*\br="` + strconv.Itoa(rowNum) + `"[^>]*>(.*?)</row>`)
	if m := rowRe.FindStringSubmatchIndex(ws); m != nil {
		// 在该 row 的 </row> 前插入
		closeTag := strings.LastIndex(ws[m[0]:m[1]], "</row>")
		insertAt := m[0] + closeTag
		return ws[:insertAt] + newCell + ws[insertAt:]
	}
	// 无该行：查找 sheetData 范围内第一个 r > rowNum 的 <row>，插在其前；
	// 否则插在最后一个 </row> 之后、</sheetData> 之前。
	sdOpen := strings.Index(ws, "<sheetData")
	sdClose := strings.LastIndex(ws, "</sheetData>")
	if sdOpen == -1 || sdClose == -1 {
		// 无 sheetData：创建
		wi := strings.LastIndex(ws, "</worksheet>")
		if wi == -1 {
			return ws + `<sheetData><row r="` + strconv.Itoa(rowNum) + `">` + newCell + `</row></sheetData>`
		}
		return ws[:wi] + `<sheetData><row r="` + strconv.Itoa(rowNum) + `">` + newCell + `</row></sheetData>` + ws[wi:]
	}
	inner := ws[sdOpen:sdClose]
	rowOpenRe := regexp.MustCompile(`<row\b[^>]*\br="(\d+)"`)
	insertAt := sdClose // 默认放 sheetData 末尾
	prevRowEnd := sdOpen
	foundAnyRow := false
	for _, rm := range rowOpenRe.FindAllStringSubmatchIndex(inner, -1) {
		foundAnyRow = true
		curNum, _ := strconv.Atoi(inner[rm[2]:rm[3]])
		if curNum > rowNum {
			// 插到这个 row 之前
			insertAt = sdOpen + rm[0]
			break
		}
		// 记录当前 row 结束位置（用于"放最后"）
		rowClose := strings.Index(inner[rm[1]:], "</row>")
		if rowClose != -1 {
			prevRowEnd = sdOpen + rm[1] + rowClose + len("</row>")
		}
	}
	if insertAt == sdClose {
		if foundAnyRow {
			insertAt = prevRowEnd
		} else {
			// sheetData 为空：把新行插入到 <sheetData ...> 起始标签之后（而非之前），
			// 避免出现「行在 <sheetData> 之前 / 之外」的非法结构。
			openEnd := strings.Index(inner, ">")
			if openEnd == -1 {
				openEnd = 0
			}
			insertAt = sdOpen + openEnd + 1
		}
	}
	return ws[:insertAt] + `<row r="` + strconv.Itoa(rowNum) + `">` + newCell + `</row>` + ws[insertAt:]
}

// readCellValue 从 worksheet XML 读取 cell 的显示值：
//   - t="s"：从 sharedStrings.xml 解析真实文本（fileMap 提供）；
//   - 含 <f> 公式：返回 <v> 结果，无结果则返回公式文本（"=" 前缀）；
//   - 其余：返回 <v> 或内联 <t>。
func readCellValue(fileMap map[string][]byte, ws, cell string) string {
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + regexp.QuoteMeta(cell) + `"[^>]*/?>(.*?)</c>`)
	m := re.FindStringSubmatch(ws)
	if m == nil {
		re2 := regexp.MustCompile(`<c\b[^>]*\br="` + regexp.QuoteMeta(cell) + `"[^>]*/>`)
		if re2.MatchString(ws) {
			return ""
		}
		return ""
	}
	cellTag := extractCellTag(ws, cell)
	inner := m[1]
	tAttr := attrOf(cellTag, "t")

	// 公式优先
	if f := regexp.MustCompile(`(?s)<f>(.*?)</f>`).FindStringSubmatch(inner); f != nil {
		if v := regexp.MustCompile(`(?s)<v>(.*?)</v>`).FindStringSubmatch(inner); v != nil {
			return v[1]
		}
		return "=" + f[1]
	}

	switch tAttr {
	case "s":
		if v := regexp.MustCompile(`(?s)<v>(.*?)</v>`).FindStringSubmatch(inner); v != nil {
			idx, err := strconv.Atoi(v[1])
			if err == nil {
				if ss, ok := fileMap["xl/sharedStrings.xml"]; ok {
					items := extractSharedStrings(string(ss))
					if idx >= 0 && idx < len(items) {
						return items[idx]
					}
				}
			}
		}
		return ""
	case "inlineStr":
		if t := regexp.MustCompile(`(?s)<t\b[^>]*>(.*?)</t>`).FindStringSubmatch(inner); t != nil {
			return unescapeXML(t[1])
		}
		return ""
	default:
		if v := regexp.MustCompile(`(?s)<v>(.*?)</v>`).FindStringSubmatch(inner); v != nil {
			return v[1]
		}
		return ""
	}
}

// extractCellTag 提取 worksheet 中指定单元格的 <c ...> 起始标签原文（含属性）。
func extractCellTag(ws, cell string) string {
	re := regexp.MustCompile(`<c\b[^>]*\br="` + regexp.QuoteMeta(cell) + `"[^>]*>`)
	m := re.FindString(ws)
	return m
}

// escapeXML / unescapeXML 处理单元格文本中的 XML 特殊字符。
func escapeXML(s string) string {
	r := strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;", `"`, "&quot;")
	return r.Replace(s)
}

func unescapeXML(s string) string {
	r := strings.NewReplacer("&quot;", `"`, "&lt;", "<", "&gt;", ">", "&amp;", "&")
	return r.Replace(s)
}
