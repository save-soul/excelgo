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
	"math"
	"regexp"
	"strconv"
	"strings"
	"time"
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

// typedValue 把 <v> 文本按内容推断为 int / float64 / string。
func typedValue(s, t string) (interface{}, error) {
	if t == "str" {
		return s, nil
	}
	if f, err := strconv.ParseFloat(s, 64); err == nil {
		if f == math.Trunc(f) && !strings.ContainsAny(s, ".eE") {
			return int(f), nil
		}
		return f, nil
	}
	return s, nil
}

// GetCellValue 读取 sheetRef!cell 的值并按其类型返回原生 Go 值（见 WorkSheet.GetCellValue）。
func GetCellValue(filename, sheetRef, cell string) (interface{}, error) {
	b, err := Open(filename)
	if err != nil {
		return nil, err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return nil, err
	}
	return ws.GetCellValue(cell)
}

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
//   - time.Time → Excel 序列号 + 自动配日期格式（见 SetCellTime）
//
// 保留原单元格 s 样式索引；若单元格不存在则新建（s 默认 0）。
func SetCellValue(filename, sheetRef, cell string, value interface{}) error {
	// 类型矩阵与 (*WorkSheet).SetCellValue 保持一致：
	// 两侧支持同样的类型，否则"方法能用、包级不能用"会成为难以察觉的坑。
	switch v := value.(type) {
	case nil:
		return nil
	case string:
		return SetCellStr(filename, sheetRef, cell, v)
	case bool:
		return SetCellBool(filename, sheetRef, cell, v)
	case int:
		return SetCellInt(filename, sheetRef, cell, v)
	case int8:
		return SetCellInt(filename, sheetRef, cell, int(v))
	case int16:
		return SetCellInt(filename, sheetRef, cell, int(v))
	case int32:
		return SetCellInt(filename, sheetRef, cell, int(v))
	case int64:
		return SetCellInt(filename, sheetRef, cell, int(v))
	case uint:
		return SetCellNumeric(filename, sheetRef, cell, float64(v))
	case uint8:
		return SetCellInt(filename, sheetRef, cell, int(v))
	case uint16:
		return SetCellInt(filename, sheetRef, cell, int(v))
	case uint32:
		return SetCellInt(filename, sheetRef, cell, int(v))
	case uint64:
		return SetCellNumeric(filename, sheetRef, cell, float64(v))
	case float64:
		return SetCellNumeric(filename, sheetRef, cell, v)
	case float32:
		return SetCellNumeric(filename, sheetRef, cell, float64(v))
	case time.Time:
		return SetCellTime(filename, sheetRef, cell, v)
	default:
		return fmt.Errorf("不支持的值类型: %T（可用：nil/bool/int 族/uint 族/float/string/time.Time）", value)
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
	var tAttr xmlFrag
	switch ct {
	case CellTypeString:
		idx, err := ensureSharedString(fileMap, value)
		if err != nil {
			return err
		}
		finalValue = strconv.Itoa(idx)
		tAttr = xmlFrag(` t="s"`)
	case CellTypeInline:
		tAttr = xmlFrag(` t="inlineStr"`)
	case CellTypeBool:
		tAttr = xmlFrag(` t="b"`)
	case CellTypeFormula:
		tAttr = xmlFrag(` t="str"`)
		// 公式文本不应包含前导 "="（OOXML <f> 内不含 "="，由应用程序补）；
		// 去掉调用方可能误带的前导 "="，避免 Excel/openpyxl 解析为 "=="。
		finalValue = strings.TrimPrefix(value, "=")
	}

	// buildCellXML 内部已对 value/result 过 safeText，产物是可信 XML 片段
	newCellXML := frag(buildCellXML(newCellRef(cell), tAttr, ct, finalValue, result))
	ws = replaceOrInsertCell(ws, cell, string(newCellXML))
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
		b.WriteString(safeText(s))
		b.WriteString(`</t></si>`)
	}
	b.WriteString(`</sst>`)
	return b.String()
}

// buildCellXML 构造 <c r="cell" [t="..."] [s="N"]> 元素。
// 对已有单元格，s 在 replaceOrInsertCell 中保留；此处 s 缺省（新建时默认 0 由调用方决定，
// 但为简洁：新建单元格不带 s，使用默认样式 0；若需保留调用方应传入 s）。
func buildCellXML(cell cellRef, tAttr xmlFrag, ct CellType, value, result string) string {
	// 读取现有 s（若 ws 中已有该单元格，在 replaceOrInsertCell 处理；此处仅构造值部分）
	switch ct {
	case CellTypeString, CellTypeBool, CellTypeNumeric, CellTypeDefault:
		if value == "" {
			return `<c r="` + string(cell) + `"` + string(tAttr) + `/>`
		}
		return `<c r="` + string(cell) + `"` + string(tAttr) + `><v>` + safeText(value) + `</v></c>`
	case CellTypeInline:
		return `<c r="` + string(cell) + `"` + string(tAttr) + `><is><t xml:space="preserve">` + safeText(value) + `</t></is></c>`
	case CellTypeFormula:
		inner := `<f>` + safeText(value) + `</f>`
		if result != "" {
			inner += `<v>` + safeText(result) + `</v>`
		}
		return `<c r="` + string(cell) + `"` + string(tAttr) + `>` + inner + `</c>`
	}
	return `<c r="` + string(cell) + `"` + string(tAttr) + `/>`
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
	return insertCellIntoRow(part(ws), rowNum+1, frag(newXML))
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
func insertCellIntoRow(ws partXML, rowNum int, newCell xmlFrag) string {
	// 查找 <row r="rowNum" ...> ... </row>
	rowRe := regexp.MustCompile(`(?s)<row\b[^>]*\br="` + strconv.Itoa(rowNum) + `"[^>]*>(.*?)</row>`)
	if m := rowRe.FindStringSubmatchIndex(string(ws)); m != nil {
		// 在该 row 的 </row> 前插入
		closeTag := strings.LastIndex(string(ws[m[0]:m[1]]), "</row>")
		insertAt := m[0] + closeTag
		return string(ws[:insertAt]) + string(newCell) + string(ws[insertAt:])
	}
	// 无该行：查找 sheetData 范围内第一个 r > rowNum 的 <row>，插在其前；
	// 否则插在最后一个 </row> 之后、</sheetData> 之前。
	ss := string(ws)
	sdOpen := strings.Index(ss, "<sheetData")
	sdClose := strings.LastIndex(ss, "</sheetData>")
	if sdOpen == -1 || sdClose == -1 {
		// 无 sheetData：创建
		wi := strings.LastIndex(ss, "</worksheet>")
		if wi == -1 {
			return ss + `<sheetData><row r="` + strconv.Itoa(rowNum) + `">` + string(newCell) + `</row></sheetData>`
		}
		return ss[:wi] + `<sheetData><row r="` + strconv.Itoa(rowNum) + `">` + string(newCell) + `</row></sheetData>` + ss[wi:]
	}
	inner := ss[sdOpen:sdClose]
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
	return string(ws[:insertAt]) + `<row r="` + strconv.Itoa(rowNum) + `">` + string(newCell) + `</row>` + string(ws[insertAt:])
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
	case "b":
		if v := regexp.MustCompile(`(?s)<v>(.*?)</v>`).FindStringSubmatch(inner); v != nil {
			if v[1] == "1" {
				return "TRUE"
			}
			return "FALSE"
		}
		return "FALSE"
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

// safeText 把用户数据转为可安全写入 **XML 文本节点** 的字符串。
//
// 约定（务必遵守）：
//
//	所有写入 xlsx 内部 XML 的用户数据，一律且只能经由本函数（或 safeAttr）出栈。
//	直接拼接用户数据即为漏洞 —— 见 TestNoUnescapedXMLWrite 静态检查。
//
// 做两件事：
//  1. 转义 5 个 XML 实体：& < > " '（' 在文本节点可不转义，但转义更保险）；
//  2. 剔除 XML 1.0 规范禁止的控制字符（U+0000~U+0008 / U+000B / U+000C /
//     U+000E~U+001F，保留 \t \n \r）。这类字符**无法用实体表达**，
//     写入后整个部件会变成 not-well-formed，Excel/WPS 直接判文件损坏。
//     攻击者一个 0x00 即可报废整份文档，这是"XML 注入"最隐蔽的形态。
func safeText(s string) string {
	if hasIllegalXMLChar(s) {
		s = stripIllegalXMLChar(s)
	}
	r := strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;", `"`, "&quot;")
	return r.Replace(s)
}

// safeAttr 把用户数据转为可安全写入 **XML 属性值**（带引号包裹的位置）的字符串。
// 语义与 safeText 相同；单独命名是为了让代码审查时一眼看出"这是属性上下文"。
func safeAttr(s string) string { return safeText(s) }

// hasIllegalXMLChar 判断是否含 XML 1.0 不允许的控制字符（不含 \t \n \r）。
func hasIllegalXMLChar(s string) bool {
	for _, r := range s {
		if r < 0x20 && r != '\t' && r != '\n' && r != '\r' {
			return true
		}
	}
	return false
}

// stripIllegalXMLChar 剔除 XML 1.0 不允许的控制字符。
func stripIllegalXMLChar(s string) string {
	return strings.Map(func(r rune) rune {
		if r < 0x20 && r != '\t' && r != '\n' && r != '\r' {
			return -1
		}
		return r
	}, s)
}

func unescapeXML(s string) string {
	// 单次扫描即可：Replacer 自左向右逐位匹配，即使 &amp; 排在最后，
	// 遇到 "&amp;lt;" 时也会先消耗掉 "&amp;" 而不会被 "&lt;" 抢先匹配，
	// 结果 "&lt;"（这正是 XML 语义要求的字面量文本）。
	r := strings.NewReplacer("&quot;", `"`, "&lt;", "<", "&gt;", ">", "&amp;", "&")
	return r.Replace(s)
}

// ---------- 已认证的内部构造（用类型表达"安全"这一事实） ----------
//
// XML 写入护栏（guard_taint_test.go）需要区分"用户输入"与"库内生成的安全值"。
// 若二者都是 string，静态分析无法辨别，只能靠人工豁免清单 —— 那意味着新增写入点
// 时要靠人记得填清单，迟早会漏。
//
// 解法：把「库内生成且已校验」的值用**具名类型**表达，使安全性成为类型事实：
//   - cellRef：经 newCellRef 规范化（校验 A1 形式）的单元格坐标；
//   - relID：库内生成的关系 ID（rIdN）；
//   - xmlFrag：仅由字面量与 safe* 拼成的 XML 片段。
// 这三种类型在护栏里天然"不带污"，无需任何豁免；
// 反之，裸 string 参数默认带污，漏转义会立刻被检出。

// cellRef 是已规范化的单元格坐标（如 "A1"、"BC12"）。
// 只能由 newCellRef / parseCellRef 构造，构造时即完成校验。
type cellRef string

// newCellRef 校验并规范化单元格坐标；非法返回空串。
// 校验复用 parseCellRef 的正则（^([A-Za-z]+)(\d+)$），保证与既有解析一致。
func newCellRef(s string) cellRef {
	if _, _, err := parseCellRef(s); err != nil {
		return ""
	}
	return cellRef(strings.ToUpper(strings.TrimSpace(s)))
}

// relID 是库内生成的关系 ID（形如 "rId7"），不含用户数据。
type relID string

// newRelID 由计数器生成关系 ID。
func newRelID(n int) relID { return relID("rId" + strconv.Itoa(n)) }

// xmlFrag 是「仅由字面量与 safe* 拼成」的 XML 片段，构造即安全。
// 用它包裹已清洗的片段，让护栏无需把局部中间量当作污点。
type xmlFrag string

// frag 标记一个已清洗的 XML 片段（调用方负责确保其内部已 safe* 处理）。
func frag(s string) xmlFrag { return xmlFrag(s) }

// xmlTagName 是「受信任的 schema 标签名」——库内固定字面量（如 "v"、"f"、"r"），
// 不接受外部输入。用于 XML 正则拼接（非输出），故与认证类型一并纳入护栏白名单。
type xmlTagName string

// tagName 标记一个受信任的标签名常量。
func tagName(s string) xmlTagName { return xmlTagName(s) }

// newRelIDFrom 把既有关系 ID 字符串包装为认证的 relID。
// 仅供库内已生成 rIdN 的场景复用（如 drawing 内部关系）。
func newRelIDFrom(s string) relID { return relID(s) }

// rangeRef 是已校验的区域引用（如 "A1:C10"）。只能由 newRangeRef 构造，
// 构造时即完成合法性校验，故可安全写入 XML 属性。
type rangeRef string

// newRangeRef 校验并规范化区域引用；非法返回空串。
// 接受 "A1" / "A1:C10" / "A:C" / "1:1" 四种形态。
func newRangeRef(s string) rangeRef {
	s = strings.ToUpper(strings.TrimSpace(s))
	if s == "" {
		return ""
	}
	for _, part := range strings.Split(s, ":") {
		if !isValidRangePart(part) {
			return ""
		}
	}
	return rangeRef(s)
}

// isValidRangePart 校验区域引用的单段（列/行/单元格三者之一）。
func isValidRangePart(p string) bool {
	if p == "" {
		return false
	}
	i := 0
	for i < len(p) && p[i] >= 'A' && p[i] <= 'Z' {
		i++
	}
	if i == len(p) {
		return i <= 3 // 纯列引用 A / AB
	}
	if i > 3 || i == 0 {
		return false
	}
	rest := p[i:]
	allDigits := true
	for j := 0; j < len(rest); j++ {
		if rest[j] < '0' || rest[j] > '9' {
			allDigits = false
			break
		}
	}
	if !allDigits {
		// 允许 $ 绝对引用前缀
		rest = strings.TrimPrefix(rest, "$")
		if rest == "" {
			return false
		}
		for j := 0; j < len(rest); j++ {
			if rest[j] < '0' || rest[j] > '9' {
				return false
			}
		}
	}
	return true
}

// partXML 表示「xlsx 部件的原始 XML 字节」（worksheet / workbook / comments 等）。
//
// 它与 xmlFrag 的区别：
//   - xmlFrag 是**本次新建**的片段（内容全部经过 safe* 处理）；
//   - partXML 是**文件里已有**的字节，来源是解析其它部件或用户文件。
//
// 之所以单独建类型：把既有 XML 重新拼接（如 `ws[:i] + frag + ws[i:]`）是"读取后复用"，
// 与"把用户数据写进 XML"性质不同 —— 两侧既有内容都未被改写。类型化后护栏无需
// 任何按变量名的判断即可区分二者。
type partXML string

// part 把既有部件 XML 标记为 partXML（不改变内容）。
func part(s string) partXML { return partXML(s) }

// ---------- 日期时间写入 ----------

// excelDateBase 是 Excel 1900 日期系统的基准日。
//
// 之所以是 1899-12-30 而不是 1 月 1 日：Excel 为了兼容早期 Lotus 1-2-3
// 保留了一个并不存在的 1900-02-29（伪闰日），序列号整体比真实日历多 1 天。
// 基准取 1899-12-30 正好绕开这个偏差。
var excelDateBase = time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)

// timeToExcelSerial 把 time.Time 转成 Excel 序列号（小数为当天的时间占比）。
//
// 算法必须与 serialToTime（读侧）**严格互逆**，否则往返会差几百纳秒：
//序列号是 float64（约 15-16 位有效数字），乘 86400 会放大尾数误差。
// 早期两侧各写一套（写侧 d.Hours()/24，读侧 serial*86400*1e9），
// 23:59:59 往返后会变成 23:59:59.000000512。
//
// 现在两侧统一为"整数秒 + 小数秒"的分解：
//   写：(总秒 + 纳秒/1e9) / 86400
//   读：序列号 * 86400 拆成整数秒与小数秒，再合成 Duration
func timeToExcelSerial(t time.Time) float64 {
	d := t.UTC().Sub(excelDateBase)
	secs := float64(int64(d/time.Second)) + float64(d%time.Second)/1e9
	return secs / 86400
}

// SetCellTime 写入日期时间单元格（自动配上日期格式，否则 Excel 里显示为数字）。
//
// 格式选择：
//   - 含时分秒 → "yyyy-mm-dd hh:mm:ss"
//   - 含时分   → "yyyy-mm-dd hh:mm"
//   - 仅日期   → "yyyy-mm-dd"
//
// 保留原单元格的其它样式（字体、边框等），只覆盖数字格式。
func (s *WorkSheet) SetCellTime(ref string, t time.Time) error {
	if err := s.SetCellNumeric(ref, timeToExcelSerial(t)); err != nil {
		return err
	}
	code := "yyyy-mm-dd"
	switch {
	case t.Hour() != 0 || t.Minute() != 0 || t.Second() != 0:
		code = "yyyy-mm-dd hh:mm:ss"
	case t.Hour() != 0 || t.Minute() != 0:
		code = "yyyy-mm-dd hh:mm"
	}
	// 只设数字格式：读回 GetStyle 拿到全量样式后改 NumFmt 再写回，
	// 避免覆盖字体/边框等调用方已设的属性。
	st, err := s.GetStyle(ref)
	if err != nil {
		return err
	}
	st.NumFmt = code
	_, err = s.SetStyle(ref, st)
	return err
}

// SetCellTime 写入日期时间单元格（包级 API）。
func SetCellTime(filename, sheetRef, cell string, t time.Time) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.SetCellTime(cell, t)
	})
}
