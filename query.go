package excelgo

// query.go 提供工作表查询与整表读取能力。
//
// 与复制/合并/单元格读写一脉相承：所有操作基于 readZipToMap → 修改 → writeMapToZip，
// 且尽量用「字符串/正则精确定位」策略，避免 encoding/xml 整体反序列化破坏原布局。

import (
	"regexp"
	"sort"
	"strconv"
	"strings"
)

// GetSheetList 返回工作簿中所有工作表名称（按 workbook.xml 中的顺序）。
func GetSheetList(filename string) ([]string, error) {
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return nil, err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return nil, err
	}
	names := make([]string, 0, len(wb.Sheets.Sheet))
	for _, s := range wb.Sheets.Sheet {
		names = append(names, s.Name)
	}
	return names, nil
}

// GetSheetIndex 返回 sheetRef（名称或 1 基索引）对应的 1 基序号；不存在返回 0 与错误。
func GetSheetIndex(filename, sheetRef string) (int, error) {
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return 0, err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return 0, err
	}
	_, _, idx, _, err := locateSheetInMap(fileMap, wb, sheetRef)
	return idx, err
}

// SheetExists 判断 sheetRef 是否存在于工作簿。
func SheetExists(filename, sheetRef string) (bool, error) {
	_, err := GetSheetIndex(filename, sheetRef)
	if err != nil {
		if strings.Contains(err.Error(), "未找到") {
			return false, nil
		}
		return false, err
	}
	return true, nil
}

// GetRows 返回 sheetRef 工作表中所有行的单元格值（二维切片，按行序）。
// 每行是 []string，空单元格以 "" 表示。共享字符串解析为文本；公式取 <v> 结果。
// 仅含数据的行（中间空行会被保留为 nil 或空切片，由调用方判断）。
func GetRows(filename, sheetRef string) ([][]string, error) {
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return nil, err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return nil, err
	}
	file, _, _, _, err := locateSheetInMap(fileMap, wb, sheetRef)
	if err != nil {
		return nil, err
	}
	ws := string(fileMap[file])
	return extractRows(ws, fileMap), nil
}

// GetSheetData 等价于 GetRows，命名参考 excelize（返回 [][]string）。
func GetSheetData(filename, sheetRef string) ([][]string, error) {
	return GetRows(filename, sheetRef)
}

// extractRows 从 worksheet XML 提取所有行的单元格值，转为二维 [][]string。
// 行按 r 属性升序；列按 A,B,C... 顺序，中间空列以 "" 占位。
func extractRows(ws string, fileMap map[string][]byte) [][]string {
	// 取 <sheetData> 内所有 <row>
	sdRe := regexp.MustCompile(`(?s)<sheetData\b[^>]*>(.*?)</sheetData>`)
	sd := sdRe.FindStringSubmatch(ws)
	if sd == nil {
		return nil
	}
	inner := sd[1]
	// 同时匹配有内容的 <row>...</row> 与自闭合空行 <row r="N"/>
	rowRe := regexp.MustCompile(`(?s)<row\b[^>]*\br="(\d+)"[^>]*?(?:/>|>(.*?)</row>)`)
	rowMatches := rowRe.FindAllStringSubmatch(inner, -1)

	type rowData struct {
		rowNum int
		cells  map[int]string // colIdx(0基) -> value
	}
	var rows []rowData
	maxCol := 0
	for _, rm := range rowMatches {
		rowNum, _ := strconv.Atoi(rm[1])
		rowInner := rm[2]
		cells := map[int]string{}
		// 匹配整个 <c ...>...</c>（含属性，便于读取 t 类型）
		cRe := regexp.MustCompile(`(?s)<c\b[^>]*\br="([A-Z]+\d+)"[^>]*?(?:/>|>(.*?)</c>)`)
		for _, cm := range cRe.FindAllStringSubmatch(rowInner, -1) {
			ref := cm[1]
			fullCell := cm[0] // 整个 <c> 元素
			col, _, err := parseCellRef(ref)
			if err != nil {
				continue
			}
			val := cellDisplayValue(fullCell, fileMap)
			cells[col] = val
			if col > maxCol {
				maxCol = col
			}
		}
		rows = append(rows, rowData{rowNum: rowNum, cells: cells})
	}
	if len(rows) == 0 {
		return nil
	}
	// 按 rowNum 升序
	sort.Slice(rows, func(i, j int) bool { return rows[i].rowNum < rows[j].rowNum })
	out := make([][]string, len(rows))
	for i, r := range rows {
		line := make([]string, maxCol+1)
		for c := 0; c <= maxCol; c++ {
			line[c] = r.cells[c]
		}
		out[i] = line
	}
	return out
}

// cellDisplayValue 从单个 <c> 内部 XML 提取显示值。
// 处理：t="s" 共享字符串、t="inlineStr" 内联、t="b" 布尔、t="str" 公式字符串、默认数值、公式结果。
func cellDisplayValue(cellInner string, fileMap map[string][]byte) string {
	tAttr := attrOf(cellInner, "t")
	switch tAttr {
	case "s":
		// 共享字符串索引
		if v := firstTag(cellInner, "v"); v != "" {
			if idx, err := strconv.Atoi(v); err == nil {
				return sharedStringAt(fileMap, idx)
			}
		}
		return ""
	case "inlineStr":
		if is := regexp.MustCompile(`(?s)<is\b[^>]*>(.*?)</is>`).FindStringSubmatch(cellInner); is != nil {
			if t := regexp.MustCompile(`(?s)<t\b[^>]*>(.*?)</t>`).FindStringSubmatch(is[1]); t != nil {
				return unescapeXML(t[1])
			}
		}
		return ""
	case "b":
		if v := firstTag(cellInner, "v"); v == "1" {
			return "TRUE"
		}
		return "FALSE"
	case "str":
		// 公式字符串结果
		if v := firstTag(cellInner, "v"); v != "" {
			return v
		}
		// 退化：无 v 返回公式文本
		if f := firstTag(cellInner, "f"); f != "" {
			return "=" + f
		}
		return ""
	default:
		// 数值或公式结果
		if v := firstTag(cellInner, "v"); v != "" {
			return v
		}
		// 有 <f> 但无缓存结果（如 <v /> 自闭合）：返回公式文本
		if f := firstTag(cellInner, "f"); f != "" {
			return "=" + f
		}
		return ""
	}
}

// sharedStringAt 读取 sharedStrings.xml 中第 idx 条（不存在返回 ""）。
func sharedStringAt(fileMap map[string][]byte, idx int) string {
	data, ok := fileMap["xl/sharedStrings.xml"]
	if !ok {
		return ""
	}
	items := extractSharedStrings(string(data))
	if idx >= 0 && idx < len(items) {
		return items[idx]
	}
	return ""
}

// firstTag 提取第一个 <tag>...</tag> 的文本内容（若标签自闭合返回 ""）。
func firstTag(s, tag string) string {
	re := regexp.MustCompile(`(?s)<` + tag + `\b[^>]*>(.*?)</` + tag + `>`)
	if m := re.FindStringSubmatch(s); m != nil {
		return strings.TrimSpace(m[1])
	}
	return ""
}
