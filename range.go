package excelgo

// range.go 提供按矩形区域读取与写入单元格的能力。
//
// 与全库一致：基于 readZipToMap → 修改 → writeMapToZip，且区域写入复用 setCellInMap
// （在内存 map 上逐单元格改写，最终一次性写回），不重复读盘；写入保留原单元格 s 样式索引。

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// CellFormula 用于区域写入时表达一个公式单元格（Formula 为公式文本，Result 为可选预计算结果）。
type CellFormula struct {
	Formula string
	Result  string
}

// parseRangeRef 解析 "A1" 或 "A1:C10" 形式的范围，返回 0 基列号与行号（闭区间）。
// 单列/单行（如 "B2"）等价于起止相同的 1x1 区域。
func parseRangeRef(ref string) (c1, r1, c2, r2 int, err error) {
	ref = strings.TrimSpace(ref)
	parts := strings.Split(ref, ":")
	if len(parts) == 1 {
		c, r, e := parseCellRef(parts[0])
		if e != nil {
			return 0, 0, 0, 0, e
		}
		return c, r, c, r, nil
	}
	if len(parts) != 2 {
		return 0, 0, 0, 0, fmt.Errorf("无效的区域引用: %q", ref)
	}
	a, b, e := parseCellRef(parts[0])
	if e != nil {
		return 0, 0, 0, 0, e
	}
	c, d, e := parseCellRef(parts[1])
	if e != nil {
		return 0, 0, 0, 0, e
	}
	if a > c || b > d {
		return 0, 0, 0, 0, fmt.Errorf("区域引用 %q 起止顺序非法", ref)
	}
	return a, b, c, d, nil
}

// GetRange 读取 sheetRef 工作表中 rangeRef（如 "A1:C10"）矩形区域的单元格值。
// 返回二维切片，行/列与区域一一对应，空单元格以 "" 表示。公式取 <v> 结果。
func GetRange(filename, sheetRef, rangeRef string) ([][]string, error) {
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
	c1, r1, c2, r2, err := parseRangeRef(rangeRef)
	if err != nil {
		return nil, err
	}
	return extractRangeRows(ws, fileMap, c1, r1, c2, r2), nil
}

// SetRange 把 values 二维切片写入 sheetRef 工作表的 rangeRef 区域（左上角对齐）。
// 支持的值类型：string（共享字符串）/ int / float64（数值）/ bool / nil（跳过，保留原值）/
// CellFormula（写入公式，Result 为可选预计算结果）。保留被覆盖单元格的原 s 样式索引。
func SetRange(filename, sheetRef, rangeRef string, values [][]interface{}) error {
	if len(values) == 0 {
		return nil
	}
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
	sc := &sharedStringsCache{}
	if err := setRangeInMap(fileMap, file, rangeRef, values, sc); err != nil {
		return err
	}
	return writeMapToZip(filename, fileMap)
}

// setRangeInMap 在内存 fileMap 上把 values 写入 file 工作表的 rangeRef（不复写磁盘）。
func setRangeInMap(fileMap map[string][]byte, file, rangeRef string, values [][]interface{}, sc *sharedStringsCache) error {
	if len(values) == 0 {
		return nil
	}
	c1, r1, c2, r2, err := parseRangeRef(rangeRef)
	if err != nil {
		return err
	}
	cols := c2 - c1 + 1
	rows := r2 - r1 + 1
	if strings.Contains(rangeRef, ":") {
		// 显式区域：校验数据不越界
		if len(values) > rows {
			return fmt.Errorf("数据行数 %d 超出区域行数 %d", len(values), rows)
		}
		for i, rowVals := range values {
			if len(rowVals) > cols {
				return fmt.Errorf("第 %d 行数据列数 %d 超出区域列数 %d", i+1, len(rowVals), cols)
			}
		}
	} else {
		// 单格起点：按数据尺寸自动扩展区域（与 excelize 行为一致）
		if need := len(values); need > rows {
			rows = need
			r2 = r1 + need - 1
		}
		if need := maxLen(values); need > cols {
			cols = need
			c2 = c1 + need - 1
		}
	}
	for i, rowVals := range values {
		for j, val := range rowVals {
			if val == nil {
				continue
			}
			if j >= cols || i >= rows {
				continue
			}
			cell := colNumToLetters(c1+j) + strconv.Itoa(r1+i+1)
			if err := writeRangeCell(fileMap, file, cell, val, sc); err != nil {
				return err
			}
		}
	}
	return nil
}

// writeRangeCell 把单个值（按类型分派）写入 fileMap 中的 worksheet。
func writeRangeCell(fileMap map[string][]byte, file, cell string, val interface{}, sc *sharedStringsCache) error {
	switch v := val.(type) {
	case string:
		return setCellInMap(fileMap, file, cell, CellTypeString, v, "", sc)
	case int:
		return setCellInMap(fileMap, file, cell, CellTypeNumeric, strconv.Itoa(v), "", sc)
	case int64:
		return setCellInMap(fileMap, file, cell, CellTypeNumeric, strconv.FormatInt(v, 10), "", sc)
	case float64:
		return setCellInMap(fileMap, file, cell, CellTypeNumeric, strconv.FormatFloat(v, 'f', -1, 64), "", sc)
	case float32:
		return setCellInMap(fileMap, file, cell, CellTypeNumeric, strconv.FormatFloat(float64(v), 'f', -1, 64), "", sc)
	case bool:
		b := "0"
		if v {
			b = "1"
		}
		return setCellInMap(fileMap, file, cell, CellTypeBool, b, "", sc)
	case CellFormula:
		return setCellInMap(fileMap, file, cell, CellTypeFormula, v.Formula, v.Result, sc)
	default:
		return fmt.Errorf("SetRange 不支持的值类型: %T", val)
	}
}

// maxLen 返回二维切片中最长一行的长度。
func maxLen(values [][]interface{}) int {
	n := 0
	for _, r := range values {
		if len(r) > n {
			n = len(r)
		}
	}
	return n
}

// extractRangeRows 从 worksheet XML 提取 [r1..r2]x[c1..c2] 区域内单元格显示值。
func extractRangeRows(ws string, fileMap map[string][]byte, c1, r1, c2, r2 int) [][]string {
	grid := make([][]string, r2-r1+1)
	for i := range grid {
		grid[i] = make([]string, c2-c1+1)
	}
	sdRe := regexp.MustCompile(`(?s)<sheetData\b[^>]*>(.*?)</sheetData>`)
	sd := sdRe.FindStringSubmatch(ws)
	if sd == nil {
		return grid
	}
	inner := sd[1]
	cRe := regexp.MustCompile(`(?s)<c\b[^>]*\br="([A-Z]+\d+)"[^>]*?(?:/>|>(.*?)</c>)`)
	for _, cm := range cRe.FindAllStringSubmatch(inner, -1) {
		ref := cm[1]
		fullCell := cm[0]
		col, row, err := parseCellRef(ref)
		if err != nil {
			continue
		}
		if col < c1 || col > c2 || row < r1 || row > r2 {
			continue
		}
		grid[row-r1][col-c1] = cellDisplayValue(fullCell, fileMap)
	}
	return grid
}
