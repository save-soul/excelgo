package excelgo

// coords.go 提供单元格坐标与列名的双向转换（excelize 同名 API 的对应实现）。
//
// 术语与基约定：
//   - 「列名」：由字母构成的列标识，如 "A"、"AK"（不区分大小写输入）；
//   - 「索引」：本库内部统一用 0 基（colNumToLetters / parseCellRef 均如此）；
//   - 「坐标」：列名 + 行号，如 "AK74"，行号为 1 基（与 OOXML 的 r 属性一致）。
//
// 对外导出的六个函数遵循 excelize 的**1 基**约定，便于从 excelize 迁移的代码
// 直接替换包名即可运行；而库内既有 helper 保持 0 基不变，故本文件显式做基转换。

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// maxColumn 是 Excel 支持的最大列数（XFD）。
const maxColumn = 16384

// maxRow 是 Excel 支持的最大行号。
const maxRow = 1048576

// SplitCellName 把单元格坐标切分为列名与行号。
//
//	SplitCellName("AK74") // "AK", 74, nil
//
// 行号为 1 基。列名保持输入的大小写（"ak74" → "ak"）。
func SplitCellName(cell string) (string, int, error) {
	m := coordSplitRe.FindStringSubmatch(strings.TrimSpace(cell))
	if m == nil {
		return "", 0, fmt.Errorf("无效的单元格引用: %q", cell)
	}
	row, err := strconv.Atoi(m[2])
	if err != nil {
		return "", 0, fmt.Errorf("无效的单元格引用: %q", cell)
	}
	if row < 1 || row > maxRow {
		return "", 0, fmt.Errorf("行号超出范围 (1-%d): %q", maxRow, cell)
	}
	return m[1], row, nil
}

// JoinCellName 把列名与 1 基行号组合为单元格坐标。
//
//	JoinCellName("AK", 74) // "AK74", nil
//
// 列名不区分大小写；行号须在 1-maxRow 之间。
func JoinCellName(col string, row int) (string, error) {
	if _, err := ColumnNameToNumber(col); err != nil {
		return "", err
	}
	if row < 1 || row > maxRow {
		return "", fmt.Errorf("行号超出范围 (1-%d): %d", maxRow, row)
	}
	return strings.ToUpper(strings.TrimSpace(col)) + strconv.Itoa(row), nil
}

// ColumnNameToNumber 把列名转换为 1 基索引。
//
//	ColumnNameToNumber("AK") // 37, nil
//
// 不区分大小写。非法列名返回错误。
func ColumnNameToNumber(name string) (int, error) {
	n, err := lettersToColNum(name)
	if err != nil {
		return 0, err
	}
	return n + 1, nil
}

// ColumnNumberToName 把 1 基列索引转换为列名。
//
//	ColumnNumberToName(37) // "AK", nil
//
// 索引须在 1-maxColumn 之间。
func ColumnNumberToName(num int) (string, error) {
	if num < 1 || num > maxColumn {
		return "", fmt.Errorf("列索引超出范围 (1-%d): %d", maxColumn, num)
	}
	return colNumToLetters(num - 1), nil
}

// CellNameToCoordinates 把单元格坐标转换为 1 基 [列, 行]。
//
//	CellNameToCoordinates("A1")  // 1, 1, nil
//	CellNameToCoordinates("Z3")  // 26, 3, nil
func CellNameToCoordinates(cell string) (int, int, error) {
	col, row, err := SplitCellName(cell)
	if err != nil {
		return 0, 0, err
	}
	c, err := ColumnNameToNumber(col)
	if err != nil {
		return 0, 0, err
	}
	return c, row, nil
}

// CoordinatesToCellName 把 1 基 [列, 行] 转换为单元格坐标。
//
//	CoordinatesToCellName(1, 1)       // "A1", nil
//	CoordinatesToCellName(1, 1, true) // "$A$1", nil
//
// abs 为可选参数，true 时两侧都加 $ 绝对引用标记（$A$1），
// false 或不传时为相对引用（A1）。
func CoordinatesToCellName(col, row int, abs ...bool) (string, error) {
	name, err := ColumnNumberToName(col)
	if err != nil {
		return "", err
	}
	if row < 1 || row > maxRow {
		return "", fmt.Errorf("行号超出范围 (1-%d): %d", maxRow, row)
	}
	if len(abs) > 0 && abs[0] {
		return fmt.Sprintf("$%s$%d", name, row), nil
	}
	return name + strconv.Itoa(row), nil
}

// lettersToColNum 把列名转为 0 基列号（字母表基数 26 的累积展开）。
// 不区分大小写；空串或非法字符返回错误。
func lettersToColNum(name string) (int, error) {
	s := strings.ToUpper(strings.TrimSpace(name))
	if s == "" {
		return 0, fmt.Errorf("列名不能为空")
	}
	n := 0
	for i := 0; i < len(s); i++ {
		c := s[i]
		if c < 'A' || c > 'Z' {
			return 0, fmt.Errorf("无效的列名: %q", name)
		}
		n = n*26 + int(c-'A'+1)
	}
	n-- // 转 0 基
	if n < 0 || n >= maxColumn {
		return 0, fmt.Errorf("列名超出范围 (A-%s): %q", colNumToLetters(maxColumn-1), name)
	}
	return n, nil
}

// coordSplitRe 匹配单元格坐标并捕获列名/行号两组：字母列名 + 数字行号（不锚定大小写）。
var coordSplitRe = regexp.MustCompile(`^([A-Za-z]+)(\d+)$`)
