package excelgo

// rows_cols.go 提供行列插入与删除能力。
//
// 设计原则（与全库一致——不破坏原有格式/数据/布局）：
//   - 仅平移受影响的元素属性（<row r>、<c r>、<mergeCell ref>、<col min/max>、分页符 <brk>、
//     打印区域 definedName 中的单元格引用），其余 XML 原样保留。
//   - 单元格 s 样式索引不变；合并单元格、列宽定义、分页符均随位置平移。
//   - 公式单元格内的 A1 引用字符串不做改写（与 excelize 行为一致，由调用方负责公式修正）。
//   - 浮动图片锚点（oneCellAnchor / twoCellAnchor 的 from/to col/row）随行列操作平移，
//     与 Excel 行为一致；锚点落入被删区间的图片随内容一并移除。

import (
	"fmt"
	"path"
	"regexp"
	"strconv"
	"strings"
)

// colNumToLetters 把 0 基列号转为字母（0->A, 25->Z, 26->AA）。
func colNumToLetters(col int) string {
	col++ // 内部用 1 基计算更直观
	var sb strings.Builder
	for col > 0 {
		col--
		sb.WriteByte(byte('A' + col%26))
		col /= 26
	}
	// 反转
	r := []byte(sb.String())
	for i, j := 0, len(r)-1; i < j; i, j = i+1, j-1 {
		r[i], r[j] = r[j], r[i]
	}
	return string(r)
}

// wsMaxRow 返回工作表中实际出现的最大行号（1 基）；空表返回 0。
// 同时考虑 <row r=".."> 与单元格的 <c r="A10"> 引用，避免漏判只有单元格没有 row 属性的表。
func wsMaxRow(ws string) int {
	max := 0
	for _, m := range regexp.MustCompile(`<row\b[^>]*\br="(\d+)"`).FindAllStringSubmatch(ws, -1) {
		if v, err := strconv.Atoi(m[1]); err == nil && v > max {
			max = v
		}
	}
	for _, m := range regexp.MustCompile(`<c\b[^>]*\br="[A-Za-z]{1,3}(\d+)"`).FindAllStringSubmatch(ws, -1) {
		if v, err := strconv.Atoi(m[1]); err == nil && v > max {
			max = v
		}
	}
	return max
}

// wsMaxCol 返回工作表中实际出现的最大列号（1 基）；空表返回 0。
// 同时考虑 <col min/max> 定义、单元格 <c r="J1"> 引用与 <row spans> 声明。
func wsMaxCol(ws string) int {
	max := 0
	for _, m := range regexp.MustCompile(`<col\b[^>]*\bmin="(\d+)"[^>]*\bmax="(\d+)"`).FindAllStringSubmatch(ws, -1) {
		lo, e1 := strconv.Atoi(m[1])
		hi, e2 := strconv.Atoi(m[2])
		if e1 == nil && lo > max {
			max = lo
		}
		if e2 == nil && hi > max {
			max = hi
		}
	}
	for _, m := range regexp.MustCompile(`<c\b[^>]*\br="([A-Za-z]{1,3})(\d+)"`).FindAllStringSubmatch(ws, -1) {
		if v, err := colLettersToNum(m[1]); err == nil && v > max {
			max = v
		}
	}
	for _, m := range regexp.MustCompile(`<row\b[^>]*\bspans="(\d+):(\d+)"`).FindAllStringSubmatch(ws, -1) {
		if v, err := strconv.Atoi(m[2]); err == nil && v > max {
			max = v
		}
	}
	return max
}

// InsertRows 在 sheetRef 的第 row 行（1 基）之前插入 n 行（n>=1）。
// 原有 row 及之后的行整体下移，行内单元格引用（含公式 <f> 内部引用）同步平移，样式不变。
//
// row 允许等于「最大行号 + 1」（即在表尾之后追加空行，符合 Excel 行为）；
// 超过该值则返回错误，避免误填空白导致文件里出现大片无意义空行。
func InsertRows(filename, sheetRef string, row, n int) error {
	if row < 1 || n < 1 {
		return fmt.Errorf("row 须>=1 且 n 须>=1")
	}
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return err
	}
	file, name, _, _, err := locateSheetInMap(fileMap, wb, sheetRef)
	if err != nil {
		return err
	}
	ws := string(fileMap[file])
	if mr := wsMaxRow(ws); row > mr+1 {
		return fmt.Errorf("插入行 %d 超出表尾（当前最大行 %d，最多可在 %d 处插入）", row, mr, mr+1)
	}
	ws, err = shiftRows(ws, row, n, nil)
	if err != nil {
		return err
	}
	// 在 row 位置插入 n 个空 <row>（带 r 属性，保持升序）
	ws = insertEmptyRows(ws, row, n)
	fileMap[file] = []byte(ws)
	// 浮动图片锚点随行下移（与 Excel 行为一致）
	shiftDrawingAnchorsForSheet(fileMap, file, row, n, nil, true)
	if err := writeMapToZip(filename, fileMap); err != nil {
		return err
	}
	_ = name
	return nil
}

// RemoveRows 删除 sheetRef 的第 row 行起共 n 行（1 基，n>=1）。
// 后续行整体上移，行内单元格引用同步平移，被删行的数据与样式一并移除。
func RemoveRows(filename, sheetRef string, row, n int) error {
	if row < 1 || n < 1 {
		return fmt.Errorf("row 须>=1 且 n 须>=1")
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
	ws := string(fileMap[file])
	// 校验删除范围不超出表尾
	if mr := wsMaxRow(ws); row+n-1 > mr {
		return fmt.Errorf("删除第 %d~%d 行超出表尾（当前最大行 %d）", row, row+n-1, mr)
	}
	// 先删 [row, row+n-1] 的 <row> 块
	ws = deleteRowRange(ws, row, n)
	// 后续行 r 减 n，单元格引用行号同步减 n；被删区间内的公式引用转 #REF!
	delRows := [2]int{row, row + n - 1}
	ws, err = shiftRows(ws, row+n, -n, &delRows)
	if err != nil {
		return err
	}
	fileMap[file] = []byte(ws)
	// 浮动图片锚点随行上移；落入被删行区间的图片一并移除（与 Excel 行为一致）
	shiftDrawingAnchorsForSheet(fileMap, file, row, -n, &delRows, true)
	return writeMapToZip(filename, fileMap)
}

// InsertCols 在 sheetRef 的第 col 列（1 基）之前插入 n 列（n>=1）。
// 原有 col 及之后的列整体右移，单元格引用、合并区域、列宽、列分页符同步平移。
//
// col 允许等于「最大列号 + 1」（即在表尾之后追加空列，符合 Excel 行为）；
// 超过该值则返回错误。
func InsertCols(filename, sheetRef string, col, n int) error {
	if col < 1 || n < 1 {
		return fmt.Errorf("col 须>=1 且 n 须>=1")
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
	ws := string(fileMap[file])
	if mc := wsMaxCol(ws); col > mc+1 {
		return fmt.Errorf("插入列 %d 超出表尾（当前最大列 %d，最多可在 %d 处插入）", col, mc, mc+1)
	}
	ws = shiftCols(ws, col, n, nil)
	fileMap[file] = []byte(ws)
	// 浮动图片锚点随列右移（与 Excel 行为一致）
	shiftDrawingAnchorsForSheet(fileMap, file, col, n, nil, false)
	return writeMapToZip(filename, fileMap)
}

// RemoveCols 删除 sheetRef 的第 col 列起共 n 列（1 基，n>=1）。
func RemoveCols(filename, sheetRef string, col, n int) error {
	if col < 1 || n < 1 {
		return fmt.Errorf("col 须>=1 且 n 须>=1")
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
	ws := string(fileMap[file])
	// 校验删除范围不超出表尾
	if mc := wsMaxCol(ws); col+n-1 > mc {
		return fmt.Errorf("删除第 %d~%d 列超出表尾（当前最大列 %d）", col, col+n-1, mc)
	}
	delCols := [2]int{col, col + n - 1}
	ws = shiftCols(ws, col, -n, &delCols)
	fileMap[file] = []byte(ws)
	// 浮动图片锚点随列左移；落入被删列区间的图片一并移除（与 Excel 行为一致）
	shiftDrawingAnchorsForSheet(fileMap, file, col, -n, &delCols, false)
	return writeMapToZip(filename, fileMap)
}

// ---------- 行平移 ----------

// shiftRows 把 worksheet 中所有引用行号 >= pivot 的 <row r> 及 <c r="Xrow"> 平移 delta，
// 同时平移公式 <f> 内部的行引用。delRows 为被删除的 [起,止] 行区间（1 基闭区间），
// 引用落入该区间则转 #REF!；插入场景传 nil。
func shiftRows(ws string, pivot, delta int, delRows *[2]int) (string, error) {
	if delta == 0 {
		return ws, nil
	}
	// <row r="N"> 标签
	ws = shiftRowTags(ws, pivot, delta)
	// <c r="XN"> 单元格引用行号 + 公式内引用
	ws = shiftCellRowRefs(ws, pivot, delta, delRows)
	// 合并单元格 ref（含行号）
	ws = shiftMergeRefs(ws, pivot, delta, false)
	// 打印区域 definedName 中的行号
	ws = shiftPrintAreaRows(ws, pivot, delta)
	// 行分页符 <brk id="N" />
	ws = shiftRowBreaks(ws, pivot, delta)
	return ws, nil
}

// shiftRowTags 平移 <row r="N"> 的 N（仅 N>=pivot）。
func shiftRowTags(ws string, pivot, delta int) string {
	re := regexp.MustCompile(`(<row\b[^>]*\br=")(\d+)(")`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		r, _ := strconv.Atoi(sub[2])
		if r >= pivot {
			r += delta
			if r < 1 {
				return "" // 上移越界则移除该 row 标签（内容随后清理）
			}
		}
		return sub[1] + strconv.Itoa(r) + sub[3]
	})
}

// shiftCellRowRefs 平移所有 <c r="COLrow"> 的 row 部分（row>=pivot），
// 并改写其 <f> 公式内的行引用（含 #REF! 处理）。
func shiftCellRowRefs(ws string, pivot, delta int, delRows *[2]int) string {
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="([A-Za-z]+)(\d+)"[^>]*?(?:/>|>(.*?)</c>)`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		colLetters := sub[1]
		rowNum, _ := strconv.Atoi(sub[2])
		newRow := rowNum
		if rowNum >= pivot {
			newRow += delta
			if newRow < 1 {
				return "" // 越界移除
			}
		}
		// 改写 r 属性
		newR := colLetters + strconv.Itoa(newRow)
		m2 := regexp.MustCompile(`\br="[A-Za-z]+\d+"`).ReplaceAllString(m, `r="`+newR+`"`)
		// 改写公式内引用（列维度不变：pivotCol=0, deltaCol=0）
		if strings.Contains(m2, "<f") {
			m2 = shiftFormulaInCell(m2, pivot, 0, delta, 0, delRows, nil)
		}
		return m2
	})
}

// deleteRowRange 删除 [row, row+n-1] 的所有 <row> 块（含内容）。
// 必须同时兼容自闭合 <row r="2"/> 与带内容的 <row r="3">...</row>，
// 否则在存在空行（InsertRows 产生的自闭合 <row/>）时，正则的 .*? 会越过 /> 边界
// 把下一行一并吞入同一匹配，导致误删多行、数据丢失。
func deleteRowRange(ws string, row, n int) string {
	re := regexp.MustCompile(`(?s)<row\b[^>]*\br="(\d+)"[^>]*?(?:/>|>(.*?)</row>)`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		r, _ := strconv.Atoi(sub[1])
		if r >= row && r < row+n {
			return ""
		}
		return m
	})
}

// insertEmptyRows 在 sheetData 内的 row 位置插入 n 个空 <row r="...">。
func insertEmptyRows(ws string, row, n int) string {
	// 找 sheetData 中第一个 r >= row 的 <row>，插在其前；否则插在 </sheetData> 前
	sdOpen := strings.Index(ws, "<sheetData")
	sdClose := strings.LastIndex(ws, "</sheetData>")
	if sdOpen == -1 || sdClose == -1 {
		return ws
	}
	inner := ws[sdOpen:sdClose]
	rowOpenRe := regexp.MustCompile(`<row\b[^>]*\br="(\d+)"`)
	var insertAt int = sdClose
	for _, rm := range rowOpenRe.FindAllStringSubmatchIndex(inner, -1) {
		// rm[2],rm[3] 是捕获组1（行号）的位置
		numStr := inner[rm[2]:rm[3]]
		r, _ := strconv.Atoi(numStr)
		if r >= row {
			// 该 <row> 起始位置（在 inner 中的偏移）
			insertAt = sdOpen + rm[0]
			break
		}
	}
	var b strings.Builder
	for i := 0; i < n; i++ {
		b.WriteString(`<row r="`)
		b.WriteString(strconv.Itoa(row + i))
		b.WriteString(`"/>`)
	}
	return ws[:insertAt] + b.String() + ws[insertAt:]
}

// shiftMergeRefs 平移 <mergeCell ref="A1:B2"> 的行列（按 pivot/delta，onlyRow=true 只平移行）。
func shiftMergeRefs(ws string, pivot, delta int, onlyRow bool) string {
	re := regexp.MustCompile(`(<mergeCell\b[^>]*\bref=")([^"]+)(")`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		newRef := shiftRef(sub[2], pivot, delta, onlyRow)
		return sub[1] + newRef + sub[3]
	})
}

// shiftPrintAreaRows 平移打印区域 definedName 中的行号（localSheetId 不变）。
func shiftPrintAreaRows(ws string, pivot, delta int) string {
	re := regexp.MustCompile(`(?s)(<definedName\b[^>]*\bname="_xlnm\.Print_Area"[^>]*>)([^<]*)(</definedName>)`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		newRange := shiftRef(sub[2], pivot, delta, false)
		return sub[1] + newRange + sub[3]
	})
}

// shiftRowBreaks 平移行分页符 <brk id="N" />（类型 行，axis 隐含）。只平移 >= pivot 的。
func shiftRowBreaks(ws string, pivot, delta int) string {
	re := regexp.MustCompile(`(<brk\b[^>]*\bid=")(\d+)(")`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		id, _ := strconv.Atoi(sub[2])
		// 只处理行分页符（通常不带 axis 或 axis="0"/"rot"）
		if strings.Contains(m, `axis="1"`) {
			return m // 列分页符不在此平移
		}
		if id >= pivot {
			id += delta
			if id < 1 {
				return ""
			}
		}
		return sub[1] + strconv.Itoa(id) + sub[3]
	})
}

// ---------- 列平移 ----------

// shiftCols 平移 worksheet 中所有列 >= pivot 的引用。delta>0 右移，<0 左移（删除）。
// 同时平移公式 <f> 内部的列引用。delCols 为被删除的 [起,止] 列区间（1 基闭区间），
// 引用落入该区间则转 #REF!；插入场景传 nil。
func shiftCols(ws string, pivot, delta int, delCols *[2]int) string {
	if delta == 0 {
		return ws
	}
	// 单元格引用列字母 + 公式内列引用
	ws = shiftCellColRefs(ws, pivot, delta, delCols)
	// 合并单元格区域列
	ws = shiftMergeRefsCols(ws, pivot, delta)
	// 列宽定义 <col min max>
	ws = shiftColDefs(ws, pivot, delta)
	// 列分页符 <brk id axis="1">
	ws = shiftColBreaks(ws, pivot, delta)
	// 打印区域列字母
	ws = shiftPrintAreaCols(ws, pivot, delta)
	return ws
}

// shiftCellColRefs 平移所有 <c> 单元格的列引用（列号 >= pivot）。
// 当 delta<0（删除列）且列落在被删区间 [pivot, pivot-delta-1] 内时，移除该 <c>（返回 ""）。
// 注意：必须匹配整个 <c> 元素（自闭合或带子元素），保留其余属性与子内容；
// 同时改写其 <f> 公式内的列引用（含 #REF! 处理）。
func shiftCellColRefs(ws string, pivot, delta int, delCols *[2]int) string {
	// 匹配完整 <c>：自闭合 <c .../> 或 <c ...>...</c>
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="([A-Za-z]+)(\d+)"[^>]*?(?:/>|>(.*?)</c>)`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		colLetters := sub[1]
		rowStr := sub[2]
		col, err := colLettersToNum(colLetters)
		if err != nil {
			return m
		}
		if delta < 0 && col >= pivot && col <= pivot-delta-1 {
			return "" // 删除该列单元格
		}
		newCol := col
		if col >= pivot {
			newCol = col + delta
			if newCol < 1 {
				return ""
			}
		}
		// 替换 r 属性中的列字母
		newRef := colNumToLetters(newCol-1) + rowStr
		m2 := regexp.MustCompile(`\br="[A-Za-z]+\d+"`).ReplaceAllString(m, `r="`+newRef+`"`)
		// 改写公式内引用（行维度不变：pivotRow=0, deltaRow=0）
		if strings.Contains(m2, "<f") {
			m2 = shiftFormulaInCell(m2, 0, pivot, 0, delta, nil, delCols)
		}
		return m2
	})
}

// shiftMergeRefsCols 平移 <mergeCell ref> 的列部分。
func shiftMergeRefsCols(ws string, pivot, delta int) string {
	re := regexp.MustCompile(`(<mergeCell\b[^>]*\bref=")([^"]+)(")`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		newRef := shiftRefColsOnly(sub[2], pivot, delta)
		return sub[1] + newRef + sub[3]
	})
}

// shiftColDefs 平移 <col min=".." max=".."> 的 min/max（>= pivot）。
func shiftColDefs(ws string, pivot, delta int) string {
	re := regexp.MustCompile(`(<col\b[^>]*\bmin=")(\d+)("[^>]*\bmax=")(\d+)(")`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		min, _ := strconv.Atoi(sub[2])
		max, _ := strconv.Atoi(sub[4])
		if min >= pivot {
			min += delta
			max += delta
			if min < 1 {
				return "" // 越界移除
			}
		}
		return sub[1] + strconv.Itoa(min) + sub[3] + strconv.Itoa(max) + sub[5]
	})
}

// shiftColBreaks 平移列分页符 <brk id="N" axis="1" />。
func shiftColBreaks(ws string, pivot, delta int) string {
	re := regexp.MustCompile(`(<brk\b[^>]*\bid=")(\d+)("[^>]*\baxis="1")`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		id, _ := strconv.Atoi(sub[2])
		if id >= pivot {
			id += delta
			if id < 1 {
				return ""
			}
		}
		return sub[1] + strconv.Itoa(id) + sub[3]
	})
}

// shiftPrintAreaCols 平移打印区域 definedName 中的列字母。
func shiftPrintAreaCols(ws string, pivot, delta int) string {
	re := regexp.MustCompile(`(?s)(<definedName\b[^>]*\bname="_xlnm\.Print_Area"[^>]*>)([^<]*)(</definedName>)`)
	return re.ReplaceAllStringFunc(ws, func(m string) string {
		sub := re.FindStringSubmatch(m)
		newRange := shiftRefColsOnly(sub[2], pivot, delta)
		return sub[1] + newRange + sub[3]
	})
}

// ---------- 引用平移工具 ----------

// shiftRef 平移形如 "A1:B2" 的单元格区域引用中的行号（仅行）。
func shiftRef(ref string, pivot, delta int, _ bool) string {
	parts := strings.Split(ref, ":")
	for i := range parts {
		col, row, err := parseCellRef(parts[i])
		if err != nil {
			continue
		}
		if row >= pivot {
			row += delta
			if row < 0 {
				row = 0
			}
		}
		parts[i] = colNumToLetters(col) + strconv.Itoa(row+1)
	}
	return strings.Join(parts, ":")
}

// shiftRefColsOnly 平移区域引用的列字母（仅列）。pivot 为 1 基列号。
func shiftRefColsOnly(ref string, pivot, delta int) string {
	parts := strings.Split(ref, ":")
	for i := range parts {
		col, row, err := parseCellRef(parts[i])
		if err != nil {
			continue
		}
		col1 := col + 1 // 转 1 基与 pivot 对齐
		if col1 >= pivot {
			col1 += delta
			if col1 < 1 {
				col1 = 1
			}
		}
		parts[i] = colNumToLetters(col1-1) + strconv.Itoa(row+1)
	}
	return strings.Join(parts, ":")
}

// colLettersToNum 把列字母转为 1 基列号（A->1）。
func colLettersToNum(s string) (int, error) {
	s = strings.ToUpper(strings.TrimSpace(s))
	if s == "" {
		return 0, fmt.Errorf("空列字母")
	}
	n := 0
	for _, ch := range s {
		if ch < 'A' || ch > 'Z' {
			return 0, fmt.Errorf("无效列字母: %q", s)
		}
		n = n*26 + int(ch-'A'+1)
	}
	return n, nil
}

// ---------- 浮动图片锚点平移 ----------
// Excel 中插入/删除行列时，浮动图片（oneCellAnchor / twoCellAnchor）的锚点随单元格平移，
// 与单元格逻辑一致；下方函数补齐该平移，使图片保持相对位置、删除时移除落入被删区间的图片。

// sheetDrawingFiles 返回某工作表通过 rels 关联的所有 drawing 文件路径（绝对 zip 路径）。
func sheetDrawingFiles(fileMap map[string][]byte, sheetFile string) []string {
	relsFile := "xl/worksheets/_rels/" + path.Base(sheetFile) + ".rels"
	if !fileExistsInMap(fileMap, relsFile) {
		return nil
	}
	rels := &Relationships{}
	if err := getXMLFromMap(fileMap, relsFile, rels); err != nil {
		return nil
	}
	var drawings []string
	for _, rel := range rels.Relationship {
		if strings.Contains(rel.Type, "/relationships/drawing") {
			drawings = append(drawings, resolveTarget("xl/worksheets", rel.Target))
		}
	}
	return drawings
}

// shiftDrawingAnchorsForSheet 平移 sheetFile 关联的所有 drawing 中的浮动图片锚点。
// pivot 为 1 基行列号，delta 为平移量（插入为正、删除为负），delBand 为被删行列闭区间
// （仅删除时非 nil）；锚点落入被删区间则移除该图片；isRow=true 平移行锚点，否则平移列锚点。
func shiftDrawingAnchorsForSheet(fileMap map[string][]byte, sheetFile string, pivot, delta int, delBand *[2]int, isRow bool) {
	for _, dw := range sheetDrawingFiles(fileMap, sheetFile) {
		if !fileExistsInMap(fileMap, dw) {
			continue
		}
		fileMap[dw] = []byte(shiftDrawingAnchors(string(fileMap[dw]), pivot, delta, delBand, isRow))
	}
}

// drawing 命名空间前缀可变：Excel/WPS 多用 "xdr:oneCellAnchor"，而 excelgo 自己
// 写出的 drawing 使用默认命名空间（"oneCellAnchor"）。因此所有匹配一律写成
// `(?:xdr:)?`，两种形式都能命中。
var (
	reDrawingAnchor = regexp.MustCompile(`(?s)<(?:xdr:)?(?:oneCellAnchor|twoCellAnchor)\b.*?</(?:xdr:)?(?:oneCellAnchor|twoCellAnchor)>`)
	reAnchorFrom    = regexp.MustCompile(`(?s)<(?:xdr:)?from\b.*?</(?:xdr:)?from>`)
	reAnchorBlock   = regexp.MustCompile(`(?s)<(?:xdr:)?(?:from|to)\b.*?</(?:xdr:)?(?:from|to)>`)
	reAnchorCol     = regexp.MustCompile(`(<(?:xdr:)?col>)(\d+)(</(?:xdr:)?col>)`)
	reAnchorRow     = regexp.MustCompile(`(<(?:xdr:)?row>)(\d+)(</(?:xdr:)?row>)`)
)

// shiftDrawingAnchors 平移 drawing XML 中所有 oneCellAnchor/twoCellAnchor 锚点的 col/row。
func shiftDrawingAnchors(drawingXML string, pivot, delta int, delBand *[2]int, isRow bool) string {
	return reDrawingAnchor.ReplaceAllStringFunc(drawingXML, func(anchor string) string {
		// 删除场景：若锚点 from 落入被删区间，移除整个图片锚点
		if delBand != nil {
			if fm := reAnchorFrom.FindString(anchor); fm != "" {
				col, row := anchorColRow(fm)
				v := col
				if isRow {
					v = row
				}
				// 锚点为 0 基，转成 1 基与被删闭区间比较
				if (v+1) >= delBand[0] && (v+1) <= delBand[1] {
					return ""
				}
			}
		}
		// 平移 from / to 块
		return reAnchorBlock.ReplaceAllStringFunc(anchor, func(block string) string {
			return shiftAnchorBlock(block, pivot, delta, isRow)
		})
	})
}

// anchorColRow 从 <from>...</from> 块中提取 col 与 row（均为 0 基）。
// 兼容带 "xdr:" 前缀与默认命名空间两种写法。
func anchorColRow(block string) (col, row int) {
	if m := reAnchorCol.FindStringSubmatch(block); m != nil {
		col, _ = strconv.Atoi(m[2])
	}
	if m := reAnchorRow.FindStringSubmatch(block); m != nil {
		row, _ = strconv.Atoi(m[2])
	}
	return
}

// shiftAnchorBlock 平移一个 from/to 块内的 col 或 row（colOff/rowOff 偏移量保持不变）。
// drawing 锚点的 col/row 为 0 基，而传入的 pivot 为 1 基，故比较时取 v+1。
func shiftAnchorBlock(block string, pivot, delta int, isRow bool) string {
	re := reAnchorRow
	if !isRow {
		re = reAnchorCol
	}
	return re.ReplaceAllStringFunc(block, func(m string) string {
		sub := re.FindStringSubmatch(m)
		v, _ := strconv.Atoi(sub[2])
		if v+1 >= pivot { // 0 基锚点换算成 1 基与 pivot 对齐
			v += delta
			if v < 0 {
				v = 0
			}
		}
		return sub[1] + strconv.Itoa(v) + sub[3]
	})
}
