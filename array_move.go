package excelgo

// 数组公式与区域搬移。
//
// 两者的共同点：都需要**批量改写单元格并让公式引用随之重映射**。
// 库内已有 shiftFormulaText（插删行列时用的引用平移），本文件复用它，
// 避免为 move_range 再写一套公式重写逻辑。

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// ---------- 数组公式 ----------

// SetCellArrayFormula 在 cell 写入**数组公式**（CSE 公式），作用区域为 ref。
//
// 数组公式与普通公式的区别：OOXML 里 <f> 带 t="array" 属性和 ref 属性，
// 标明该公式覆盖的区域。Excel 2007 及以前靠它实现 CSE（Ctrl+Shift+Enter）。
// 典型用途：SUMPRODUCT 之外的多单元格一次性计算、动态数组溢出。
//
//	result 传空字符串表示不写缓存结果（Excel 打开时会自行计算）。
//
// 示意：
//
//	_ = ws.SetCellArrayFormula("A1", "A1:B3", "A1:A3*2")
func (s *WorkSheet) SetCellArrayFormula(cell, ref, formula, result string) error {
	cref := newCellRef(cell)
	if cref == "" {
		return fmt.Errorf("无效的单元格引用: %q", cell)
	}
	// 数组公式的 ref 可以是单格，也可以是矩形区域
	r := newRangeRef(ref)
	if r == "" {
		return fmt.Errorf("数组公式的 ref 无效: %q", ref)
	}
	formula = strings.TrimPrefix(formula, "=")
	if formula == "" {
		return fmt.Errorf("数组公式不能为空")
	}
	ws := string(s.ws())
	inner := `<f t="array" ref="` + string(r) + `">` + safeText(formula) + `</f>`
	if result != "" {
		inner += `<v>` + safeText(result) + `</v>`
	}
	// 数组公式的结果类型由内容推断：字符串结果用 t="str"，数值不加 t
	tAttr := ""
	if isTextFormulaResult(result) {
		tAttr = ` t="str"`
	}
	col, row, _ := refCoords(string(cref))
	ws = setCellContents(ws, []cellOp{{row: row, col: col,
		xml: `<c r="` + string(cref) + `"` + tAttr + `>` + inner + `</c>`}})
	s.setWS(ws)
	return nil
}

// SetCellArrayFormula 在 cell 写入数组公式，作用区域为 ref（包级 API）。
func SetCellArrayFormula(filename, sheetRef, cell, ref, formula, result string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.SetCellArrayFormula(cell, ref, formula, result)
	})
}

// isTextFormulaResult 粗判公式结果是否为文本（决定是否写 t="str"）。
func isTextFormulaResult(result string) bool {
	if result == "" {
		return false
	}
	if _, err := strconv.ParseFloat(result, 64); err == nil {
		return false
	}
	switch strings.ToUpper(result) {
	case "TRUE", "FALSE":
		return false
	}
	return true
}

// refCoords 把 "AB12" 拆成 (列号, 行号)，均为 **0 基**（与库内 parseCellRef 一致）。
//
// 本文件的区域计算统一用 0 基，只在生成坐标字符串时由 refOf 转回 "A1" 形式。
// 混用 1 基/0 基是早期版本的 bug 根源：区域被算成 top=0 col=0，整批写入错位。
func refCoords(ref string) (col, row int, ok bool) {
	c, r, err := parseCellRef(ref)
	if err != nil {
		return 0, 0, false
	}
	return c, r, true
}

// setCellContents 按坐标批量写入/清空单元格，返回新的 sheet XML。
//
// 实现策略：**按行重建**而不是逐格字符串拼接。
// 早期版本用"找行 → 找列 → 插入"的逐格拼接，结果在 MoveRange 这种
// "先清空一批、再写入另一批"的场景下会把行号算错（曾出现 <row r="0">、
// 单元格落到相邻行）。按行重建只解析一次、只序列化一次，天然避免错位。
//
// ops 里的坐标与 xml 均为 0 基；xml 为空串表示"清空该格"。
type cellOp struct {
	row, col int // 0 基
	xml      string
}

// setCellContents 把 sheet XML 中受 ops 影响的行整体重建。
func setCellContents(ws string, ops []cellOp) string {
	if len(ops) == 0 {
		return ws
	}
	// 按行聚合
	byRow := map[int][]cellOp{}
	for _, op := range ops {
		byRow[op.row] = append(byRow[op.row], op)
	}
	rows := map[int]string{} // row -> 行 XML
	order := []int{}         // 保持行号递增
	rowRe := regexp.MustCompile(`(?s)<row r="(\d+)"([^>]*)>(.*?)</row>`)
	// 解析现有各行的单元格
	for _, m := range rowRe.FindAllStringSubmatchIndex(ws, -1) {
		rn, _ := strconv.Atoi(ws[m[2]:m[3]])
		rn-- // XML 行号 1 基 -> 内部 0 基
		attrs := ws[m[4]:m[5]]
		inner := ws[m[6]:m[7]]
		cells := map[int]string{} // col -> <c> XML
		var cols []int
		cellRe := regexp.MustCompile(`(?s)<c r="([A-Z]+)(\d+)"(?:\s[^>]*)?(?:/>|>.*?</c>)`)
		for _, cm := range cellRe.FindAllStringSubmatchIndex(inner, -1) {
			ref := inner[cm[2]:cm[3]]
			c, _, err := parseCellRef(ref)
			if err != nil {
				continue
			}
			if _, dup := cells[c]; !dup {
				cols = append(cols, c)
			}
			cells[c] = inner[cm[0]:cm[1]]
		}
		sortInts(cols)
		// 应用本行的 ops
		if list, ok := byRow[rn]; ok {
			for _, op := range list {
				if op.xml == "" {
					delete(cells, op.col)
				} else {
					if _, exists := cells[op.col]; !exists {
						cols = append(cols, op.col)
						sortInts(cols)
					}
					cells[op.col] = op.xml
				}
			}
			delete(byRow, rn)
		}
		var b strings.Builder
		// 行号输出必须 +1：OOXML 的 row r 是 1 基，写 0 会被 Excel 判损坏
		b.WriteString(`<row r="` + strconv.Itoa(rn+1) + `"` + attrs + `>`)
		for _, c := range cols {
			if x, ok := cells[c]; ok {
				b.WriteString(x)
			}
		}
		b.WriteString(`</row>`)
		rows[rn] = b.String()
		order = append(order, rn)
	}
	// 新增的行（源里不存在的）
	for rn, list := range byRow {
		colsMap := map[int]string{}
		for _, op := range list {
			if op.xml == "" {
				continue
			}
			colsMap[op.col] = op.xml
		}
		if len(colsMap) == 0 {
			continue
		}
		var cols []int
		for c := range colsMap {
			cols = append(cols, c)
		}
		sortInts(cols)
		var b strings.Builder
		b.WriteString(`<row r="` + strconv.Itoa(rn+1) + `">`)
		for _, c := range cols {
			b.WriteString(colsMap[c])
		}
		b.WriteString(`</row>`)
		rows[rn] = b.String()
		order = append(order, rn)
	}
	sortInts(order)

	// 整体替换 sheetData 内容
	sdRe := regexp.MustCompile(`(?s)<sheetData\b[^>]*>.*?</sheetData>`)
	loc := sdRe.FindStringIndex(ws)
	if loc == nil {
		return ws
	}
	// 保留原 sheetData 的属性
	openTag := regexp.MustCompile(`<sheetData\b[^>]*>`).FindString(ws[loc[0]:loc[1]])
	var b strings.Builder
	b.WriteString(openTag)
	for _, rn := range order {
		if x, ok := rows[rn]; ok {
			b.WriteString(x)
		}
	}
	b.WriteString(`</sheetData>`)
	return ws[:loc[0]] + b.String() + ws[loc[1]:]
}

// sortInts 就地升序排序（量级极小，插入排序足够）。
func sortInts(a []int) {
	for i := 1; i < len(a); i++ {
		for j := i; j > 0 && a[j-1] > a[j]; j-- {
			a[j-1], a[j] = a[j], a[j-1]
		}
	}
}

// ---------- 区域搬移 ----------

// MoveRange 把 from 区域的值（含样式）整体搬到 to 区域，公式引用随之重映射。
//
// 与"复制后删除"的区别：MoveRange 是**移动**语义 —— 源区域被清空。
// 公式重映射规则：引用被搬动区域的公式按同样的偏移调整，指向新位置；
// 引用外部的公式不变。
//
// 示例（把 A1:A10 移到 B1:B10，向右一列）：
//
//	_ = ws.MoveRange("A1:A10", "B1")
func (s *WorkSheet) MoveRange(from, to string) error {
	src, err := newRangeRect(from)
	if err != nil {
		return fmt.Errorf("源区域无效: %w", err)
	}
	dst, err := newRangeRect(to)
	if err != nil {
		return fmt.Errorf("目标位置无效: %w", err)
	}
	// 目标可以是"锚点单格"（Excel 的做法）：此时沿用源的形状。
	// 只有当目标本身带了多格且与源形状不符时才报错 —— 那多半是调用方笔误。
	if dst.h != 1 || dst.w != 1 {
		if src.h != dst.h || src.w != dst.w {
			return fmt.Errorf("源区域 %dx%d 与目标 %dx%d 形状不一致 —— "+
				"MoveRange 要求形状相同，或把目标写成锚点单格（如 \"D1\"）",
				src.w, src.h, dst.w, dst.h)
		}
	} else {
		// 锚点单格：撑成与源同形
		dst.h, dst.w = src.h, src.w
	}
	deltaRow := dst.top - src.top
	deltaCol := dst.col - src.col
	if deltaRow == 0 && deltaCol == 0 {
		return nil // 无位移
	}

	ws := string(s.ws())

	// 1) 读出源区域的值与样式
	type cellData struct {
		row, col int
		value    string
		style    int
		hasStyle bool
		hasValue bool
	}
	var cells []cellData
	for r := src.top; r < src.top+src.h; r++ {
		for c := src.col; c < src.col+src.w; c++ {
			ref := refOf(r, c)
			d := cellData{row: r, col: c}
			if v, err := s.GetCellValue(ref); err == nil && v != nil {
				d.value = displayOf(v)
				d.hasValue = true
			}
			if idx := s.GetCellStyleIndex(ref); idx > 0 {
				d.style = idx
				d.hasStyle = true
			}
			cells = append(cells, d)
		}
	}

	// 2) 一次性重建：清空源区域 + 写入目标位置（公式按位移重映射）
	//
	// 必须**合并成一次**重建。若分两步（先清空再逐格写），按行重建会各自
	// 解析/序列化一次，中途的行状态与最终结果不一致。
	ops := make([]cellOp, 0, len(cells)*2)
	for _, d := range cells {
		ops = append(ops, cellOp{row: d.row, col: d.col, xml: ""}) // 清空源
	}
	for _, d := range cells {
		nr := d.row + deltaRow
		nc := d.col + deltaCol
		if !d.hasValue {
			continue
		}
		dstRef := refOf(nr, nc)
		v := d.value
		if strings.HasPrefix(v, "=") {
			// 公式：按位移重映射引用
			v = "=" + shiftFormulaText(strings.TrimPrefix(v, "="),
				d.row, d.col, deltaRow, deltaCol, nil, nil)
		}
		xml := buildTextOrNumberCellXML(cellRef(dstRef), v)
		if d.hasStyle {
			// 保留样式索引
			xml = strings.Replace(xml, `<c r="`, `<c s="`+strconv.Itoa(d.style)+`" r="`, 1)
		}
		ops = append(ops, cellOp{row: nr, col: nc, xml: xml})
	}
	ws = setCellContents(ws, ops)
	s.setWS(ws)
	return nil
}

// MoveRange 把 from 区域的值搬到 to 位置（包级 API）。
func MoveRange(filename, sheetRef, from, to string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.MoveRange(from, to)
	})
}

// rangeRect 是一个矩形区域（1 基闭区间）。
type rangeRect struct {
	top, col int // 左上角行列（0 基）
	h, w     int // 高（行数）与宽（列数）
}

// newRangeRect 解析 "A1:C10" 形式的区域引用。
// 也接受单格 "A1"（视作 1x1）与整行整列 "A:C" / "1:1"（用最大行列兜底）。
func newRangeRect(ref string) (rangeRect, error) {
	r := newRangeRef(ref)
	if r == "" {
		return rangeRect{}, fmt.Errorf("无法解析区域引用 %q", ref)
	}
	s := string(r)
	parts := strings.Split(s, ":")
	toCol := func(p string) (int, bool) {
		c, _, ok := refCoords(p)
		return c, ok
	}
	toRow := func(p string) (int, bool) {
		_, rr, ok := refCoords(p)
		return rr, ok
	}
	if len(parts) == 1 {
		// 单格（可能是 "A1"）
		c, ok1 := toCol(parts[0])
		rr, ok2 := toRow(parts[0])
		if !ok1 || !ok2 {
			return rangeRect{}, fmt.Errorf("无法解析 %q", ref)
		}
		return rangeRect{top: rr, col: c, h: 1, w: 1}, nil
	}
	c1, ok1 := toCol(parts[0])
	r1, ok2 := toRow(parts[0])
	c2, ok3 := toCol(parts[1])
	r2, ok4 := toRow(parts[1])
	if !ok1 || !ok2 || !ok3 || !ok4 {
		return rangeRect{}, fmt.Errorf("无法解析 %q", ref)
	}
	if c2 < c1 {
		c1, c2 = c2, c1
	}
	if r2 < r1 {
		r1, r2 = r2, r1
	}
	return rangeRect{top: r1, col: c1, h: r2 - r1 + 1, w: c2 - c1 + 1}, nil
}

// refOf 由 **0 基**行列号生成 "A1" 形式的坐标：(0,0) -> "A1"、(1,3) -> "D2"。
//
// 行号在输出时 +1 —— 内部一律 0 基，但坐标字符串按 OOXML 约定是 1 基。
// 忘记 +1 会生成 "A0" 这种非法坐标，导致读回全为空。
func refOf(row, col int) string {
	if col < 0 || row < 0 {
		return ""
	}
	// 0 基列号转字母：0->A, 25->Z, 26->AA（经典 26 进制，无 0 陷阱）
	name := ""
	for c := col + 1; c > 0; {
		c--
		name = string(rune('A'+c%26)) + name
		c /= 26
	}
	return name + strconv.Itoa(row+1)
}

// buildTextOrNumberCellXML 按值形态生成 <c> 片段（数值不加 t，文本用 inlineStr）。
func buildTextOrNumberCellXML(ref cellRef, v string) string {
	if isNumericLiteral(v) {
		return `<c r="` + string(ref) + `"><v>` + safeText(v) + `</v></c>`
	}
	return `<c r="` + string(ref) + `" t="inlineStr"><is><t xml:space="preserve">` +
		safeText(v) + `</t></is></c>`
}

// isNumericLiteral 判断字符串是否为纯数值字面量。
func isNumericLiteral(s string) bool {
	if s == "" {
		return false
	}
	_, err := strconv.ParseFloat(s, 64)
	return err == nil
}

// displayOf 把 GetCellValue 的返回值转为可写回的显示文本。
func displayOf(v interface{}) string {
	switch n := v.(type) {
	case nil:
		return ""
	case string:
		return n
	case bool:
		if n {
			return "TRUE"
		}
		return "FALSE"
	case float64:
		if n == float64(int64(n)) {
			return strconv.FormatInt(int64(n), 10)
		}
		return strconv.FormatFloat(n, 'f', -1, 64)
	case int:
		return strconv.Itoa(n)
	}
	return fmt.Sprint(v)
}
