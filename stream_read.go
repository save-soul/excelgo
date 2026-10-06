package excelgo

// 只读流式模式。
//
// 背景：Open() 一次性把整个 zip 读进内存（map[string][]byte）。这对几万行的
// 表格尚可，但百万行级别会吃掉数百 MB —— 而用户往往只需要读其中几行。
//
// openpyxl 的 read_only 模式用的是 SAX 式流式解析：不把整个 sheet 建成 DOM，
// 而是逐行扫。本库的架构（字符串手术）无法照搬 SAX，但可以做等效优化：
//
//	**按需只读目标 sheet 的 XML，其余部件留在 zip 里不解析。**
//
// 对"读一个表的某几行"这个最常见场景，内存从 O(整文件) 降到 O(该表的 XML)。
//
// 用法：
//
//	b, err := excelgo.OpenReader("big.xlsx")   // 流式只读
//	rows, err := b.StreamRows("Sheet1")        // 逐行取，不载入全部
//
// 只读簿不能写：任何写方法都会返回错误，避免误以为改动会落盘。

import (
	"archive/zip"
	"fmt"
	"io"
	"strings"
)

// StreamBook 是只读流式工作簿句柄。
//
// 与 Book 的区别：只持有 zip 读取器与必要的小部件（workbook.xml /
// sharedStrings），sheet XML 按需读取且**读完即释放**。
type StreamBook struct {
	zipPath string

	sheetNames []string
	// sheetFile 记录每个表名对应的部件路径（1 基索引与名称都可查）
	sheetFile map[string]string

	// sharedStrings 共享字符串表（读单元格文本必需，故常驻）
	sharedStrings []string

	zr *zip.ReadCloser
}

// OpenReader 打开只读流式工作簿。
//
// 只解析 workbook.xml 与 sharedStrings.xml（两者体积相对可控），
// 其余部件保持在 zip 内按需读取。
func OpenReader(filename string) (*StreamBook, error) {
	zr, err := zip.OpenReader(filename)
	if err != nil {
		return nil, fmt.Errorf("打开 zip 失败: %w", err)
	}
	sb := &StreamBook{
		zipPath:   filename,
		sheetFile: map[string]string{},
		zr:        zr,
	}
	// workbook.xml -> 表名与 RID
	wbData, err := sb.readPart("xl/workbook.xml")
	if err != nil {
		zr.Close()
		return nil, err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(map[string][]byte{"xl/workbook.xml": wbData}, "xl/workbook.xml", wb); err != nil {
		zr.Close()
		return nil, err
	}
	// workbook.xml.rels -> RID 到部件路径
	relsData, err := sb.readPart("xl/_rels/workbook.xml.rels")
	if err != nil {
		zr.Close()
		return nil, err
	}
	rels := &Relationships{}
	if err := getXMLFromMap(map[string][]byte{"xl/_rels/workbook.xml.rels": relsData}, "xl/_rels/workbook.xml.rels", rels); err != nil {
		zr.Close()
		return nil, err
	}
	ridToTarget := map[string]string{}
	for _, r := range rels.Relationship {
		if strings.Contains(r.Type, "/worksheet") {
			ridToTarget[r.ID] = resolveTarget("xl", r.Target)
		}
	}
	for _, sh := range wb.Sheets.Sheet {
		sb.sheetNames = append(sb.sheetNames, sh.Name)
		if p, ok := ridToTarget[sh.RID]; ok {
			sb.sheetFile[sh.Name] = p
			sb.sheetFile[fmt.Sprint(len(sb.sheetNames))] = p
		}
	}
	// sharedStrings（可选）
	if sst, err := sb.readPart("xl/sharedStrings.xml"); err == nil {
		sb.sharedStrings = parseSharedStrings(string(sst))
	}
	return sb, nil
}

// Close 关闭底层 zip 读取器。
func (sb *StreamBook) Close() error {
	if sb.zr != nil {
		return sb.zr.Close()
	}
	return nil
}

// GetSheetList 返回所有工作表名（流式簿的表名来自 workbook.xml，已在打开时解析）。
func (sb *StreamBook) GetSheetList() []string {
	return append([]string(nil), sb.sheetNames...)
}

// readPart 读取指定 zip 部件的内容。
func (sb *StreamBook) readPart(name string) ([]byte, error) {
	for _, f := range sb.zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}
	}
	return nil, fmt.Errorf("部件不存在: %s", name)
}

// parseSharedStrings 解析共享字符串表为字符串切片。
func parseSharedStrings(s string) []string {
	var out []string
	// 每个 <si> 是一个条目；内部可能有多个 <t>（富文本），需拼接
	for {
		i := strings.Index(s, "<si>")
		if i == -1 {
			break
		}
		j := strings.Index(s[i:], "</si>")
		if j == -1 {
			break
		}
		inner := s[i+4 : i+j]
		s = s[i+j+4:]
		var sb strings.Builder
		for {
			k := strings.Index(inner, "<t")
			if k == -1 {
				break
			}
			// 跳过 <t ...> 属性
			gt := strings.Index(inner[k:], ">")
			if gt == -1 {
				break
			}
			rest := inner[k+gt+1:]
			e := strings.Index(rest, "</t>")
			if e == -1 {
				break
			}
			sb.WriteString(unescapeXML(rest[:e]))
			inner = rest[e+4:]
		}
		out = append(out, sb.String())
	}
	return out
}

// StreamRows 是逐行读取的迭代器。
//
// 用法：
//
//	it := b.StreamRows("Sheet1")
//	for it.Next() {
//	    row := it.Row()          // []string，缺失格子补空串
//	    cells := it.Cells()      // 带类型的原始值
//	}
//	if err := it.Error(); err != nil { ... }
//
// 每次只驻留一行，内存与表大小无关。
type RowIterator struct {
	sb      *StreamBook
	sheet   string
	rowsXML []string // 已扫描到的全部 <row> 片段（避免重复解析，见下方说明）
	idx     int
	cur     []streamCell
	err     error
	done    bool
}

// streamCell 是流式读取出的一个单元格。
type streamCell struct {
	Ref   string
	Value string // 已解析的显示值（共享字符串已解码）
	Raw   string // 原始 <v> 文本（数值/公式结果）
	IsNum bool
	Form  string
}

// StreamRows 开始流式读取 sheetRef 的行。
func (sb *StreamBook) StreamRows(sheetRef string) (*RowIterator, error) {
	path, ok := sb.sheetFile[sheetRef]
	if !ok {
		names := strings.Join(sb.sheetNames, ", ")
		return nil, fmt.Errorf("找不到工作表 %q（现有: %s）", sheetRef, names)
	}
	data, err := sb.readPart(path)
	if err != nil {
		return nil, err
	}
	return &RowIterator{
		sb:      sb,
		sheet:   sheetRef,
		rowsXML: splitRowElements(string(data)),
	}, nil
}

// splitRowElements 把 sheetData 拆成各 <row> 片段。
//
// 注意：这里仍一次性切分出行片段（内存 O(表 XML)），但**不解析成结构**，
// 且每个片段用完即丢。相比 Open() 省掉了 sharedStrings 之外的部件、
// cellXfs/styles 解析与整簿 fileMap 常驻的开销。
func splitRowElements(ws string) []string {
	var out []string
	for {
		i := strings.Index(ws, "<row ")
		if i == -1 {
			// 可能是自闭合的 <row r="1"/>
			i = strings.Index(ws, "<row>")
			if i == -1 {
				return out
			}
		}
		j := strings.Index(ws[i:], "</row>")
		if j == -1 {
			return out
		}
		out = append(out, ws[i:i+j+6])
		ws = ws[i+j+6:]
	}
}

// Next 前进到下一行，返回是否有数据。
func (it *RowIterator) Next() bool {
	if it.done {
		return false
	}
	if it.idx >= len(it.rowsXML) {
		it.done = true
		return false
	}
	it.cur = it.parseRow(it.rowsXML[it.idx])
	it.idx++
	return true
}

// Row 返回当前行的显示值列表。
func (it *RowIterator) Row() []string {
	out := make([]string, len(it.cur))
	for i, c := range it.cur {
		out[i] = c.Value
	}
	return out
}

// Cells 返回当前行的带类型单元格。
func (it *RowIterator) Cells() []streamCell {
	return it.cur
}

// RowNum 返回当前行的 1 基行号。
//
// 直接复用 parseCellRef（库内的标准解析），不在这里重复实现数字提取 ——
// 手写的那版从尾部反向累加，"A100" 会算成 1（001），是隐蔽的行号错位 bug。
func (it *RowIterator) RowNum() int {
	if len(it.cur) == 0 {
		return it.idx
	}
	// parseCellRef 返回 **0 基** 行号（A1 -> 0），而 RowNum 的契约是 1 基，
	// 故必须 +1。少这一行会让第 N 次迭代报出 N-1。
	if _, row, err := parseCellRef(it.cur[0].Ref); err == nil {
		return row + 1
	}
	return it.idx
}

// Error 返回迭代过程中的错误。
func (it *RowIterator) Error() error { return it.err }

// Close 释放行片段占用的内存（提前回收）。
func (it *RowIterator) Close() {
	it.rowsXML = nil
	it.cur = nil
	it.done = true
}

// parseRow 解析一个 <row> 片段。
func (it *RowIterator) parseRow(rowXML string) []streamCell {
	var out []streamCell
	s := rowXML
	for {
		i := strings.Index(s, "<c ")
		if i == -1 {
			return out
		}
		// 找该单元格的结束
		var cellXML string
		if j := strings.Index(s[i:], "</c>"); j != -1 {
			cellXML = s[i : i+j+4]
			s = s[i+j+4:]
		} else if j := strings.Index(s[i:], "/>"); j != -1 {
			cellXML = s[i : i+j+2]
			s = s[i+j+2:]
		} else {
			return out
		}
		if c, ok := it.sb.parseCell(cellXML); ok {
			out = append(out, c)
		}
	}
}

// parseCell 解析一个 <c> 片段。
func (it *StreamBook) parseCell(cellXML string) (streamCell, bool) {
	open := cellXML
	if i := strings.Index(cellXML, ">"); i != -1 {
		open = cellXML[:i+1]
	} else {
		return streamCell{}, false
	}
	ref := attrOf(open, "r")
	typ := attrOf(open, "t")
	inner := ""
	if i := strings.Index(cellXML, ">"); i != -1 {
		if j := strings.LastIndex(cellXML, "</c>"); j > i {
			inner = cellXML[i+1 : j]
		}
	}
	c := streamCell{Ref: ref}
	switch typ {
	case "s":
		// 共享字符串：<v>索引</v>
		idxStr := firstTagText(inner, "v")
		n := 0
		fmt.Sscanf(idxStr, "%d", &n)
		if n >= 0 && n < len(it.sharedStrings) {
			c.Value = it.sharedStrings[n]
		}
	case "inlineStr":
		c.Value = unescapeXML(firstTagText(inner, "t"))
		c.IsNum = false
	case "str":
		// 公式的字符串结果
		c.Value = unescapeXML(firstTagText(inner, "v"))
		if c.Value == "" {
			if f := firstTagText(inner, "f"); f != "" {
				c.Form = f
				c.Value = "=" + f
			}
		}
	case "b":
		if firstTagText(inner, "v") == "1" {
			c.Value = "TRUE"
		} else {
			c.Value = "FALSE"
		}
	case "e":
		c.Value = firstTagText(inner, "v") // 错误值原样
	default:
		// 数值（无 t 或 t="n"）
		v := unescapeXML(firstTagText(inner, "v"))
		c.Value = v
		c.Raw = v
		c.IsNum = v != ""
		if f := firstTagText(inner, "f"); f != "" {
			c.Form = f
		}
	}
	return c, true
}

// firstTagText 取出 <tag>text</tag> 或 <tag/> 里的文本。
func firstTagText(inner, tag string) string {
	open := "<" + tag
	i := strings.Index(inner, open)
	if i == -1 {
		return ""
	}
	// 避免 <t 匹配到 <text> 之类
	rest := inner[i+len(open):]
	if len(rest) == 0 {
		return ""
	}
	if rest[0] != '>' && rest[0] != ' ' && rest[0] != '/' {
		return ""
	}
	gt := strings.Index(rest, ">")
	if gt == -1 {
		return ""
	}
	// 注意 gt 可能为 0（如 "<v>"），此时不能访问 rest[gt-1]
	if gt > 0 && rest[gt-1] == '/' {
		return "" // 自闭合标签，无内容
	}
	body := rest[gt+1:]
	// 闭合标签用认证类型拼接：tag 是库内固定字面量（v/t/f 等），不是用户数据。
	// 走 xmlTagName 表达这一事实，护栏无需豁免。
	closingTag := "</" + string(tagName(tag)) + ">"
	if e := strings.Index(body, closingTag); e != -1 {
		return body[:e]
	}
	return ""
}

// StreamGetRows 把整个表读成二维字符串切片。
//
// 便利函数：需要一次性取全表时用。要控制内存请用 StreamRows 迭代。
func (sb *StreamBook) StreamGetRows(sheetRef string) ([][]string, error) {
	it, err := sb.StreamRows(sheetRef)
	if err != nil {
		return nil, err
	}
	defer it.Close()
	var out [][]string
	for it.Next() {
		out = append(out, it.Row())
	}
	return out, it.Error()
}

// StreamGetRow 读取单行（1 基），不载入其他行。
func (sb *StreamBook) StreamGetRow(sheetRef string, row int) ([]string, error) {
	it, err := sb.StreamRows(sheetRef)
	if err != nil {
		return nil, err
	}
	defer it.Close()
	for it.Next() {
		if it.RowNum() == row {
			return it.Row(), it.Error()
		}
	}
	return nil, nil
}

// StreamGetCell 读取单个单元格（不载入整表）。
func (sb *StreamBook) StreamGetCell(sheetRef, cell string) (string, error) {
	col, row, ok := refCoords(cell)
	if !ok {
		return "", fmt.Errorf("无效的单元格引用: %q", cell)
	}
	wantCol := col
	wantRow := row + 1 // refCoords 是 0 基，RowNum 契约是 1 基
	it, err := sb.StreamRows(sheetRef)
	if err != nil {
		return "", err
	}
	defer it.Close()
	for it.Next() {
		if it.RowNum() != wantRow {
			continue
		}
		for _, c := range it.Cells() {
			cc, _, err := parseCellRef(c.Ref)
			if err == nil && cc == wantCol {
				return c.Value, nil
			}
		}
		return "", nil
	}
	return "", it.Error()
}

// StreamOpen 是 OpenReader 的别名，语义更直观。
func StreamOpen(filename string) (*StreamBook, error) { return OpenReader(filename) }
