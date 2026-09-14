package excelgo

// book.go 提供面向对象的「文件 -> 工作表 -> 单元格」链式 API（excelgo.Book / Sheet / Cell / Range）。
//
// 设计原则与全库一致：纯 Go、不依赖 Office；所有修改基于内存中的 fileMap，最后由 Save/SaveAs
// 一次性写回，相较原有「每个全局函数各自读盘再写盘」的方式更省 IO，且便于在一个工作簿上
// 连续多次修改。原全局函数（InsertRows/SetCellValue/...）保持不变，仍可继续使用；
// 本文件的类型方法复用同一套底层字符串/正则精确定位原语，不破坏原有样式/布局/数据/图片。

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// ---------- Book：工作簿句柄 ----------

// Book 是面向对象工作簿句柄，持有内存 fileMap 与文件名（未保存时文件名为空）。
type Book struct {
	filename string
	fileMap  map[string][]byte
}

// Open 打开一个已有 xlsx 文件，读入内存。
func Open(filename string) (*Book, error) {
	fm, err := readZipToMap(filename)
	if err != nil {
		return nil, err
	}
	return &Book{filename: filename, fileMap: fm}, nil
}

// Create 新建一个空白工作簿（含一个默认 Sheet1），尚未落盘，需 SaveAs 写文件。
func Create() (*Book, error) {
	fm, err := blankWorkbookMap()
	if err != nil {
		return nil, err
	}
	return &Book{fileMap: fm}, nil
}

// Save 写回原文件。若由 Create 创建尚未指定文件名，应改用 SaveAs。
func (b *Book) Save() error {
	if b.filename == "" {
		return fmt.Errorf("尚未指定文件名，请使用 SaveAs")
	}
	return writeMapToZip(b.filename, b.fileMap)
}

// SaveAs 以新文件名写盘（同时更新 Book.filename，之后 Save 沿用此名）。
func (b *Book) SaveAs(filename string) error {
	if err := writeMapToZip(filename, b.fileMap); err != nil {
		return err
	}
	b.filename = filename
	return nil
}

// Filename 返回当前文件名（可能为空，表示尚未落盘）。
func (b *Book) Filename() string { return b.filename }

// SheetNames 返回所有工作表名称（按 workbook.xml 顺序）。
func (b *Book) SheetNames() []string {
	wb := b.wbStruct()
	names := make([]string, 0, len(wb.Sheets.Sheet))
	for _, s := range wb.Sheets.Sheet {
		names = append(names, s.Name)
	}
	return names
}

// Sheet 获取名为 name 的工作表（name 可为名称或 1 基索引字符串）。不存在返回错误。
func (b *Book) Sheet(name string) (*WorkSheet, error) {
	file, _, _, _, err := locateSheetInMap(b.fileMap, b.wbStruct(), name)
	if err != nil {
		return nil, err
	}
	return &WorkSheet{book: b, ref: name, file: file}, nil
}

// MustSheet 同 Sheet，但不存在时按 name 新建并返回（容错便捷方法）。
func (b *Book) MustSheet(name string) (*WorkSheet, error) {
	s, err := b.Sheet(name)
	if err == nil {
		return s, nil
	}
	if _, e := b.AddSheet(name); e != nil {
		return nil, e
	}
	return b.Sheet(name)
}

// AddSheet 在工作簿末尾新建名为 name 的空白工作表，返回该 Sheet 句柄。
func (b *Book) AddSheet(name string) (*WorkSheet, error) {
	file, _, err := newSheetInMap(b.fileMap, name)
	if err != nil {
		return nil, err
	}
	return &WorkSheet{book: b, ref: name, file: file}, nil
}

// RemoveSheet 删除名为 name 的工作表（至少保留一个）。
func (b *Book) RemoveSheet(name string) error {
	return deleteSheetInMap(b.fileMap, name)
}

// RenameSheet 把工作表 oldName 改名为 newName（同名自动避让）。
func (b *Book) RenameSheet(oldName, newName string) error {
	b.fileMap["xl/workbook.xml"] = renameSheetInWorkbookXML(b.fileMap["xl/workbook.xml"], oldName, newName)
	return nil
}

// MoveSheet 把 name 工作表移动到 1 基序号 toIndex（1 表示最前）。
func (b *Book) MoveSheet(name string, toIndex int) error {
	_, _, idx, _, err := locateSheetInMap(b.fileMap, b.wbStruct(), name)
	if err != nil {
		return err
	}
	b.fileMap["xl/workbook.xml"] = moveSheetInWorkbookXML(b.fileMap["xl/workbook.xml"], idx, toIndex-1)
	return nil
}

// wbStruct 把当前 workbook.xml 解析为 Workbook 结构（仅供定位/列举）。
func (b *Book) wbStruct() *Workbook {
	wb := &Workbook{}
	_ = getXMLFromMap(b.fileMap, "xl/workbook.xml", wb)
	return wb
}

// blankWorkbookMap 构造最小合法的空白 xlsx 部件集合（供 Create 使用）。
func blankWorkbookMap() (map[string][]byte, error) {
	const (
		ct = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
			`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
			`<Default Extension="xml" ContentType="application/xml"/>` +
			`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>` +
			`<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>` +
			`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>` +
			`<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>` +
			`</Types>`
		rootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
			`</Relationships>`
		wb = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
			`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
			`<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`
		xlRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
			`<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>` +
			`<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>` +
			`</Relationships>`
		ss = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"/>`
	)
	return map[string][]byte{
		"[Content_Types].xml":        []byte(ct),
		"_rels/.rels":                []byte(rootRels),
		"xl/workbook.xml":            []byte(wb),
		"xl/_rels/workbook.xml.rels": []byte(xlRels),
		"xl/worksheets/sheet1.xml":   []byte(blankWorksheet()),
		"xl/styles.xml":              []byte(defaultStylesXML()),
		"xl/sharedStrings.xml":       []byte(ss),
	}, nil
}

// ---------- Sheet：工作表句柄 ----------

// Sheet 是面向对象工作表句柄。
type WorkSheet struct {
	book *Book
	ref  string // 名称或 1 基索引字符串（创建时即名称）
	file string // worksheet XML 在 fileMap 中的路径
}

func (s *WorkSheet) fm() map[string][]byte { return s.book.fileMap }
func (s *WorkSheet) ws() string            { return string(s.book.fileMap[s.file]) }
func (s *WorkSheet) setWS(x string)        { s.book.fileMap[s.file] = []byte(x) }

// setTabColor 在工作簿 workbook.xml 的 <sheet name="..."> 上设置/移除标签色（ARGB）。
// 空 rgb 表示移除已有标签色。标签色属于 workbook.xml 的 sheet 节点，而非 worksheet.xml。
func (s *WorkSheet) setTabColor(rgb string) {
	wb := s.book.wbStruct()
	_, name, _, _, err := locateSheetInMap(s.fm(), wb, s.ref)
	if err != nil {
		return
	}
	xml := string(s.fm()["xl/workbook.xml"])
	// 匹配该 sheet 的 <sheet ... name="name" ...> 起始标签
	re := regexp.MustCompile(`(?s)<sheet\b[^>]*\bname="` + regexp.QuoteMeta(name) + `"[^>]*>`)
	m := re.FindStringSubmatch(xml)
	if m == nil {
		return
	}
	tag := m[0]
	newTag := regexp.MustCompile(`\s+tabColor[^>/]*`).ReplaceAllString(tag, "")
	if rgb != "" {
		// 始终输出自闭合形式 <sheet ... tabColor="..."/>，避免未闭合的开标签
		if strings.HasSuffix(newTag, "/>") {
			newTag = newTag[:len(newTag)-2] + ` tabColor="` + escapeAttr(rgb) + `"/>`
		} else {
			newTag = newTag[:len(newTag)-1] + ` tabColor="` + escapeAttr(rgb) + `"/>`
		}
	}
	xml = strings.Replace(xml, tag, newTag, 1)
	s.fm()["xl/workbook.xml"] = []byte(xml)
}

// Name 返回该工作表名称。
func (s *WorkSheet) Name() string {
	_, name, _, _, err := locateSheetInMap(s.fm(), s.book.wbStruct(), s.ref)
	if err != nil {
		return s.ref
	}
	return name
}

// Cell 获取 ref 单元格句柄（如 "A1"）。
func (s *WorkSheet) Cell(ref string) *Cell { return &Cell{sheet: s, ref: ref} }

// Range 获取 ref 矩形区域句柄（如 "A1:C10"）。
func (s *WorkSheet) Range(ref string) *Range { return &Range{sheet: s, ref: ref} }

// NewSheet 在当前工作簿内新建名为 name 的工作表，返回其句柄。
func (b *Book) NewSheet(name string) (*WorkSheet, error) {
	file, _, err := newSheetInMap(b.fileMap, name)
	if err != nil {
		return nil, err
	}
	return &WorkSheet{book: b, ref: name, file: file}, nil
}

// ---------- 行列插入/删除（复用底层平移原语，公式引用自动平移） ----------

// InsertRows 在第 row 行（1 基）前插入 n 行。
func (s *WorkSheet) InsertRows(row, n int) error {
	if row < 1 || n < 1 {
		return fmt.Errorf("row 须>=1 且 n 须>=1")
	}
	ws, err := shiftRows(s.ws(), row, n, nil)
	if err != nil {
		return err
	}
	s.setWS(insertEmptyRows(ws, row, n))
	return nil
}

// RemoveRows 删除第 row 行起共 n 行。
func (s *WorkSheet) RemoveRows(row, n int) error {
	if row < 1 || n < 1 {
		return fmt.Errorf("row 须>=1 且 n 须>=1")
	}
	ws := deleteRowRange(s.ws(), row, n)
	delRows := [2]int{row, row + n - 1}
	ws2, err := shiftRows(ws, row+n, -n, &delRows)
	if err != nil {
		return err
	}
	s.setWS(ws2)
	return nil
}

// InsertCols 在第 col 列（1 基）前插入 n 列。
func (s *WorkSheet) InsertCols(col, n int) error {
	if col < 1 || n < 1 {
		return fmt.Errorf("col 须>=1 且 n 须>=1")
	}
	s.setWS(shiftCols(s.ws(), col, n, nil))
	return nil
}

// RemoveCols 删除第 col 列起共 n 列。
func (s *WorkSheet) RemoveCols(col, n int) error {
	if col < 1 || n < 1 {
		return fmt.Errorf("col 须>=1 且 n 须>=1")
	}
	delCols := [2]int{col, col + n - 1}
	s.setWS(shiftCols(s.ws(), col, -n, &delCols))
	return nil
}

// ---------- 单元格读写 ----------

// GetCell 读取单元格显示值（公式取结果）。
func (s *WorkSheet) GetCell(ref string) string {
	return readCellValue(s.fm(), s.ws(), ref)
}

// GetCellFormula 读取单元格公式（不含前导 =），非公式单元格返回 ""。
func (s *WorkSheet) GetCellFormula(ref string) string {
	ws := s.ws()
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + regexp.QuoteMeta(ref) + `"[^>]*?(?:/>|>(.*?)</c>)`)
	m := re.FindStringSubmatch(ws)
	if m == nil {
		return ""
	}
	if strings.Contains(m[0], "<f") {
		fm := regexp.MustCompile(`(?s)<f[^>]*>(.*?)</f>`).FindStringSubmatch(m[1])
		if fm != nil {
			return strings.TrimPrefix(fm[1], "=")
		}
	}
	return ""
}

// GetCellStyleIndex 读取单元格 s 样式索引（无则返回 0）。
func (s *WorkSheet) GetCellStyleIndex(ref string) int {
	ws := s.ws()
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + regexp.QuoteMeta(ref) + `"[^>]*`)
	m := re.FindStringSubmatch(ws)
	if m == nil {
		return 0
	}
	sm := regexp.MustCompile(`\bs="(\d+)"`).FindStringSubmatch(m[0])
	if sm == nil {
		return 0
	}
	n, _ := strconv.Atoi(sm[1])
	return n
}

// GetCellType 读取单元格类型（string/number/bool/formula/empty）。
func (s *WorkSheet) GetCellType(ref string) string {
	ws := s.ws()
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + regexp.QuoteMeta(ref) + `"[^>]*?(?:/>|>(.*?)</c>)`)
	m := re.FindStringSubmatch(ws)
	if m == nil {
		return "empty"
	}
	tag := m[0]
	if strings.Contains(tag, "<f") {
		return "formula"
	}
	if strings.Contains(tag, `t="s"`) {
		return "string"
	}
	if strings.Contains(tag, `t="b"`) {
		return "bool"
	}
	if strings.Contains(tag, "<v") || strings.Contains(tag, "<is") {
		return "number"
	}
	return "empty"
}

// SetCellValue 按值类型自动写入（string→共享字符串，bool→布尔，int/float64→数值）。
func (s *WorkSheet) SetCellValue(ref string, value interface{}) error {
	switch v := value.(type) {
	case string:
		return s.SetCellStr(ref, v)
	case bool:
		return s.SetCellBool(ref, v)
	case int:
		return s.SetCellInt(ref, v)
	case int64:
		return s.SetCellInt(ref, int(v))
	case float64:
		return s.SetCellNumeric(ref, v)
	case float32:
		return s.SetCellNumeric(ref, float64(v))
	default:
		return fmt.Errorf("不支持的值类型: %T", value)
	}
}

// SetCellStr 写入共享字符串。
func (s *WorkSheet) SetCellStr(ref, val string) error {
	return setCellInMap(s.fm(), s.file, ref, CellTypeString, val, "")
}

// SetCellInline 写入内联字符串。
func (s *WorkSheet) SetCellInline(ref, val string) error {
	return setCellInMap(s.fm(), s.file, ref, CellTypeInline, val, "")
}

// SetCellInt 写入整数数值。
func (s *WorkSheet) SetCellInt(ref string, val int) error {
	return setCellInMap(s.fm(), s.file, ref, CellTypeNumeric, strconv.Itoa(val), "")
}

// SetCellNumeric 写入浮点数值。
func (s *WorkSheet) SetCellNumeric(ref string, val float64) error {
	return setCellInMap(s.fm(), s.file, ref, CellTypeNumeric, strconv.FormatFloat(val, 'f', -1, 64), "")
}

// SetCellBool 写入布尔值。
func (s *WorkSheet) SetCellBool(ref string, val bool) error {
	b := "0"
	if val {
		b = "1"
	}
	return setCellInMap(s.fm(), s.file, ref, CellTypeBool, b, "")
}

// SetCellFormula 写入公式（可选 result 作为预计算缓存值）。
func (s *WorkSheet) SetCellFormula(ref, formula, result string) error {
	return setCellInMap(s.fm(), s.file, ref, CellTypeFormula, formula, result)
}

// GetRange 读取矩形区域的值（二维 [][]string）。
func (s *WorkSheet) GetRange(ref string) ([][]string, error) {
	c1, r1, c2, r2, err := parseRangeRef(ref)
	if err != nil {
		return nil, err
	}
	return extractRangeRows(s.ws(), s.fm(), c1, r1, c2, r2), nil
}

// SetRange 把二维值写入矩形区域（左上角对齐，单格起点自动扩展）。
func (s *WorkSheet) SetRange(ref string, values [][]interface{}) error {
	return setRangeInMap(s.fm(), s.file, ref, values)
}

// SetStyle 把样式应用到单个单元格，返回 s 索引。
func (s *WorkSheet) SetStyle(ref string, st Style) (int, error) {
	return setStyleRangeInMap(s.fm(), s.file, ref, st)
}

// SetStyleRange 把样式应用到矩形区域，返回 s 索引。
func (s *WorkSheet) SetStyleRange(ref string, st Style) (int, error) {
	return setStyleRangeInMap(s.fm(), s.file, ref, st)
}

// ---------- Cell：单元格句柄 ----------

// Cell 是面向对象单元格句柄。
type Cell struct {
	sheet *WorkSheet
	ref   string
}

// Get 读取显示值。
func (c *Cell) Get() string { return c.sheet.GetCell(c.ref) }

// Set 写入值（自动推断类型）。
func (c *Cell) Set(value interface{}) error { return c.sheet.SetCellValue(c.ref, value) }

// SetStr 写入共享字符串。
func (c *Cell) SetStr(v string) error { return c.sheet.SetCellStr(c.ref, v) }

// SetFormula 写入公式。
func (c *Cell) SetFormula(formula, result string) error {
	return c.sheet.SetCellFormula(c.ref, formula, result)
}

// SetStyle 设置本单元格样式，返回 s 索引。
func (c *Cell) SetStyle(st Style) (int, error) { return c.sheet.SetStyle(c.ref, st) }

// GetStr 读取单元格字符串值。
func (c *Cell) GetStr() string { return c.sheet.GetCell(c.ref) }

// GetNum 读取单元格数值（无法解析时返回 0）。
func (c *Cell) GetNum() float64 {
	v := c.sheet.GetCell(c.ref)
	f, _ := strconv.ParseFloat(v, 64)
	return f
}

// GetFormula 读取单元格公式（不含前导 =）。
func (c *Cell) GetFormula() string { return c.sheet.GetCellFormula(c.ref) }

// GetStyle 读取单元格样式索引。
func (c *Cell) GetStyle() (int, error) { return c.sheet.GetCellStyleIndex(c.ref), nil }

// MergeTo 把本单元格与 toRef 合并为一块合并区域。
func (c *Cell) MergeTo(toRef string) error { return c.sheet.MergeCells(c.ref + ":" + toRef) }

// ---------- Range：区域句柄 ----------

// Range 是面向对象区域句柄。
type Range struct {
	sheet *WorkSheet
	ref   string
}

// Get 读取区域值。
func (r *Range) Get() ([][]string, error) { return r.sheet.GetRange(r.ref) }

// Set 写入区域值。
func (r *Range) Set(values [][]interface{}) error { return r.sheet.SetRange(r.ref, values) }

// SetStyle 设置区域样式，返回 s 索引。
func (r *Range) SetStyle(st Style) (int, error) { return r.sheet.SetStyleRange(r.ref, st) }

// Merge 把本区域合并为一块合并单元格。
func (r *Range) Merge() error { return r.sheet.MergeCells(r.ref) }
