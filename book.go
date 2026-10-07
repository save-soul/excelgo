package excelgo

// book.go 提供面向对象的「文件 -> 工作表 -> 单元格」链式 API（excelgo.Book / Sheet / Cell / Range）。
//
// 设计原则与全库一致：纯 Go、不依赖 Office；所有修改基于内存中的 fileMap，最后由 Save/SaveAs
// 一次性写回，相较原有「每个全局函数各自读盘再写盘」的方式更省 IO，且便于在一个工作簿上
// 连续多次修改。原全局函数（InsertRows/SetCellValue/...）保持不变，仍可继续使用；
// 本文件的类型方法复用同一套底层字符串/正则精确定位原语，不破坏原有样式/布局/数据/图片。

import (
	"bytes"
	"fmt"
	"math"
	"regexp"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

// ---------- Book：工作簿句柄 ----------

// Book 是面向对象工作簿句柄，持有内存 fileMap 与文件名（未保存时文件名为空）。
//
// 命名上与 excelize 的 *File 对应，故另提供类型别名 File = Book，
// 调用方可按习惯写成 *excelgo.File。Open/Create 返回 *Book（亦即 *File）。
type Book struct {
	filename string
	fileMap  map[string][]byte
	// sstCache 是 sharedStrings.xml 的解析索引缓存，仅作写入加速。
	// 以部件字节为指纹，外部改动该部件会自动使其失效，故不影响正确性。
	// 生命周期与本 Book 一致（打开的文件已含 SST 时才会在首次写入时填充）。
	sstCache *sharedStringsCache
}

// sst 返回本工作簿的共享字符串缓存（惰性创建）。
func (b *Book) sst() *sharedStringsCache {
	if b.sstCache == nil {
		b.sstCache = &sharedStringsCache{}
	}
	return b.sstCache
}

// File 即 Book 的别名，与 excelize 的 *File 命名保持一致，便于从 excelize 迁移。
type File = Book

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

// maxSheetNameLen 为 Excel 规定的工作表名长度上限（字符数）。
const maxSheetNameLen = 31

// validateSheetName 校验工作表名是否合法。Excel 限制表名最长 31 个字符，
// 超长时 Excel/WPS 会直接拒绝打开文件；此处提前拦截，避免产出无法打开的文件。
// 另拒绝空名与含非法字符（\ / ? * [ ] :）的名字。
func validateSheetName(name string) error {
	if name == "" {
		return fmt.Errorf("工作表名不能为空")
	}
	if n := utf8.RuneCountInString(name); n > maxSheetNameLen {
		return fmt.Errorf("工作表名 %q 长度 %d 超过 Excel 上限 %d 个字符", name, n, maxSheetNameLen)
	}
	if strings.ContainsAny(name, `\/?*[]:`) {
		return fmt.Errorf("工作表名 %q 含 Excel 不允许的字符（\\ / ? * [ ] :）", name)
	}
	return nil
}

// RenameSheet 把工作表 oldName 改名为 newName（同名自动避让）。
// 若 oldName 不存在，返回错误而非静默无操作。newName 超长或含非法字符时返回错误。
func (b *Book) RenameSheet(oldName, newName string) error {
	if err := validateSheetName(newName); err != nil {
		return err
	}
	before := b.fileMap["xl/workbook.xml"]
	after := renameSheetInWorkbookXML(before, oldName, newName)
	if bytes.Equal(before, after) {
		return fmt.Errorf("工作表 %q 不存在，无法重命名", oldName)
	}
	// 同步改掉 definedName 里对该表的引用（打印区域/重复打印标题等）。
	// 漏掉这步会让改完名后 definedName 悬空指向旧表名，Excel 修复时静默丢弃打印设置。
	after = renameDefinedNamesInWorkbookXML(after, oldName, newName)
	b.fileMap["xl/workbook.xml"] = after
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

// ---------- 工作表可见性（visible/hidden/veryHidden） ----------

// 工作表可见性状态常量。
const (
	SheetStateVisible    = "visible"
	SheetStateHidden     = "hidden"
	SheetStateVeryHidden = "veryHidden"
)

// SetSheetVisible 设置工作表 name 的可见性状态：
//   - SheetStateVisible    → 正常显示（移除 state 属性）
//   - SheetStateHidden     → 可经 Excel UI 取消隐藏
//   - SheetStateVeryHidden → 仅能经 VBA/代码取消隐藏
func (b *Book) SetSheetVisible(name, state string) error {
	xml := string(b.fileMap["xl/workbook.xml"])
	re := regexp.MustCompile(`(?s)<sheet\b[^>]*\bname="` + regexp.QuoteMeta(name) + `"[^>]*>`)
	m := re.FindStringSubmatch(xml)
	if m == nil {
		return fmt.Errorf("工作表 %q 不存在", name)
	}
	tag := m[0]
	newTag := regexp.MustCompile(`\s+state="[^"]*"`).ReplaceAllString(tag, "")
	if state != SheetStateVisible && state != "" {
		if strings.HasSuffix(newTag, "/>") {
			newTag = newTag[:len(newTag)-2] + ` state="` + safeAttr(state) + `"/>`
		} else {
			newTag = newTag[:len(newTag)-1] + ` state="` + safeAttr(state) + `">`
		}
	}
	xml = strings.Replace(xml, tag, newTag, 1)
	b.fileMap["xl/workbook.xml"] = []byte(xml)
	return nil
}

// GetSheetVisible 读取工作表 name 的可见性状态（默认 SheetStateVisible）。
func (b *Book) GetSheetVisible(name string) (string, error) {
	xml := string(b.fileMap["xl/workbook.xml"])
	re := regexp.MustCompile(`(?s)<sheet\b[^>]*\bname="` + regexp.QuoteMeta(name) + `"[^>]*>`)
	m := re.FindStringSubmatch(xml)
	if m == nil {
		return "", fmt.Errorf("工作表 %q 不存在", name)
	}
	st := attrOf(m[0], "state")
	if st == "" {
		return SheetStateVisible, nil
	}
	return st, nil
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

// ws 返回工作表部件的原始 XML 字节。类型为 partXML —— 护栏据此区分
// 「复用既有 XML」（安全）与「把用户数据写进 XML」（需转义）。
func (s *WorkSheet) ws() partXML    { return partXML(s.book.fileMap[s.file]) }
func (s *WorkSheet) setWS(x string) { s.book.fileMap[s.file] = []byte(x) }

// setTabColor 在工作簿 workbook.xml 的 <sheet name="..."> 上设置/移除标签色（ARGB）。
// 空 rgb 表示移除已有标签色。标签色属于 workbook.xml 的 sheet 节点，而非 worksheet.xml。
// setTabColor 设置工作表标签颜色（符合 OOXML 规范：写入 worksheet XML 的
// <sheetPr><tabColor rgb="..."/>，而非 workbook.xml <sheet> 的属性——后者非标准，
// 会被严格解析器（如 openpyxl）拒绝）。
func (s *WorkSheet) setTabColor(rgb string) {
	ws := string(s.ws())
	ws = setTabColorInWS(ws, rgb)
	s.setWS(ws)
}

// setTabColorInWS 在 worksheet XML 中设置/清除 <sheetPr><tabColor>。
func setTabColorInWS(ws, rgb string) string {
	// 先移除已有的 <tabColor> 子元素
	ws = regexp.MustCompile(`(?s)<tabColor\b[^>]*/?>`).ReplaceAllString(ws, "")
	if rgb == "" {
		return ws
	}
	tabColorTag := `<tabColor rgb="` + safeAttr(rgb) + `"/>`
	if m := regexp.MustCompile(`(?s)(<sheetPr\b[^>]*>)`).FindStringSubmatch(ws); m != nil {
		// 在 <sheetPr> 起始标签之后插入 <tabColor>
		tag := m[0]
		return strings.Replace(ws, tag, tag+tabColorTag, 1)
	}
	// 没有 <sheetPr>：在 <worksheet ...> 之后（或 <sheetData 之前）创建 <sheetPr>
	block := `<sheetPr>` + tabColorTag + `</sheetPr>`
	if i := strings.Index(ws, "<sheetData"); i != -1 {
		return ws[:i] + block + ws[i:]
	}
	if i := strings.Index(ws, "</worksheet>"); i != -1 {
		return ws[:i] + block + ws[i:]
	}
	return ws + block
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
	cur := string(s.ws())
	if mr := wsMaxRow(cur); row > mr+1 {
		return fmt.Errorf("插入行 %d 超出表尾（当前最大行 %d，最多可在 %d 处插入）", row, mr, mr+1)
	}
	ws, err := shiftRows(cur, row, n, nil)
	if err != nil {
		return err
	}
	s.setWS(insertEmptyRows(ws, row, n))
	// 浮动图片锚点随行下移（与 Excel 行为一致）
	shiftDrawingAnchorsForSheet(s.book.fileMap, s.file, row, n, nil, true)
	return nil
}

// RemoveRows 删除第 row 行起共 n 行。
func (s *WorkSheet) RemoveRows(row, n int) error {
	if row < 1 || n < 1 {
		return fmt.Errorf("row 须>=1 且 n 须>=1")
	}
	cur := string(s.ws())
	if mr := wsMaxRow(cur); row+n-1 > mr {
		return fmt.Errorf("删除第 %d~%d 行超出表尾（当前最大行 %d）", row, row+n-1, mr)
	}
	ws := deleteRowRange(cur, row, n)
	delRows := [2]int{row, row + n - 1}
	ws2, err := shiftRows(ws, row+n, -n, &delRows)
	if err != nil {
		return err
	}
	s.setWS(ws2)
	// 浮动图片锚点随行上移；落入被删行区间的图片一并移除（与 Excel 行为一致）
	shiftDrawingAnchorsForSheet(s.book.fileMap, s.file, row, -n, &delRows, true)
	return nil
}

// InsertCols 在第 col 列（1 基）前插入 n 列。
func (s *WorkSheet) InsertCols(col, n int) error {
	if col < 1 || n < 1 {
		return fmt.Errorf("col 须>=1 且 n 须>=1")
	}
	cur := string(s.ws())
	if mc := wsMaxCol(cur); col > mc+1 {
		return fmt.Errorf("插入列 %d 超出表尾（当前最大列 %d，最多可在 %d 处插入）", col, mc, mc+1)
	}
	s.setWS(shiftCols(cur, col, n, nil))
	// 浮动图片锚点随列右移（与 Excel 行为一致）
	shiftDrawingAnchorsForSheet(s.book.fileMap, s.file, col, n, nil, false)
	return nil
}

// RemoveCols 删除第 col 列起共 n 列。
func (s *WorkSheet) RemoveCols(col, n int) error {
	if col < 1 || n < 1 {
		return fmt.Errorf("col 须>=1 且 n 须>=1")
	}
	cur := string(s.ws())
	if mc := wsMaxCol(cur); col+n-1 > mc {
		return fmt.Errorf("删除第 %d~%d 列超出表尾（当前最大列 %d）", col, col+n-1, mc)
	}
	delCols := [2]int{col, col + n - 1}
	s.setWS(shiftCols(cur, col, -n, &delCols))
	// 浮动图片锚点随列左移；落入被删列区间的图片一并移除（与 Excel 行为一致）
	shiftDrawingAnchorsForSheet(s.book.fileMap, s.file, col, -n, &delCols, false)
	return nil
}

// ---------- 单元格读写 ----------

// GetCell 读取单元格显示值（公式取结果）。
func (s *WorkSheet) GetCell(ref string) string {
	return readCellValue(s.fm(), string(s.ws()), ref)
}

// GetCellFormula 读取单元格公式（不含前导 =），非公式单元格返回 ""。
func (s *WorkSheet) GetCellFormula(ref string) string {
	ws := string(s.ws())
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + regexp.QuoteMeta(ref) + `"[^>]*?(?:/>|>(.*?)</c>)`)
	m := re.FindStringSubmatch(ws)
	if m == nil {
		return ""
	}
	if strings.Contains(m[0], "<f") {
		fm := regexp.MustCompile(`(?s)<f[^>]*>(.*?)</f>`).FindStringSubmatch(m[1])
		if fm != nil {
			// 公式在 XML 中以转义形式存储（如 &amp;），读出时需还原
			return unescapeXML(strings.TrimPrefix(fm[1], "="))
		}
	}
	return ""
}

// GetCellStyleIndex 读取单元格 s 样式索引（无则返回 0）。
func (s *WorkSheet) GetCellStyleIndex(ref string) int {
	ws := string(s.ws())
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

// GetCellValue 读取单元格值并按其类型返回原生 Go 值：
//   - 空单元格 → nil
//   - 共享/内联字符串、公式字符串 → string
//   - 布尔 t="b" → bool
//   - 数值（含公式缓存值）→ int（整数外形）或 float64
//   - 公式 → 有缓存值则按其类型返回，否则返回 "=" 前缀的公式文本
func (s *WorkSheet) GetCellValue(ref string) (interface{}, error) {
	ws := string(s.ws())
	re := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + regexp.QuoteMeta(ref) + `"[^>]*/?>(.*?)</c>`)
	m := re.FindStringSubmatch(ws)
	if m == nil {
		if regexp.MustCompile(`<c\b[^>]*\br="` + regexp.QuoteMeta(ref) + `"[^>]*/>`).MatchString(ws) {
			return nil, nil
		}
		return nil, nil
	}
	cellTag := extractCellTag(ws, ref)
	inner := m[1]
	t := attrOf(cellTag, "t")

	if strings.Contains(inner, "<f") {
		if v := regexp.MustCompile(`(?s)<v>(.*?)</v>`).FindStringSubmatch(inner); v != nil {
			return typedValue(v[1], t)
		}
		if f := regexp.MustCompile(`(?s)<f>(.*?)</f>`).FindStringSubmatch(inner); f != nil {
			return "=" + f[1], nil
		}
		return nil, nil
	}

	switch t {
	case "s":
		if v := regexp.MustCompile(`(?s)<v>(.*?)</v>`).FindStringSubmatch(inner); v != nil {
			idx, err := strconv.Atoi(v[1])
			if err == nil {
				if ss, ok := s.fm()["xl/sharedStrings.xml"]; ok {
					items := extractSharedStrings(string(ss))
					if idx >= 0 && idx < len(items) {
						return items[idx], nil
					}
				}
			}
		}
		return nil, nil
	case "inlineStr", "str":
		if t2 := regexp.MustCompile(`(?s)<t\b[^>]*>(.*?)</t>`).FindStringSubmatch(inner); t2 != nil {
			return unescapeXML(t2[1]), nil
		}
		return nil, nil
	case "b":
		if v := regexp.MustCompile(`(?s)<v>(.*?)</v>`).FindStringSubmatch(inner); v != nil {
			return v[1] == "1", nil
		}
		return false, nil
	default:
		if v := regexp.MustCompile(`(?s)<v>(.*?)</v>`).FindStringSubmatch(inner); v != nil {
			return typedValue(v[1], t)
		}
		return nil, nil
	}
}

// GetCellType 读取单元格类型（string/number/bool/formula/empty）。
func (s *WorkSheet) GetCellType(ref string) string {
	ws := string(s.ws())
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
	case nil:
		return nil // 空值：跳过该格（与 openpyxl 的 None 一致）
	case string:
		return s.SetCellStr(ref, v)
	case bool:
		return s.SetCellBool(ref, v)
	case int:
		return s.SetCellInt(ref, v)
	case int8:
		return s.SetCellInt(ref, int(v))
	case int16:
		return s.SetCellInt(ref, int(v))
	case int32:
		return s.SetCellInt(ref, int(v))
	case int64:
		return s.SetCellInt(ref, int(v))
	case uint:
		return s.SetCellNumeric(ref, float64(v))
	case uint8:
		return s.SetCellInt(ref, int(v))
	case uint16:
		return s.SetCellInt(ref, int(v))
	case uint32:
		return s.SetCellInt(ref, int(v))
	case uint64:
		return s.SetCellNumeric(ref, float64(v))
	case float64:
		return s.SetCellNumeric(ref, v)
	case float32:
		return s.SetCellNumeric(ref, float64(v))
	case time.Time:
		return s.SetCellTime(ref, v)
	default:
		return fmt.Errorf("不支持的值类型: %T（可用：nil/bool/int 族/uint 族/float/string/time.Time）", value)
	}
}

// AppendRow 在工作表末尾追加一行，元素类型与 SetCellValue 相同。
//
// 从 A 列开始按序填充；nil 元素跳过（留空）。这是写大表时最顺手的入口 ——
// 逐格调 SetCellValue 需要调用方自己算坐标。
func (s *WorkSheet) AppendRow(values []interface{}) error {
	rowNum := s.rowsUsed() + 1
	for i, v := range values {
		if v == nil {
			continue
		}
		if err := s.SetCellValue(refOf(rowNum-1, i), v); err != nil {
			return fmt.Errorf("追加第 %d 行第 %d 列失败: %w", rowNum, i+1, err)
		}
	}
	return nil
}

// rowsUsed 返回已使用的最大行号（无数据时为 0）。
func (s *WorkSheet) rowsUsed() int {
	maxRow := 0
	for _, m := range regexp.MustCompile(`<row r="(\d+)"`).
		FindAllStringSubmatch(string(s.ws()), -1) {
		if n, err := strconv.Atoi(m[1]); err == nil && n > maxRow {
			maxRow = n
		}
	}
	return maxRow
}

// SetCellStr 写入共享字符串。
func (s *WorkSheet) SetCellStr(ref, val string) error {
	return setCellInMap(s.fm(), s.file, ref, CellTypeString, val, "", s.book.sst())
}

// SetCellInline 写入内联字符串。
func (s *WorkSheet) SetCellInline(ref, val string) error {
	return setCellInMap(s.fm(), s.file, ref, CellTypeInline, val, "", s.book.sst())
}

// SetCellInt 写入整数数值。
func (s *WorkSheet) SetCellInt(ref string, val int) error {
	return setCellInMap(s.fm(), s.file, ref, CellTypeNumeric, strconv.Itoa(val), "", s.book.sst())
}

// SetCellNumeric 写入浮点数值。
func (s *WorkSheet) SetCellNumeric(ref string, val float64) error {
	return setCellInMap(s.fm(), s.file, ref, CellTypeNumeric, strconv.FormatFloat(val, 'f', -1, 64), "", s.book.sst())
}

// SetCellBool 写入布尔值。
func (s *WorkSheet) SetCellBool(ref string, val bool) error {
	b := "0"
	if val {
		b = "1"
	}
	return setCellInMap(s.fm(), s.file, ref, CellTypeBool, b, "", s.book.sst())
}

// SetCellFormula 写入公式（可选 result 作为预计算缓存值）。
func (s *WorkSheet) SetCellFormula(ref, formula, result string) error {
	return setCellInMap(s.fm(), s.file, ref, CellTypeFormula, formula, result, s.book.sst())
}

// GetRange 读取矩形区域的值（二维 [][]string）。
func (s *WorkSheet) GetRange(ref string) ([][]string, error) {
	c1, r1, c2, r2, err := parseRangeRef(ref)
	if err != nil {
		return nil, err
	}
	return extractRangeRows(string(s.ws()), s.fm(), c1, r1, c2, r2), nil
}

// SetRange 把二维值写入矩形区域（左上角对齐，单格起点自动扩展）。
func (s *WorkSheet) SetRange(ref string, values [][]interface{}) error {
	return setRangeInMap(s.fm(), s.file, ref, values, s.book.sst())
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

// GetValue 读取单元格值（按类型返回原生 Go 值，见 WorkSheet.GetCellValue）。
func (c *Cell) GetValue() (interface{}, error) { return c.sheet.GetCellValue(c.ref) }

// GetInt 读取整数值（非整数/非数值时返回 0）。
func (c *Cell) GetInt() int {
	v, err := c.sheet.GetCellValue(c.ref)
	if err != nil {
		return 0
	}
	switch n := v.(type) {
	case int:
		return n
	case float64:
		return int(n)
	case string:
		if i, e := strconv.Atoi(n); e == nil {
			return i
		}
	}
	return 0
}

// GetFloat 读取浮点值（无法解析时返回 0）。
func (c *Cell) GetFloat() float64 {
	v, err := c.sheet.GetCellValue(c.ref)
	if err != nil {
		return 0
	}
	switch n := v.(type) {
	case int:
		return float64(n)
	case float64:
		return n
	case string:
		if f, e := strconv.ParseFloat(n, 64); e == nil {
			return f
		}
	}
	return 0
}

// GetBool 读取布尔值（t="b" 时 "1" 为真）。
func (c *Cell) GetBool() bool {
	v, err := c.sheet.GetCellValue(c.ref)
	if err != nil {
		return false
	}
	if b, ok := v.(bool); ok {
		return b
	}
	return false
}

// GetTime 读取日期时间值。仅当单元格数字格式指示为日期/时间时才转换，
// 返回 1900 日期系统下的 time.Time；否则返回零值与 false。
//
// 日期判定优先依据单元格样式的 numFmtId（内置日期时间格式 14-22/27-31/45-47/50-58/71-81），
// 再回退到格式代码文本判定（剥离字面量与方括号区段后，含 y/d/h/s 或带同伴的 m 即视为日期时间），
// 比单纯子串匹配 "ymd" 更可靠，可避免 "0.0\"m\""（米）之类误判为日期，也能识别自定义中文日期格式。
func (c *Cell) GetTime() (time.Time, bool) {
	st, err := c.sheet.GetStyle(c.ref)
	isDate := false
	if err == nil {
		isDate = isDateFormat(st.NumFmtId, st.NumFmt)
	}
	// t="d" 是 ISO8601 日期时间（OOXML 的写法，流式写入用）。
	// 它已经携带完整时刻，无需也不该走"序列号 + 格式"的路径。
	if c.cellType() == "d" {
		return c.parseISODate()
	}
	v, err := c.sheet.GetCellValue(c.ref)
	if err != nil || !isDate {
		return time.Time{}, false
	}
	serial, ok := v.(float64)
	if !ok {
		if i, ok2 := v.(int); ok2 {
			serial = float64(i)
		} else {
			return time.Time{}, false
		}
	}
	// Excel 1900 日期系统：基准 1899-12-30（含 1900-02-29 伪闰日）
	//
	// 注意：必须先在**浮点秒**上算，最后一步才转 time.Duration。
	// 反过来（先 Duration 化再除）会丢精度 —— 序列号的浮点尾数在小数位
	// 上承载着秒以下的信息（如 23:59:59 对应 .99998842593，
	// 直接 `serial * 86400 * 1e9` 取整会差出几百纳秒，往返后
	// 23:59:59 变成 23:59:59.000000512。
	base := time.Date(1899, 12, 30, 0, 0, 0, 0, time.UTC)
	secs := serial * 86400
	whole := math.Floor(secs)
	nanos := (secs - whole) * 1e9
	return base.Add(time.Duration(int64(whole)) * time.Second).
		Add(time.Duration(int64(nanos))), true
}

// isDateFormat 判断某数字格式是否为日期/时间格式。
func isDateFormat(numFmtId int, code string) bool {
	// 内置日期时间格式 ID（ECMA-376）
	switch {
	case numFmtId >= 14 && numFmtId <= 22:
		return true
	case numFmtId >= 27 && numFmtId <= 31:
		return true
	case numFmtId >= 45 && numFmtId <= 47:
		return true
	case numFmtId >= 50 && numFmtId <= 58:
		return true
	case numFmtId >= 71 && numFmtId <= 81:
		return true
	}
	if code == "" {
		return false
	}
	// 自定义格式：剥离字面量（"..."）与方括号区段（[...]，含区域/颜色/条件），
	// 避免其中出现的 y/m/d/h/s 干扰判定。
	clean := stripFormatLiterals(code)
	clean = strings.ToLower(clean)
	if strings.ContainsAny(clean, "y") || strings.ContainsAny(clean, "d") {
		return true
	}
	if strings.ContainsAny(clean, "h") || strings.ContainsAny(clean, "s") {
		return true
	}
	// 仅含 m：仅当 m 与 y/d/h/s 同时出现才视为日期（否则可能是普通数字格式占位）
	if strings.Contains(clean, "m") && strings.ContainsAny(clean, "ydhs") {
		return true
	}
	return false
}

// stripFormatLiterals 去掉数字格式代码中的字面量文本（双引号括起）与方括号区段（[...]）。
func stripFormatLiterals(code string) string {
	var b strings.Builder
	inQuote := false
	inBracket := 0
	for _, r := range code {
		switch {
		case r == '"':
			inQuote = !inQuote
		case r == '[':
			inBracket++
		case r == ']':
			if inBracket > 0 {
				inBracket--
			}
		case inQuote || inBracket > 0:
			// 跳过字面量 / 方括号区段内容
		default:
			b.WriteRune(r)
		}
	}
	return b.String()
}

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

// cellType 返回单元格的 t 属性（"s"/"b"/"d"/"inlineStr"/"str"/"" 等）。
func (c *Cell) cellType() string {
	cellXML := extractCellTag(string(c.sheet.ws()), c.ref)
	return attrOf(cellXML, "t")
}

// parseISODate 解析 t="d" 单元格的 ISO8601 日期时间文本。
//
// 格式形如 2026-03-15T12:30:00Z（UTC）。解析失败返回 ok=false，
// 而不是返回一个看似合理却错误的时间 —— 日期读错比读不到更危险。
func (c *Cell) parseISODate() (time.Time, bool) {
	raw := c.cellText()
	for _, layout := range []string{
		time.RFC3339,
		"2006-01-02T15:04:05",
		"2006-01-02T15:04",
		"2006-01-02",
	} {
		if t, err := time.Parse(layout, raw); err == nil {
			return t, true
		}
	}
	return time.Time{}, false
}

// cellText 返回单元格的文本内容（<v> 或 <is><t>）。
func (c *Cell) cellText() string {
	ws := string(c.sheet.ws())
	ref := regexp.QuoteMeta(c.ref)
	// 自闭合的 <c r="X"/> 视为空
	if regexp.MustCompile(`<c\b[^>]*\br="` + ref + `"[^>]*/>`).MatchString(ws) {
		return ""
	}
	// 取整个 <c ...>...</c>：extractCellTag 只返回开标签，子元素不在其中
	m := regexp.MustCompile(`(?s)<c\b[^>]*\br="` + ref + `"[^>]*>(.*?)</c>`).
		FindStringSubmatch(ws)
	if m == nil {
		return ""
	}
	inner := m[1]
	if v := firstTagText(inner, "v"); v != "" {
		return unescapeXML(v)
	}
	if t := firstTagText(inner, "t"); t != "" {
		return unescapeXML(t)
	}
	return ""
}
