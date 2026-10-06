package excelgo

import (
	"archive/zip"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"time"
)

// 流式（write-only）写入。
//
// 为什么需要：常规写入把每张工作表的 XML 全程放在内存的 map[string][]byte 里，
// 且**每次写单元格都重扫整份 sheet XML** —— 复杂度 O(单元格数 × 文档长度)，
// 实测 3000 行 × 4 列耗时 4 分多钟。写大表时这不可接受。
//
// 流式的做法与 openpyxl 的 write_only 一致：
//
//	w := excelgo.NewStreamWriter("out.xlsx")
//	ws, _ := w.NewSheet("数据")
//	ws.Append([]interface{}{"名称", "数量"})
//	ws.Append([]interface{}{"甲", 10})
//	w.Save()
//
// 约束（与 openpyxl 相同，源自格式本身）：
//   - **必须先声明表名再写行** —— sheet 名与 sheetN.xml 的对应关系在
//     workbook.xml 里定死，边写边改名会让已写的数据无处安放。
//   - **行必须按行号递增写入** —— <row> 是顺序流，随机回填需要重扫全文档，
//     正是流式要消除的操作。
//   - **Close/Save 之前文件不完整**。
//
// 与常规写入的差别：不共享字符串表 —— 每个字符串单元格直接写
// t="inlineStr"（内联）。这样无需在写行时回查 SST 索引，字符串去重可延后，
// 也省掉一次全表扫。代价是文件略大，对写大表是划算的取舍。

// StreamWriter 是流式写入器。
type StreamWriter struct {
	w        *zip.Writer
	f        *os.File
	tmpDir   string // 暂存各表 sheetData 的临时目录（Save/Abort 时清理）
	sheets   []*streamSheet
	sheetNum int

	// 样式表：复用常规路径的 applyStyleToStylesXML —— 它已处理枚举校验、
	// 组件去重与 numFmt 自增。流式只需从一个最小基底开始逐个累积。
	stylesXML []byte

	cols   []string // <col> 片段（SetColWidth 累积）
	merges []string // 合并区域
	freeze string   // 冻结窗格 XML
	rows   int      // 已写行数（仅统计，不做限制）
}

// streamSheet 是流式模式下的一张工作表。
//
// 数据行写入**临时文件**而非内存 —— 这是 zip 格式的硬约束决定的：
// archive/zip 的 Writer 一次只能有一个打开的条目，Create 新条目会关闭前一个。
// 所以多表不能"交错写"，只能先各自落盘、最后按序输出。
// 用临时文件而非内存，正好也维持了"内存不随数据量增长"的目标。
type streamSheet struct {
	w        *StreamWriter
	name     string
	file     partPath // xl/worksheets/sheetN.xml（库内生成，非用户输入）
	tmp      *os.File // 暂存 sheetData 之前的 head + 各行
	rows     int
	maxCol   int
	finished bool
}

// NewStreamWriter 创建一个流式写入器（同时创建目标文件）。
func NewStreamWriter(filename string) (*StreamWriter, error) {
	f, err := os.Create(filename)
	if err != nil {
		return nil, fmt.Errorf("无法创建文件 %q: %w", filename, err)
	}
	tmpDir, err := os.MkdirTemp("", "excelgo-stream-*")
	if err != nil {
		f.Close()
		os.Remove(filename)
		return nil, fmt.Errorf("无法创建临时目录: %w", err)
	}
	return &StreamWriter{
		w:         zip.NewWriter(f),
		f:         f,
		tmpDir:    tmpDir,
		stylesXML: []byte(defaultStylesXML()),
	}, nil
}

// NewSheet 声明一张新工作表并返回流式写入句柄。
//
// 表名必须在写数据前确定 —— workbook.xml 里 sheet 名与 sheetN.xml 的对应关系
// 是固定的，写完行再改名会让已写的数据无处安放。
func (sw *StreamWriter) NewSheet(name string) (*StreamSheet, error) {
	if name == "" {
		return nil, fmt.Errorf("工作表名不能为空")
	}
	for _, s := range sw.sheets {
		if s.name == name {
			return nil, fmt.Errorf("工作表 %q 已存在", name)
		}
	}
	sw.sheetNum++
	file := partPath(fmt.Sprintf("xl/worksheets/sheet%d.xml", sw.sheetNum))
	ss := &streamSheet{w: sw, name: name, file: file}

	// 暂存到临时文件，落盘时再作为 zip 条目输出
	tmp, err := os.Create(filepath.Join(sw.tmpDir,
		fmt.Sprintf("sheet%d.xml", sw.sheetNum)))
	if err != nil {
		return nil, fmt.Errorf("无法创建暂存文件: %w", err)
	}
	ss.tmp = tmp
	if err := ss.writeHead(); err != nil {
		tmp.Close()
		return nil, err
	}
	sw.sheets = append(sw.sheets, ss)
	return &StreamSheet{s: ss}, nil
}

// SheetNames 返回已声明的工作表名（按声明顺序）。
func (sw *StreamWriter) SheetNames() []string {
	out := make([]string, 0, len(sw.sheets))
	for _, s := range sw.sheets {
		out = append(out, s.name)
	}
	return out
}

// AddStyle 预登记一个样式，返回其索引，供 StreamCell.Style 引用。
//
// 流式模式下不能"写到某格时才建样式" —— 那时该格可能已经写出去了。
// 因此样式必须在写数据之前全部登记。
//
// 实现复用常规路径的 applyStyleToStylesXML（已含枚举校验、组件去重、
// numFmt 自增），只是从最小基底开始累积。
func (sw *StreamWriter) AddStyle(st Style) (int, error) {
	next, idx, err := applyStyleToStylesXML(sw.stylesXML, st)
	if err != nil {
		return 0, err
	}
	sw.stylesXML = next
	return idx, nil
}

// SetColWidth 声明列宽（写数据前调用）。
//
// 列宽写在 <cols> 里，必须位于 <sheetData> 之前 —— 这正是流式的硬约束之一。
// 流式模式下不能像常规写入那样"先写数据再改列宽"。
func (sw *StreamWriter) SetColWidth(startCol, endCol int, width float64) error {
	if startCol < 1 || endCol < startCol {
		return fmt.Errorf("无效的列范围 %d:%d", startCol, endCol)
	}
	sw.cols = append(sw.cols, fmt.Sprintf(
		`<col min="%d" max="%d" width="%s" customWidth="1"/>`,
		startCol, endCol, strconv.FormatFloat(width, 'f', -1, 64)))
	return nil
}

// MergeCell 声明合并区域（写数据前调用）。
func (sw *StreamWriter) MergeCell(ref string) error {
	if _, _, ok := refCoords(ref); !ok {
		return fmt.Errorf("无效的合并区域 %q", ref)
	}
	sw.merges = append(sw.merges, `<mergeCell ref="`+safeAttr(ref)+`"/>`)
	return nil
}

// SetFreezePanes 冻结窗格，如 "B2" 表示冻结首行首列。
func (sw *StreamWriter) SetFreezePanes(ref string) error {
	col, row, ok := refCoords(ref)
	if !ok {
		return fmt.Errorf("无效的冻结位置 %q", ref)
	}
	sw.freeze = fmt.Sprintf(
		`<pane xSplit="%d" ySplit="%d" topLeftCell="%s" activePane="bottomRight" state="frozen"/>`+
			`<selection pane="bottomRight"/>`, col, row, safeAttr(ref))
	return nil
}

// Save 收尾并写盘。调用后写入器不可再用。
//
// 顺序：先收尾各表并把元部件写进临时区，最后才按序输出 zip 条目 ——
// archive/zip 一次只允许一个打开的条目，条目必须串行、不可交错。
func (sw *StreamWriter) Save() error {
	// 1) 收尾各表（把完整 sheet XML 落到临时文件）
	for _, s := range sw.sheets {
		if err := s.writeTail(); err != nil {
			sw.cleanup()
			return fmt.Errorf("收尾工作表 %q 失败: %w", s.name, err)
		}
	}
	// 2) 元部件先写进 zip（它们都在各表条目之前，内容与数据量无关）
	if err := sw.writeMetadata(); err != nil {
		sw.cleanup()
		return err
	}
	// 3) 各表条目按序输出
	for _, s := range sw.sheets {
		if err := sw.writeSheetEntry(s); err != nil {
			sw.cleanup()
			return err
		}
	}
	if err := sw.w.Close(); err != nil {
		sw.f.Close()
		sw.removeTmp()
		return err
	}
	err := sw.f.Close()
	sw.removeTmp()
	return err
}

// writeSheetEntry 把暂存的 sheet XML 输出为 zip 条目。
func (sw *StreamWriter) writeSheetEntry(s *streamSheet) error {
	fh, err := sw.w.Create(string(s.file))
	if err != nil {
		return err
	}
	tmp, err := os.Open(filepath.Join(sw.tmpDir,
		filepath.Base(string(s.file))))
	if err != nil {
		return err
	}
	defer tmp.Close()
	_, err = io.Copy(fh, tmp)
	return err
}

// Abort 中止写入并删除半成品文件。
//
// 流式写入途中出错时文件必然不完整；直接返回错误会让用户留下一个
// 看似存在、实则打不开的文件。Abort 明确表示"这个文件不算数"。
func (sw *StreamWriter) Abort() {
	sw.cleanup()
}

// cleanup 关闭并删除半成品（忽略错误，尽力而为）。
func (sw *StreamWriter) cleanup() {
	sw.w.Close()
	sw.f.Close()
	os.Remove(sw.f.Name())
	sw.removeTmp()
}

// removeTmp 清理暂存目录（尽力而为，忽略错误）。
func (sw *StreamWriter) removeTmp() {
	if sw.tmpDir == "" {
		return
	}
	os.RemoveAll(sw.tmpDir)
	sw.tmpDir = ""
}

// writeHead 写 sheet XML 的开头部分（含 cols/冻结窗格，必须在 sheetData 之前）。
func (ss *streamSheet) writeHead() error {
	var b strings.Builder
	b.WriteString(xmlDecl)
	b.WriteString(`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" `)
	b.WriteString(`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`)

	if len(ss.w.cols) > 0 {
		b.WriteString(`<cols>`)
		for _, c := range ss.w.cols {
			b.WriteString(c)
		}
		b.WriteString(`</cols>`)
	}
	if ss.w.freeze != "" {
		b.WriteString(`<sheetViews><sheetView workbookViewId="0" tabSelected="0">`)
		// freeze 在 SetFreezePanes 时已格式化为完整片段，直接拼接
		b.WriteString(ss.w.freeze)
		b.WriteString(`</sheetView></sheetViews>`)
	}
	b.WriteString(`<sheetData>`)
	_, err := ss.tmp.Write([]byte(b.String()))
	return err
}

// writeTail 收尾 sheet XML（关闭 sheetData 并补后续元素）。
func (ss *streamSheet) writeTail() error {
	if ss.finished {
		return nil
	}
	ss.finished = true
	var b strings.Builder
	b.WriteString(`</sheetData>`)
	if len(ss.w.merges) > 0 {
		fmt.Fprintf(&b, `<mergeCells count="%d">`, len(ss.w.merges))
		for _, m := range ss.w.merges {
			b.WriteString(m)
		}
		b.WriteString(`</mergeCells>`)
	}
	// OOXML 要求 dimension 在 sheetData 之前；这里省略它（可选元素），
	// 不少消费方（含 openpyxl）都能容忍。
	b.WriteString(`</worksheet>`)
	if _, err := ss.tmp.Write([]byte(b.String())); err != nil {
		return err
	}
	return ss.tmp.Close()
}

// writeMetadata 写 workbook / rels / styles / Content_Types。
func (sw *StreamWriter) writeMetadata() error {
	// --- [Content_Types].xml ---
	var ct strings.Builder
	ct.WriteString(xmlDecl)
	ct.WriteString(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`)
	ct.WriteString(`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`)
	ct.WriteString(`<Default Extension="xml" ContentType="application/xml"/>`)
	ct.WriteString(`<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`)
	ct.WriteString(`<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`)
	for _, s := range sw.sheets {
		ct.WriteString(`<Override PartName="/` + string(s.file) + `" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`)
	}
	ct.WriteString(`</Types>`)

	// --- _rels/.rels ---
	rootRels := xmlDecl +
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
		`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
		`</Relationships>`

	// --- xl/workbook.xml ---
	var wb strings.Builder
	wb.WriteString(xmlDecl)
	wb.WriteString(`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
		`xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>`)
	for i, s := range sw.sheets {
		fmt.Fprintf(&wb, `<sheet name="%s" sheetId="%d" r:id="rId%d"/>`,
			safeAttr(s.name), i+1, i+1)
	}
	wb.WriteString(`</sheets></workbook>`)

	// --- xl/_rels/workbook.xml.rels ---
	var wbr strings.Builder
	wbr.WriteString(xmlDecl)
	wbr.WriteString(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`)
	for i := range sw.sheets {
		fmt.Fprintf(&wbr, `<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>`,
			i+1, i+1)
	}
	fmt.Fprintf(&wbr, `<Relationship Id="rId%d" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
		len(sw.sheets)+1)
	wbr.WriteString(`</Relationships>`)

	// --- xl/styles.xml ---
	styles := string(sw.stylesXML)

	parts := [][2]string{
		{"[Content_Types].xml", ct.String()},
		{"_rels/.rels", rootRels},
		{"xl/workbook.xml", wb.String()},
		{"xl/_rels/workbook.xml.rels", wbr.String()},
		{"xl/styles.xml", styles},
	}
	for _, p := range parts {
		fh, err := sw.w.Create(p[0])
		if err != nil {
			return err
		}
		if _, err := fh.Write([]byte(p[1])); err != nil {
			return err
		}
	}
	return nil
}

// StreamSheet 是流式模式下工作表的写入句柄。
type StreamSheet struct {
	s *streamSheet
}

// Name 返回工作表名。
func (sh *StreamSheet) Name() string { return sh.s.name }

// RowCount 返回已写入的行数。
func (sh *StreamSheet) RowCount() int { return sh.s.rows }

// Append 追加一行。
//
// values 的元素支持：nil、bool、int 族、float64、string、time.Time、StreamCell。
// 超出当前最大列的部分会被自动计入列数。
func (sh *StreamSheet) Append(values []interface{}) error {
	if sh.s.finished {
		return fmt.Errorf("工作表 %q 已收尾，不能继续写入", sh.s.name)
	}
	rowNum := sh.s.rows + 1
	var b strings.Builder
	fmt.Fprintf(&b, `<row r="%d">`, rowNum)
	col := 0
	for _, v := range values {
		col++
		c, err := streamCellXML(v, refOf(rowNum-1, col-1))
		if err != nil {
			return fmt.Errorf("写入第 %d 行第 %d 列失败: %w", rowNum, col, err)
		}
		if c == "" {
			continue // nil：跳过该格（与 openpyxl 的 None 一致）
		}
		b.WriteString(c)
	}
	b.WriteString(`</row>`)
	if _, err := sh.s.tmp.Write([]byte(b.String())); err != nil {
		return err
	}
	sh.s.rows = rowNum
	if col > sh.s.maxCol {
		sh.s.maxCol = col
	}
	return nil
}

// AppendRow 追加一行并返回错误（openpyxl 风格别名）。
func (sh *StreamSheet) AppendRow(values []interface{}) error { return sh.Append(values) }

// StreamCell 是带样式的流式单元格。
type StreamCell struct {
	Value interface{}
	Style int // 由 StreamWriter.AddStyle 返回的索引；0=默认
}

// TimeCell 构造带样式的日期时间单元格（t=d）。
func TimeCell(v time.Time, style int) StreamCell { return StreamCell{Value: v, Style: style} }

// streamCellXML 把一个值转成 <c> 片段；返回空串表示该格不写。
func streamCellXML(v interface{}, ref string) (string, error) {
	styleIdx := 0
	if sc, ok := v.(StreamCell); ok {
		styleIdx = sc.Style
		v = sc.Value
	}
	sAttr := ""
	if styleIdx > 0 {
		sAttr = ` s="` + string(styleIndex(strconv.Itoa(styleIdx))) + `"`
	}
	// open 只由认证类型拼成：cellRef（经 newCellRef 校验）+ styleIndex（库内生成）
	open := cellOpen(`<c r="` + string(ref) + `"` + string(sAttr))
	// 拼 t 属性时要去掉尾部空格
	openNum := cellOpen(strings.TrimRight(string(open), " "))
	switch val := v.(type) {
	case nil:
		return "", nil
	case bool:
		// 布尔落成 t="b"，值只能是 0/1 两个字面量
		b := "0"
		if val {
			b = "1"
		}
		return numCell(openNum+` t="b"`, numLit(b)), nil
	case int:
		return numCell(openNum, numLit(strconv.FormatInt(int64(val), 10))), nil
	case int8:
		return numCell(openNum, numLit(strconv.FormatInt(int64(val), 10))), nil
	case int16:
		return numCell(openNum, numLit(strconv.FormatInt(int64(val), 10))), nil
	case int32:
		return numCell(openNum, numLit(strconv.FormatInt(int64(val), 10))), nil
	case int64:
		return numCell(openNum, numLit(strconv.FormatInt(val, 10))), nil
	case uint:
		return numCell(openNum, numLit(strconv.FormatUint(uint64(val), 10))), nil
	case uint8:
		return numCell(openNum, numLit(strconv.FormatUint(uint64(val), 10))), nil
	case uint16:
		return numCell(openNum, numLit(strconv.FormatUint(uint64(val), 10))), nil
	case uint32:
		return numCell(openNum, numLit(strconv.FormatUint(uint64(val), 10))), nil
	case uint64:
		return numCell(openNum, numLit(strconv.FormatUint(val, 10))), nil
	case float32:
		return numCell(openNum, numLit(strconv.FormatFloat(float64(val), 'f', -1, 32))), nil
	case float64:
		return numCell(openNum, numLit(strconv.FormatFloat(val, 'f', -1, 64))), nil
	case time.Time:
		// ISO8601 UTC（OOXML 的 t="d" 约定）
		return numCell(openNum+` t="d"`,
			numLit(val.UTC().Format("2006-01-02T15:04:05Z"))), nil
	case string:
		// 内联字符串：流式模式不做共享去重（那需要每次回查整张表）。
		// safeText 是转义函数，走它即完成清洗。
		return string(openNum) + ` t="inlineStr"><is><t xml:space="preserve">` +
			string(safeText(val)) + `</t></is></c>`, nil
	default:
		return "", fmt.Errorf("不支持的类型 %T（可用：nil/bool/整数/浮点/string/time.Time/StreamCell）", v)
	}
}

// ---------- 数值字面量：让"这是纯数字"成为代码事实 ----------

// numLit 表示一段**纯数值字面量**文本。
//
// 为什么要这个类型：数值是从 interface{} 里断言出来的，类型系统不知道
// strconv.FormatX 的结果只可能含数字与符号。污点护栏会（正确地）把
// "未经清洗的字符串拼进 XML"当成风险点 —— 与其加豁免说明、不如让认证类型
// 把这件事写进代码里。
type numLit string

// cellOpen 是 <c> 标签的开头部分（属性部分），如 `<c r="A1" s="2"`。
//
// 认证依据：只可能由 cellRef（经 newCellRef 校验的坐标）与 styleIndex
//（库内生成的样式索引）拼成，不含用户输入。
type cellOpen string

// numCell 拼出数值单元格：<c ...><v>数字</v></c>。
func numCell(open cellOpen, n numLit) string {
	return string(open) + `><v>` + string(n) + `</v></c>`
}

// partPath 表示一个由库内生成的 xlsx 部件路径（如 xl/worksheets/sheet1.xml）。
//
// 认证依据：完全由 sheetNum 格式化而来，不含任何用户输入；字符集被限制在
// [a-zA-Z0-9/_.-]，即便调用方误传也不会引入 XML 特殊字符。
type partPath string
