package excelgo

// style_read.go 提供样式「读取/克隆」能力，与 style.go 的写入能力对称。
//
// 通过解析 xl/styles.xml：从单元格的 s 索引 → cellXfs[idx] → fontId/fillId/borderId/
// numFmtId 与 alignment，再反查 fonts/fills/borders/numFmts 各块，重建出 Style 结构。
// 这样即可"读取 A1 的样式并应用到 B2"（格式复制），无需手工逐项设置。

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// orDefault 空字符串时返回默认值。
func orDefault(s, d string) string {
	if s == "" {
		return d
	}
	return s
}

// builtinNumFmt 返回内置 numFmtId 对应的格式代码（仅常用子集）。
// 不在表中的内置格式返回 ""（调用方视为常规）。
func builtinNumFmt(id int) string {
	switch id {
	case 0:
		return "General"
	case 1:
		return "0"
	case 2:
		return "0.00"
	case 3:
		return "#,##0"
	case 4:
		return "#,##0.00"
	case 9:
		return "0%"
	case 10:
		return "0.00%"
	case 11:
		return "0.00E+00"
	case 12:
		return "# ?/?"
	case 14:
		return "mm-dd-yy"
	case 15:
		return "d-mmm-yy"
	case 16:
		return "d-mmm"
	case 17:
		return "mmm-yy"
	case 18:
		return "h:mm AM/PM"
	case 19:
		return "h:mm:ss AM/PM"
	case 20:
		return "h:mm"
	case 21:
		return "h:mm:ss"
	case 22:
		return "m/d/yy h:mm"
	case 37:
		return "#,##0 ;(#,##0)"
	case 38:
		return "#,##0 ;[Red](#,##0)"
	case 39:
		return "#,##0.00;(#,##0.00)"
	case 40:
		return "#,##0.00;[Red](#,##0.00)"
	case 45:
		return "mm:ss"
	case 46:
		return "[h]:mm:ss"
	case 47:
		return "mmss.0"
	case 48:
		return "##0.0E+0"
	case 49:
		return "@"
	}
	return ""
}

// parseStyleFromIndex 从 styles.xml 字节反构第 idx 条 cellXf 对应的 Style。
func parseStyleFromIndex(stylesXML []byte, idx int) (Style, error) {
	st := string(stylesXML)
	cellXfs := extractItems(extractBlock(st, "cellXfs"), "xf")
	if idx < 0 || idx >= len(cellXfs) {
		return Style{}, fmt.Errorf("样式索引 %d 越界（cellXfs 共 %d 条）", idx, len(cellXfs))
	}
	xf := cellXfs[idx]
	var s Style

	fontId, _ := strconv.Atoi(orDefault(attrOf(xf, "fontId"), "0"))
	fillId, _ := strconv.Atoi(orDefault(attrOf(xf, "fillId"), "0"))
	borderId, _ := strconv.Atoi(orDefault(attrOf(xf, "borderId"), "0"))
	numFmtId, _ := strconv.Atoi(orDefault(attrOf(xf, "numFmtId"), "0"))

	if fonts := extractItems(extractBlock(st, "fonts"), "font"); fontId < len(fonts) {
		if f := parseFontStyle(fonts[fontId]); f != nil {
			s.Font = f
		}
	}
	if fills := extractItems(extractBlock(st, "fills"), "fill"); fillId < len(fills) {
		if fl := parseFillStyle(fills[fillId]); fl != nil {
			s.Fill = fl
		}
	}
	if borders := extractItems(extractBlock(st, "borders"), "border"); borderId < len(borders) {
		if b := parseBorderStyle(borders[borderId]); b != nil {
			s.Border = b
		}
	}
	if numFmtId != 0 {
		s.NumFmtId = numFmtId
		if code := builtinNumFmt(numFmtId); code != "" {
			s.NumFmt = code
		} else if numFmts := extractItems(extractBlock(st, "numFmts"), "numFmt"); len(numFmts) > 0 {
			for _, nf := range numFmts {
				if attrOf(nf, "numFmtId") == strconv.Itoa(numFmtId) {
					s.NumFmt = attrOf(nf, "formatCode")
					break
				}
			}
		}
	}
	if strings.Contains(xf, "<alignment") {
		if a := parseAlignment(xf); a != nil {
			s.Alignment = a
		}
	}
	return s, nil
}

// parseFontStyle 从单个 <font> 元素反构 FontStyle；无任何字体属性时返回 nil。
func parseFontStyle(font string) *FontStyle {
	f := &FontStyle{}
	set := false
	if strings.Contains(font, "<b/>") || strings.Contains(font, "<b ") {
		f.Bold = true
		set = true
	}
	if strings.Contains(font, "<i/>") || strings.Contains(font, "<i ") {
		f.Italic = true
		set = true
	}
	if m := regexp.MustCompile(`(?s)<u\b[^>]*?(?:/>|>(.*?)</u>)`).FindStringSubmatch(font); m != nil {
		val := attrOf(m[0], "val")
		if val == "" {
			val = "single"
		}
		f.Underline = val
		set = true
	}
	if strings.Contains(font, "<strike") {
		f.Strike = true
		set = true
	}
	if m := regexp.MustCompile(`<sz\b[^>]*val="([^"]*)"`).FindStringSubmatch(font); m != nil {
		if v, err := strconv.ParseFloat(m[1], 64); err == nil {
			f.Size = v
			set = true
		}
	}
	if m := regexp.MustCompile(`<color\b[^>]*rgb="([^"]*)"`).FindStringSubmatch(font); m != nil {
		f.Color = m[1]
		set = true
	}
	if m := regexp.MustCompile(`<name\b[^>]*val="([^"]*)"`).FindStringSubmatch(font); m != nil {
		f.Name = m[1]
		set = true
	}
	if !set {
		return nil
	}
	return f
}

// parseFillStyle 从单个 <fill> 元素反构 FillStyle；none/gray125 等无填充返回 nil。
func parseFillStyle(fill string) *FillStyle {
	if !strings.Contains(fill, "patternFill") {
		return nil
	}
	pt := attrOf(fill, "patternType")
	if pt == "none" || pt == "gray125" {
		return nil
	}
	color := ""
	if m := regexp.MustCompile(`<fgColor\b[^>]*rgb="([^"]*)"`).FindStringSubmatch(fill); m != nil {
		color = m[1]
	}
	if color == "" && pt == "" {
		return nil
	}
	return &FillStyle{Color: color, PatternType: orDefault(pt, "solid")}
}

// parseBorderStyle 从单个 <border> 元素反构 BorderStyle；四边均无样式返回 nil。
func parseBorderStyle(border string) *BorderStyle {
	b := &BorderStyle{}
	set := false
	for _, name := range []xmlTagName{"left", "right", "top", "bottom", "diagonal"} {
		side := sideFromXML(border, name)
		if side.Style == "" {
			continue
		}
		set = true
		switch string(name) {
		case "left":
			b.Left = side
		case "right":
			b.Right = side
		case "top":
			b.Top = side
		case "bottom":
			b.Bottom = side
		case "diagonal":
			b.Diagonal = side
		}
	}
	if !set {
		return nil
	}
	return b
}

// sideFromXML 提取某一边框元素（含子 <color>）的样式与颜色。
func sideFromXML(border string, name xmlTagName) BorderSide {
	n := regexp.QuoteMeta(string(name))
	re := regexp.MustCompile(`(?s)<` + n + `\b[^>]*?(?:/>|>(.*?)</` + n + `>)`)
	m := re.FindStringSubmatch(border)
	if m == nil {
		return BorderSide{}
	}
	full := m[0]
	st := attrOf(full, "style")
	if st == "" {
		return BorderSide{}
	}
	color := ""
	if cm := regexp.MustCompile(`<color\b[^>]*rgb="([^"]*)"`).FindStringSubmatch(full); cm != nil {
		color = cm[1]
	}
	return BorderSide{Style: st, Color: color}
}

// parseAlignment 从 cellXf 片段反构 AlignmentStyle；无对齐属性返回 nil。
func parseAlignment(xf string) *AlignmentStyle {
	m := regexp.MustCompile(`(?s)<alignment\b[^>]*/?>`).FindStringSubmatch(xf)
	if m == nil {
		return nil
	}
	tag := m[0]
	a := &AlignmentStyle{}
	set := false
	if h := attrOf(tag, "horizontal"); h != "" {
		a.Horizontal = h
		set = true
	}
	if v := attrOf(tag, "vertical"); v != "" {
		a.Vertical = v
		set = true
	}
	if attrOf(tag, "wrapText") == "1" {
		a.WrapText = true
		set = true
	}
	if !set {
		return nil
	}
	return a
}

// ---------- 面向对象 API ----------

// GetStyle 读取 ref 单元格的完整样式（从 styles.xml 反构）。
func (s *WorkSheet) GetStyle(ref string) (Style, error) {
	idx := s.GetCellStyleIndex(ref)
	styles, ok := s.fm()["xl/styles.xml"]
	if !ok {
		return Style{}, fmt.Errorf("工作簿不含 styles.xml")
	}
	return parseStyleFromIndex(styles, idx)
}

// CopyStyle 把 from 单元格的样式克隆到 to 单元格（仅改 s 索引，保留数据与公式）。
func (s *WorkSheet) CopyStyle(from, to string) error {
	st, err := s.GetStyle(from)
	if err != nil {
		return err
	}
	_, err = s.SetStyle(to, st)
	return err
}

// GetStyleStruct 读取单元格完整样式（同 WorkSheet.GetStyle）。
func (c *Cell) GetStyleStruct() (Style, error) { return c.sheet.GetStyle(c.ref) }

// ---------- 全局函数（按文件读写） ----------

// GetCellStyle 读取 filename 中 sheetRef!cell 的样式（返回 Style 结构）。
func GetCellStyle(filename, sheetRef, cell string) (Style, error) {
	b, err := Open(filename)
	if err != nil {
		return Style{}, err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return Style{}, err
	}
	return ws.GetStyle(cell)
}

// CopyStyle 把 filename 中 sheetRef!from 的样式克隆到同表 to 单元格，并写回磁盘。
func CopyStyle(filename, sheetRef, from, to string) error {
	b, err := Open(filename)
	if err != nil {
		return err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return err
	}
	if err := ws.CopyStyle(from, to); err != nil {
		return err
	}
	return b.Save()
}
