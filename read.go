package excelgo

// read.go 提供与写入 API 对称的「读取」能力，覆盖合并单元格、超链接、数据验证、
// 条件格式、批注、图片等，使「读模板 → 加工 → 重写」闭环完整。
//
// 所有读取均直接解析 xl/worksheets/sheetN.xml 及其关系文件（_rels），不依赖 Office。

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// ---------- 结构体 ----------

// Hyperlink 超链接信息。
type Hyperlink struct {
	Cell     string // 所在单元格（如 "A1"）
	URL      string // 目标地址（外部链接为 URL；内部链接为相对路径）
	Display  string // 显示文本（可能为空）
	External bool   // 是否为外部链接（TargetMode="External"）
}

// DataValidation 数据验证信息。
type DataValidation struct {
	Ref        string // 作用区域（如 "A1:A10"）
	Type       string // 类型：list/whole/decimal/date/textLength/custom 等
	Operator   string // 比较运算符（可能为空）
	Formula1   string // 公式1
	Formula2   string // 公式2（可能为空）
	AllowBlank bool   // 允许空白
}

// ConditionalFormat 条件格式规则。
type ConditionalFormat struct {
	Ref      string // 作用区域（sqref）
	Type     string // 规则类型：cellIs/expression/top10/duplicateValues 等
	Operator string // 运算符（可能为空）
	Priority int    // 优先级
	Formula  string // 规则公式
}

// Comment 批注信息。
type Comment struct {
	Cell   string // 所在单元格
	Text   string // 批注文本
	Author string // 作者（可能为空）
}

// PictureInfo 工作表中的图片信息。
type PictureInfo struct {
	CellRef   string // 锚定起始单元格（如 "A1"）
	Name      string // 媒体文件名（如 "image1.png"）
	FileInZip string // 在 xlsx 包内的完整路径（如 "xl/media/image1.png"）
	WidthPx   int    // 宽度（像素）
	HeightPx  int    // 高度（像素）
}

// ---------- 内部辅助 ----------

// sheetRelsPath 由 worksheet 文件路径推导其关系文件路径。
func sheetRelsPath(sheetFile string) string {
	i := strings.LastIndex(sheetFile, "/")
	dir := sheetFile[:i]
	base := sheetFile[i+1:]
	return dir + "/_rels/" + base + ".rels"
}

// resolveRel 把相对 Target（可能含 ".."）解析为 xlsx 包内绝对路径。
// 基准目录为 rels 文件所在目录（即源部件目录，OOXML 规范：关系 Target 相对
// 于其描述部件的目录，而非 _rels 目录）。
func resolveRel(relsPath, target string) string {
	baseDir := relsPath[:strings.LastIndex(relsPath, "/")]
	if i := strings.LastIndex(baseDir, "/_rels"); i >= 0 {
		baseDir = baseDir[:i]
	}
	segs := strings.Split(baseDir, "/")
	for _, p := range strings.Split(target, "/") {
		switch p {
		case "..":
			if len(segs) > 0 {
				segs = segs[:len(segs)-1]
			}
		case ".", "":
			// 跳过
		default:
			segs = append(segs, p)
		}
	}
	return strings.Join(segs, "/")
}

// readRelsMap 读取关系文件，返回 rId → Relationship。
func readRelsMap(fileMap map[string][]byte, relsPath string) map[string]Relationship {
	m := map[string]Relationship{}
	data, ok := fileMap[relsPath]
	if !ok {
		return m
	}
	re := regexp.MustCompile(`(?s)<Relationship\b[^>]*?(?:/>|>(.*?)</Relationship>)`)
	for _, r := range re.FindAllString(string(data), -1) {
		id := attrOf(r, "Id")
		if id == "" {
			continue
		}
		m[id] = Relationship{
			ID:         id,
			Type:       attrOf(r, "Type"),
			Target:     attrOf(r, "Target"),
			TargetMode: attrOf(r, "TargetMode"),
		}
	}
	return m
}

// ---------- 合并单元格 ----------

func readMerges(ws string) []string {
	var out []string
	re := regexp.MustCompile(`(?s)<mergeCell\b[^>]*\bref="([^"]*)"[^>]*/?>`)
	for _, m := range re.FindAllStringSubmatch(ws, -1) {
		out = append(out, m[1])
	}
	return out
}

// GetMergeCells 读取 sheet 中所有合并区域（如 "A1:B2"）。
func (s *WorkSheet) GetMergeCells() ([]string, error) {
	return readMerges(string(s.ws())), nil
}

// GetMergeCells 读取 filename 中 sheetRef 的所有合并区域。
func GetMergeCells(filename, sheetRef string) ([]string, error) {
	b, err := Open(filename)
	if err != nil {
		return nil, err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return nil, err
	}
	return ws.GetMergeCells()
}

// ---------- 超链接 ----------

func readHyperlinks(fileMap map[string][]byte, sheetFile string) []Hyperlink {
	ws := string(fileMap[sheetFile])
	rels := readRelsMap(fileMap, sheetRelsPath(sheetFile))
	var out []Hyperlink
	re := regexp.MustCompile(`(?s)<hyperlink\b[^>]*?(?:/>|>(.*?)</hyperlink>)`)
	for _, h := range re.FindAllString(ws, -1) {
		cell := attrOf(h, "ref")
		rid := attrOf(h, "r:id")
		display := attrOf(h, "display")
		hl := Hyperlink{Cell: cell, Display: display}
		if rel, ok := rels[rid]; ok {
			hl.URL = rel.Target
			hl.External = rel.TargetMode == "External"
		}
		out = append(out, hl)
	}
	return out
}

// GetHyperlinks 读取 sheet 中所有超链接。
func (s *WorkSheet) GetHyperlinks() ([]Hyperlink, error) {
	return readHyperlinks(s.fm(), s.file), nil
}

// GetHyperlinks 读取 filename 中 sheetRef 的所有超链接。
func GetHyperlinks(filename, sheetRef string) ([]Hyperlink, error) {
	b, err := Open(filename)
	if err != nil {
		return nil, err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return nil, err
	}
	return ws.GetHyperlinks()
}

// ---------- 数据验证 ----------

func readDataValidations(ws string) []DataValidation {
	var out []DataValidation
	re := regexp.MustCompile(`(?s)<dataValidation\b[^>]*?(?:/>|>(.*?)</dataValidation>)`)
	for _, m := range re.FindAllStringSubmatch(ws, -1) {
		tag := m[0]
		dv := DataValidation{
			Ref:        attrOf(tag, "sqref"),
			Type:       attrOf(tag, "type"),
			Operator:   attrOf(tag, "operator"),
			AllowBlank: attrOf(tag, "allowBlank") == "1",
		}
		inner := m[1]
		if f1 := regexp.MustCompile(`(?s)<formula1>(.*?)</formula1>`).FindStringSubmatch(inner); f1 != nil {
			dv.Formula1 = unescapeXML(f1[1])
		}
		if f2 := regexp.MustCompile(`(?s)<formula2>(.*?)</formula2>`).FindStringSubmatch(inner); f2 != nil {
			dv.Formula2 = unescapeXML(f2[1])
		}
		out = append(out, dv)
	}
	return out
}

// GetDataValidations 读取 sheet 中所有数据验证规则。
func (s *WorkSheet) GetDataValidations() ([]DataValidation, error) {
	return readDataValidations(string(s.ws())), nil
}

// GetDataValidations 读取 filename 中 sheetRef 的所有数据验证规则。
func GetDataValidations(filename, sheetRef string) ([]DataValidation, error) {
	b, err := Open(filename)
	if err != nil {
		return nil, err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return nil, err
	}
	return ws.GetDataValidations()
}

// ---------- 条件格式 ----------

func readConditionalFormats(ws string) []ConditionalFormat {
	var out []ConditionalFormat
	re := regexp.MustCompile(`(?s)<conditionalFormatting\b[^>]*?(?:/>|>(.*?)</conditionalFormatting>)`)
	for _, m := range re.FindAllStringSubmatch(ws, -1) {
		tag := m[0]
		ref := attrOf(tag, "sqref")
		inner := m[1]
		cfRe := regexp.MustCompile(`(?s)<cfRule\b[^>]*?(?:/>|>(.*?)</cfRule>)`)
		for _, cm := range cfRe.FindAllStringSubmatch(inner, -1) {
			ctag := cm[0]
			cf := ConditionalFormat{
				Ref:      ref,
				Type:     attrOf(ctag, "type"),
				Operator: attrOf(ctag, "operator"),
			}
			if p := attrOf(ctag, "priority"); p != "" {
				if n, err := strconv.Atoi(p); err == nil {
					cf.Priority = n
				}
			}
			if formula := regexp.MustCompile(`(?s)<formula>(.*?)</formula>`).FindStringSubmatch(cm[1]); formula != nil {
				cf.Formula = unescapeXML(formula[1])
			}
			out = append(out, cf)
		}
	}
	return out
}

// GetConditionalFormats 读取 sheet 中所有条件格式规则。
func (s *WorkSheet) GetConditionalFormats() ([]ConditionalFormat, error) {
	return readConditionalFormats(string(s.ws())), nil
}

// GetConditionalFormats 读取 filename 中 sheetRef 的所有条件格式规则。
func GetConditionalFormats(filename, sheetRef string) ([]ConditionalFormat, error) {
	b, err := Open(filename)
	if err != nil {
		return nil, err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return nil, err
	}
	return ws.GetConditionalFormats()
}

// ---------- 批注 ----------

func readComments(fileMap map[string][]byte, sheetFile string) []Comment {
	rels := readRelsMap(fileMap, sheetRelsPath(sheetFile))
	// 找到 comments 关系
	var commentsPath string
	for _, rel := range rels {
		if strings.Contains(rel.Type, "comments") {
			commentsPath = resolveRel(sheetRelsPath(sheetFile), rel.Target)
			break
		}
	}
	if commentsPath == "" {
		return nil
	}
	data, ok := fileMap[commentsPath]
	if !ok {
		return nil
	}
	xml := string(data)

	// 作者表
	authors := []string{}
	if am := regexp.MustCompile(`(?s)<authors>(.*?)</authors>`).FindStringSubmatch(xml); am != nil {
		re := regexp.MustCompile(`(?s)<author[^>]*>(.*?)</author>`)
		for _, a := range re.FindAllStringSubmatch(am[1], -1) {
			authors = append(authors, unescapeXML(a[1]))
		}
	}

	var out []Comment
	re := regexp.MustCompile(`(?s)<comment\b[^>]*\bref="([^"]*)"[^>]*authorId="(\d+)"[^>]*>(.*?)</comment>`)
	for _, m := range re.FindAllStringSubmatch(xml, -1) {
		c := Comment{Cell: m[1]}
		if idx, err := strconv.Atoi(m[2]); err == nil && idx < len(authors) {
			c.Author = authors[idx]
		}
		textInner := m[3]
		// 批注文本可能在 <text>...<t>片段</t>... 或多段 <r><t>...</t></r>
		var sb strings.Builder
		for _, tm := range regexp.MustCompile(`(?s)<t\b[^>]*>(.*?)</t>`).FindAllStringSubmatch(textInner, -1) {
			sb.WriteString(unescapeXML(tm[1]))
		}
		c.Text = sb.String()
		out = append(out, c)
	}
	return out
}

// GetComments 读取 sheet 中所有批注。
func (s *WorkSheet) GetComments() ([]Comment, error) {
	return readComments(s.fm(), s.file), nil
}

// GetComments 读取 filename 中 sheetRef 的所有批注。
func GetComments(filename, sheetRef string) ([]Comment, error) {
	b, err := Open(filename)
	if err != nil {
		return nil, err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return nil, err
	}
	return ws.GetComments()
}

// ---------- 图片 ----------

// readPictures 读取 sheet 中所有图片。支持两种存储形态：
//  1. 单元格内嵌图片（x14:blip / xdr:blip 直接位于 worksheet XML，rId 对应该 sheet 自身 rels 的 image 关系）；
//  2. 浮动图片（位于 drawing 文件，rId 对应 drawing rels）。
func readPictures(fileMap map[string][]byte, sheetFile string) []PictureInfo {
	rels := readRelsMap(fileMap, sheetRelsPath(sheetFile))
	var out []PictureInfo

	// 形态 1：单元格内嵌图片
	ws := string(fileMap[sheetFile])
	blipRe := regexp.MustCompile(`(?s)<(?:x14|xdr):blip\b[^>]*?\br:embed="([^"]+)"[^>]*/?>`)
	for _, m := range blipRe.FindAllStringSubmatch(ws, -1) {
		rid := m[1]
		rel, ok := rels[rid]
		if !ok {
			continue
		}
		info := PictureInfo{FileInZip: resolveRel(sheetRelsPath(sheetFile), rel.Target)}
		info.Name = info.FileInZip[strings.LastIndex(info.FileInZip, "/")+1:]
		out = append(out, info)
	}
	if len(out) > 0 {
		return out
	}

	// 形态 2：drawing 浮动图片
	var drawingPath string
	for _, rel := range rels {
		if strings.Contains(rel.Type, "drawing") {
			drawingPath = resolveRel(sheetRelsPath(sheetFile), rel.Target)
			break
		}
	}
	if drawingPath == "" {
		return nil
	}
	ddata, ok := fileMap[drawingPath]
	if !ok {
		return nil
	}
	dxml := string(ddata)
	drels := readRelsMap(fileMap, sheetRelsPath(drawingPath))

	anchorRe := regexp.MustCompile(`(?s)<(?:xdr:)?(?:twoCellAnchor|oneCellAnchor)\b.*?</(?:xdr:)?(?:twoCellAnchor|oneCellAnchor)>`)
	for _, anchor := range anchorRe.FindAllString(dxml, -1) {
		pic := regexp.MustCompile(`(?s)<(?:xdr:)?pic\b.*?</(?:xdr:)?pic>`).FindString(anchor)
		if pic == "" {
			continue
		}
		embed := attrOf(regexp.MustCompile(`<(?:a:)?blip\b[^>]*?\br:embed="([^"]*)"[^>]*/?>`).FindString(pic), "r:embed")
		if embed == "" {
			continue
		}
		rel, ok := drels[embed]
		if !ok {
			continue
		}
		info := PictureInfo{
			FileInZip: resolveRel(sheetRelsPath(drawingPath), rel.Target),
		}
		info.Name = info.FileInZip[strings.LastIndex(info.FileInZip, "/")+1:]
		// 锚定单元格
		if fm := regexp.MustCompile(`<(?:xdr:)?from>(.*?)</(?:xdr:)?from>`).FindStringSubmatch(anchor); fm != nil {
			colM := regexp.MustCompile(`<(?:xdr:)?col>(\d+)</(?:xdr:)?col>`).FindStringSubmatch(fm[1])
			rowM := regexp.MustCompile(`<(?:xdr:)?row>(\d+)</(?:xdr:)?row>`).FindStringSubmatch(fm[1])
			if colM != nil && rowM != nil {
				col, _ := strconv.Atoi(colM[1])
				row, _ := strconv.Atoi(rowM[1])
				info.CellRef = colNumToLetters(col) + strconv.Itoa(row+1)
			}
		}
		// 尺寸（EMU → 像素，9525 EMU/px）
		if ext := regexp.MustCompile(`<(?:xdr:)?ext\b[^>]*cx="(\d+)"[^>]*cy="(\d+)"`).FindStringSubmatch(anchor); ext != nil {
			if cx, err := strconv.Atoi(ext[1]); err == nil {
				info.WidthPx = cx / 9525
			}
			if cy, err := strconv.Atoi(ext[2]); err == nil {
				info.HeightPx = cy / 9525
			}
		}
		out = append(out, info)
	}
	return out
}

// GetPictures 读取 sheet 中所有图片信息。
func (s *WorkSheet) GetPictures() ([]PictureInfo, error) {
	return readPictures(s.fm(), s.file), nil
}

// GetPictures 读取 filename 中 sheetRef 的所有图片信息。
func GetPictures(filename, sheetRef string) ([]PictureInfo, error) {
	b, err := Open(filename)
	if err != nil {
		return nil, err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return nil, err
	}
	return ws.GetPictures()
}

// ExtractPicture 导出 sheet 中指定图片的字节数据。fileInZip 取自查得的 PictureInfo.FileInZip
// （如 "xl/media/image1.png"），与 Open/Excel 内存储路径一致。
func (s *WorkSheet) ExtractPicture(fileInZip string) ([]byte, error) {
	data, ok := s.fm()[fileInZip]
	if !ok {
		return nil, fmt.Errorf("未找到图片部件 %q", fileInZip)
	}
	return data, nil
}

// ExtractPicture 读取 filename 中 sheetRef 指定图片的字节数据（fileInZip 来自 GetPictures 的结果）。
func ExtractPicture(filename, sheetRef, fileInZip string) ([]byte, error) {
	b, err := Open(filename)
	if err != nil {
		return nil, err
	}
	ws, err := b.Sheet(sheetRef)
	if err != nil {
		return nil, err
	}
	return ws.ExtractPicture(fileInZip)
}
