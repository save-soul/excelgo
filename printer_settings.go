package excelgo

import (
	"fmt"
	"path"
	"regexp"
	"strings"
)

// 打印机设置二进制（printerSettings）的读写。
//
// 它是什么：`xl/printerSettings/printerSettingsN.bin` 是 Excel/WPS 保存工作簿时
// 写下的**打印机专属二进制块**（DEVMODE 结构：纸张、驱动、边距微调、双面设置等）。
// 由 worksheet 的 <pageSetup r:id="rIdN"> 引用。
//
// 为什么本库要碰它：这类二进制**无法由本库生成** —— 它编码的是具体打印机驱动的
// 能力集。本库能做的只有两件事：
//  1. 搬运保真：复制/合并工作表时把 .bin 部件与关系一起搬走；
//  2. 不破坏：改页面设置时不得抹掉 pageSetup 上的 r:id。
//
// 关键细节：Content_Types 里必须声明的是 **Default Extension="bin"**，
// 不是 Override。bin 是扩展名级声明（同一 .bin 扩展可能被多种部件共用），
// 写成 Override 会被 Excel 判为非法。

// printerSettingsRelType 是 printerSettings 的关系类型。
const printerSettingsRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings"

// printerSettingsContentType 是 .bin 扩展对应的内容类型。
const printerSettingsContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings"

// SetPrinterSettings 把一段打印机设置二进制挂到工作表上。
//
// data 通常来自另一份工作簿（见 GetPrinterSettings）或 Excel/WPS 保存的文件，
// 用于把打印版式从一个工作簿搬到另一个。
//
// 注意：Excel 遇到无法识别的 DEVMODE 会退回默认打印机，不会判文件损坏，
// 所以这里只保证结构正确，不保证二进制内容适配目标机器。
func SetPrinterSettings(filename, sheetRef string, data []byte) error {
	if len(data) == 0 {
		return fmt.Errorf("打印机设置二进制为空：需要一个非空的 DEVMODE 块")
	}
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		s.setPrinterSettingsBytes(data)
		return nil
	})
}

// GetPrinterSettings 读取工作表挂着的打印机设置二进制；没有则返回 nil。
func GetPrinterSettings(filename, sheetRef string) ([]byte, error) {
	var out []byte
	err := withSheet(filename, sheetRef, func(s *WorkSheet) error {
		out = s.printerSettingsBytes()
		return nil
	})
	if err != nil {
		return nil, err
	}
	return out, nil
}

// setPrinterSettingsBytes 写入 .bin 部件、建立 sheet 级关系、并挂上 pageSetup 的 r:id。
func (s *WorkSheet) setPrinterSettingsBytes(data []byte) {
	fm := s.fm()

	// 1) 部件：复用本表已挂的那一个，没有则新建
	if old := s.printerSettingsPartName(fm); old != "" {
		fm[old] = data
		s.setPageSetupRelID(s.printerSettingsRelID(fm))
		return
	}
	num := nextFreeNumber(fm, "xl/printerSettings/printerSettings", ".bin")
	partName := fmt.Sprintf("xl/printerSettings/printerSettings%d.bin", num)
	fm[partName] = data
	fm["[Content_Types].xml"] = insertDefaultExtensionInContentTypes(
		fm["[Content_Types].xml"], "bin", printerSettingsContentType)

	// 2) sheet rels 里加关系
	relsPath := "xl/worksheets/_rels/" + path.Base(s.file) + ".rels"
	rels := loadRels(fm, relsPath)
	if rels == nil {
		rels = &Relationships{XMLName: xmlNameRelationships()}
	}
	rid := fmt.Sprintf("rId%d", getMaxRId(rels)+1)
	rels.Relationship = append(rels.Relationship, Relationship{
		ID:     rid,
		Type:   printerSettingsRelType,
		Target: "/" + partName,
	})
	fm[relsPath] = serializeRelationships(rels)

	// 3) <pageSetup r:id>
	s.setPageSetupRelID(rid)
}

// printerSettingsBytes 读回本表挂着的打印机设置二进制（没有则 nil）。
func (s *WorkSheet) printerSettingsBytes() []byte {
	fm := s.fm()
	p := s.printerSettingsPartName(fm)
	if p == "" {
		return nil
	}
	return fm[p]
}

// printerSettingsPartName 返回本表引用的打印机设置部件路径（没有则空串）。
func (s *WorkSheet) printerSettingsPartName(fm map[string][]byte) string {
	rid := s.printerSettingsRelID(fm)
	if rid == "" {
		return ""
	}
	relsPath := "xl/worksheets/_rels/" + path.Base(s.file) + ".rels"
	rels := loadRels(fm, relsPath)
	if rels == nil {
		return ""
	}
	for _, r := range rels.Relationship {
		if r.ID == rid && strings.Contains(r.Type, "/printerSettings") {
			return resolveTarget("xl/worksheets", r.Target)
		}
	}
	return ""
}

// printerSettingsRelID 从 <pageSetup r:id> 取出打印机设置的关系 ID。
func (s *WorkSheet) printerSettingsRelID(fm map[string][]byte) string {
	ws := string(s.ws())
	tag := regexp.MustCompile(`(?s)<pageSetup\b[^>]*?/?>`).FindString(ws)
	rid := attrOf(tag, "r:id")
	if rid == "" {
		return ""
	}
	relsPath := "xl/worksheets/_rels/" + path.Base(s.file) + ".rels"
	rels := loadRels(fm, relsPath)
	if rels == nil {
		return ""
	}
	for _, r := range rels.Relationship {
		if r.ID == rid && strings.Contains(r.Type, "/printerSettings") {
			return rid
		}
	}
	return ""
}

// setPageSetupRelID 在 <pageSetup> 上挂 r:id；不存在 pageSetup 时先建一个。
func (s *WorkSheet) setPageSetupRelID(rid string) {
	ws := string(s.ws())
	ws = ensureWorksheetNamespaces(ws)
	attr := ` r:id="` + safeAttr(rid) + `"`
	if re := regexp.MustCompile(`(?s)<pageSetup\b[^>]*?/?>`); re.MatchString(ws) {
		old := re.FindString(ws)
		var newTag string
		if regexp.MustCompile(`\br:id="[^"]*"`).MatchString(old) {
			newTag = regexp.MustCompile(`\br:id="[^"]*"`).ReplaceAllString(old, attr)
		} else if strings.HasSuffix(old, "/>") {
			newTag = strings.TrimSuffix(old, "/>") + attr + `/>`
		} else {
			newTag = strings.TrimSuffix(old, ">") + attr + `>`
		}
		s.replaceWSIn(ws, old, newTag)
		return
	}
	s.setWS(insertBeforeWorksheetClose(ws, `<pageSetup`+attr+`/>`))
}

// insertDefaultExtensionInContentTypes 插入 <Default Extension="..."/>（幂等）。
//
// 与 Override 的区别很要紧：bin 是**扩展名级**声明 —— 同一 .bin 扩展可能被
// 多种部件共用。若写成 Override 指向单个 PartName，Excel 会判为非法。
func insertDefaultExtensionInContentTypes(ctXML []byte, ext, contentType string) []byte {
	content := string(ctXML)
	if strings.Contains(content, fmt.Sprintf(`<Default Extension="%s"`, ext)) {
		return ctXML
	}
	tag := fmt.Sprintf(`<Default Extension="%s" ContentType="%s"/>`,
		safeAttr(ext), safeAttr(contentType))
	if i := strings.Index(content, "<Default"); i != -1 {
		return []byte(content[:i] + tag + content[i:])
	}
	i := strings.Index(content, "<Override")
	if i != -1 {
		return []byte(content[:i] + tag + content[i:])
	}
	// 没有任何 Default/Override：插在 <Types ...> 之后
	if j := strings.Index(content, ">"); j != -1 {
		return []byte(content[:j+1] + tag + content[j+1:])
	}
	return ctXML
}