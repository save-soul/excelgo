package excelgo

// workbook_protection.go 实现工作簿级保护（结构锁定 workbookProtection）
// 与命名区域（definedName，含工作簿级全局命名区域）的读写 API。
//
// 说明：
//   - 工作簿结构保护写入 xl/workbook.xml 的 <workbookProtection>（lockStructure/lockWindows/
//     lockRevision + 可选 workbookPassword 哈希）。位置遵循 OOXML 顺序：在 </workbookPr>
//     之后；若无 workbookPr 则置于 <bookViews> 或 <sheets> 之前。
//   - 命名区域写入 xl/workbook.xml 的 <definedNames>；工作簿级全局命名区域不带 localSheetId，
//     表级命名区域（如打印区域）由 print.go 的 setOrReplaceDefinedName 处理（带 localSheetId）。
//   - 密码采用 Excel 标准 password verifier 哈希（16 位），以十六进制字符串写入 workbookPassword 属性。

import (
	"fmt"
	"regexp"
	"strings"
)

// ---------- 工作簿保护 ----------

// WorkbookProtection 描述工作簿保护状态。
type WorkbookProtection struct {
	LockStructure bool   // 锁定工作簿结构（禁止增删/重排工作表）
	LockWindows   bool   // 锁定窗口（禁止调整窗口大小/位置）
	LockRevision  bool   // 锁定修订（共享工作簿）
	PasswordHash  string // 密码验证器哈希（4 位十六进制，空表示无密码）
}

// ProtectWorkbook 启用工作簿结构保护（可带可选密码）。
// password 为空表示不设密码；否则按 Excel 算法写入 password verifier 哈希。
func (b *Book) ProtectWorkbook(password string) error {
	return b.setWorkbookProtection(true, false, false, password)
}

// UnprotectWorkbook 解除工作簿保护（移除 <workbookProtection>）。
func (b *Book) UnprotectWorkbook() error {
	xml := string(b.fileMap["xl/workbook.xml"])
	xml = regexp.MustCompile(`(?s)<workbookProtection\b.*?/>`).ReplaceAllString(xml, "")
	xml = strings.ReplaceAll(xml, "<workbookProtection/>", "") // 兜底（理论上已被上一行覆盖）
	b.fileMap["xl/workbook.xml"] = []byte(xml)
	return nil
}

// GetWorkbookProtection 读取工作簿保护状态（未保护时返回全零结构体）。
func (b *Book) GetWorkbookProtection() (WorkbookProtection, error) {
	xml := string(b.fileMap["xl/workbook.xml"])
	wp := WorkbookProtection{}
	m := regexp.MustCompile(`(?s)<workbookProtection\b[^>]*/?>`).FindStringSubmatch(xml)
	if m == nil {
		return wp, nil
	}
	tag := m[0]
	wp.LockStructure = attrOf(tag, "lockStructure") == "1"
	wp.LockWindows = attrOf(tag, "lockWindows") == "1"
	wp.LockRevision = attrOf(tag, "lockRevision") == "1"
	wp.PasswordHash = attrOf(tag, "workbookPassword")
	return wp, nil
}

// setWorkbookProtection 内部实现：依据参数写/重写 <workbookProtection>。
func (b *Book) setWorkbookProtection(lockStruct, lockWin, lockRev bool, password string) error {
	xml := string(b.fileMap["xl/workbook.xml"])
	// 先移除已有 <workbookProtection>（无论是否自闭合）
	xml = regexp.MustCompile(`(?s)<workbookProtection\b.*?/>`).ReplaceAllString(xml, "")

	var attrs []string
	if lockStruct {
		attrs = append(attrs, `lockStructure="1"`)
	}
	if lockWin {
		attrs = append(attrs, `lockWindows="1"`)
	}
	if lockRev {
		attrs = append(attrs, `lockRevision="1"`)
	}
	if password != "" {
		attrs = append(attrs, fmt.Sprintf(`workbookPassword="%s"`, hashExcelPassword(password)))
	}
	if len(attrs) == 0 {
		b.fileMap["xl/workbook.xml"] = []byte(xml)
		return nil
	}
	tag := `<workbookProtection ` + strings.Join(attrs, " ") + `/>`
	xml = insertWorkbookProtection(xml, tag)
	b.fileMap["xl/workbook.xml"] = []byte(xml)
	return nil
}

// insertWorkbookProtection 把 <workbookProtection> 标签插入到合法位置：
// 优先 </workbookPr> 之后，否则 <bookViews 之前，否则 <sheets 之前。
func insertWorkbookProtection(xml, tag string) string {
	if i := strings.Index(xml, "</workbookPr>"); i != -1 {
		end := i + len("</workbookPr>")
		return xml[:end] + tag + xml[end:]
	}
	if i := strings.Index(xml, "<bookViews"); i != -1 {
		return xml[:i] + tag + xml[i:]
	}
	if i := strings.Index(xml, "<sheets"); i != -1 {
		return xml[:i] + tag + xml[i:]
	}
	return xml + tag
}

// hashExcelPassword 计算 Excel 工作簿/工作表保护密码哈希（password verifier，16 位）。
//  1. 对密码 UTF-16 码元自尾到首遍历：循环左移 1 位（15 位宽）后异或该码元；
//  2. 末尾再循环左移 1 位；
//  3. 异或密码长度；
//  4. 异或常量 0xCE4B。
//
// 结果以 4 位大写十六进制写入 workbookPassword 属性。
func hashExcelPassword(password string) string {
	rotl15 := func(h uint16) uint16 {
		return ((h << 1) | (h >> 14)) & 0x7FFF
	}
	var hash uint16
	runes := []rune(password)
	for i := len(runes) - 1; i >= 0; i-- {
		hash = rotl15(hash)
		hash ^= uint16(runes[i])
	}
	hash = rotl15(hash)
	hash ^= uint16(len(runes))
	hash ^= 0xCE4B
	return fmt.Sprintf("%04X", hash&0xFFFF)
}

// ---------- 命名区域（工作簿级全局） ----------

// SetDefinedName 设置一个工作簿级全局命名区域（不绑定特定工作表）。
// refersTo 为引用目标字符串，例如 "Sheet1!$A$1:$D$10" 或 "=SUM(Sheet1!$A$1:$A$10)"。
// 若同名的工作簿级区域已存在则替换；同名但带 localSheetId 的表级区域不受影响。
func (b *Book) SetDefinedName(name, refersTo string) error {
	b.fileMap["xl/workbook.xml"] = setDefinedNameGlobal(b.fileMap["xl/workbook.xml"], name, refersTo)
	return nil
}

// GetDefinedName 读取工作簿级全局命名区域的值（无 localSheetId 时返回 ""）。
func (b *Book) GetDefinedName(name string) string {
	return readDefinedNameGlobal(string(b.fileMap["xl/workbook.xml"]), name)
}

// DeleteDefinedName 删除指定名称的命名区域（全局或表级一并移除）。
func (b *Book) DeleteDefinedName(name string) error {
	xml := string(b.fileMap["xl/workbook.xml"])
	re := regexp.MustCompile(`(?s)<definedName\s+name="` + regexp.QuoteMeta(name) + `"[^>]*>.*?</definedName>`)
	xml = re.ReplaceAllString(xml, "")
	// 容器变空则移除
	if !strings.Contains(xml, "<definedName ") {
		xml = regexp.MustCompile(`(?s)<definedNames\b.*?</definedNames>`).ReplaceAllString(xml, "")
		xml = strings.ReplaceAll(xml, "<definedNames />", "")
	}
	b.fileMap["xl/workbook.xml"] = []byte(xml)
	return nil
}

// setDefinedNameGlobal 在 workbook.xml 中设置/替换工作簿级（无 localSheetId）命名区域。
// 仅匹配没有 localSheetId 的同名 definedName，避免误改同名的表级（带 localSheetId）命名区域。
// name 与 refersTo 均做 XML 转义：名称来自用户输入，含 " 或 < 时若不转义会注入额外节点。
func setDefinedNameGlobal(wbXML []byte, name, refersTo string) []byte {
	content := string(wbXML)
	re := regexp.MustCompile(`(?s)<definedName\s+name="` + regexp.QuoteMeta(safeText(name)) + `"\s*[^>]*>[^<]*</definedName>`)
	newTag := fmt.Sprintf(`<definedName name="%s">%s</definedName>`, safeText(name), safeText(refersTo))
	out := re.ReplaceAllStringFunc(content, func(m string) string {
		// 跳过表级（带 localSheetId）同名区域，只替换工作簿级
		if strings.Contains(m, "localSheetId=") {
			return m
		}
		return newTag
	})
	if out != content {
		return []byte(out)
	}
	// 无同名的：移除空占位 <definedNames /> 后再插入
	content = strings.ReplaceAll(content, "<definedNames />", "")
	tag := newTag
	if idx := strings.LastIndex(content, "</definedNames>"); idx != -1 {
		return []byte(content[:idx] + tag + content[idx:])
	}
	if idx := strings.LastIndex(content, "</sheets>"); idx != -1 {
		closeEnd := idx + len("</sheets>")
		block := `<definedNames>` + tag + `</definedNames>`
		return []byte(content[:closeEnd] + block + content[closeEnd:])
	}
	return []byte(content + `<definedNames>` + tag + `</definedNames>`)
}

// readDefinedNameGlobal 读取工作簿级（无 localSheetId）命名区域的值。
// 若存在同名表级（带 localSheetId）区域，只读取工作簿级那份，避免误读。
func readDefinedNameGlobal(wbXML, name string) string {
	re := regexp.MustCompile(`(?s)<definedName\s+name="` + regexp.QuoteMeta(name) + `"\s*[^>]*>([^<]*)</definedName>`)
	for _, m := range re.FindAllStringSubmatch(wbXML, -1) {
		if strings.Contains(m[0], "localSheetId=") {
			continue // 表级同名，跳过
		}
		return unescapeXML(m[1])
	}
	return ""
}

// ---------- 全局便捷函数（直接操作磁盘文件，与 ops.go / print.go 风格一致） ----------

// ProtectWorkbook 对 filename 启用工作簿结构保护（password 可为空）。
func ProtectWorkbook(filename, password string) error {
	b, err := Open(filename)
	if err != nil {
		return err
	}
	if err := b.ProtectWorkbook(password); err != nil {
		return err
	}
	return b.Save()
}

// UnprotectWorkbook 对 filename 解除工作簿保护。
func UnprotectWorkbook(filename string) error {
	b, err := Open(filename)
	if err != nil {
		return err
	}
	if err := b.UnprotectWorkbook(); err != nil {
		return err
	}
	return b.Save()
}

// GetWorkbookProtection 读取 filename 的工作簿保护状态。
func GetWorkbookProtection(filename string) (WorkbookProtection, error) {
	b, err := Open(filename)
	if err != nil {
		return WorkbookProtection{}, err
	}
	return b.GetWorkbookProtection()
}

// SetDefinedName 在 filename 中设置/替换工作簿级全局命名区域。
func SetDefinedName(filename, name, refersTo string) error {
	b, err := Open(filename)
	if err != nil {
		return err
	}
	if err := b.SetDefinedName(name, refersTo); err != nil {
		return err
	}
	return b.Save()
}

// GetDefinedName 读取 filename 中工作簿级全局命名区域的值。
func GetDefinedName(filename, name string) (string, error) {
	b, err := Open(filename)
	if err != nil {
		return "", err
	}
	return b.GetDefinedName(name), nil
}

// DeleteDefinedName 删除 filename 中指定名称的命名区域。
func DeleteDefinedName(filename, name string) error {
	b, err := Open(filename)
	if err != nil {
		return err
	}
	if err := b.DeleteDefinedName(name); err != nil {
		return err
	}
	return b.Save()
}
