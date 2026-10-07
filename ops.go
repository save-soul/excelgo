package excelgo

// ops.go 提供工作簿/工作表层面的就地编辑能力：
//   - 工作表生命周期：新建 NewSheet、删除 DeleteSheet、移动 MoveSheet、改名 RenameSheet
//   - 单元格读写：GetCell / SetCellValue（及 SetCellStr/Int/Bool/Formula 等细分格式）
//   - 图片插入：AddPicture（浮动，可设锚点位置/偏移/缩放）、AddCellPicture（WPS 单元格内嵌）
//
// 设计原则：纯 Go、不依赖 Office；所有操作基于「读入内存 map → 修改 → 写回」，
// 仅改动必要部件（workbook.xml / rels / content types / 对应 worksheet / media / drawing），
// 绝不触碰其他工作表的内容、样式、布局、图片，保证原有数据格式不被破坏。

import (
	"fmt"
	"path"
	"regexp"
	"strconv"
	"strings"
)

// ---------- 工作表定位辅助 ----------

// locateSheetInMap 在 fileMap 的工作簿结构中定位工作表，返回其信息。
//   - file:     worksheet XML 路径（如 "xl/worksheets/sheet3.xml"）
//   - name:     工作表名
//   - index:    0 基序号
//   - sheetIdx: 在 wb.Sheets.Sheet 中的下标（与 index 相同，便于改结构体）
//
// worksheetRelType 是 worksheet 关系类型常量。
const worksheetRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

// locateSheetInMap 在 fileMap 的工作簿结构中定位工作表，返回其信息。
//   - file:     worksheet XML 路径（如 "xl/worksheets/sheet3.xml"）
//   - name:     工作表名
//   - index:    0 基序号（在 wb.Sheets.Sheet 中的下标）
//   - sheetIdx: 同 index，便于改结构体
//
// 工作表文件由其 RID 经 workbook.xml.rels 唯一确定，与 <sheet> 在列表中的顺序/编号毫无关系。
// WPS 等软件在删除、新建、重排工作表后，xl/worksheets/sheet{N}.xml 的文件名编号与列表顺序会
// 解耦（例如第 1 个 <sheet> 实际指向 sheet3.xml）。因此本函数以 rels 解析为唯一权威依据；
// 若 RID 无法在 rels 中解析、或解析出的文件不存在，直接报错，绝不按位置命名臆测/兜底，
// 以确保定位结果唯一确定、不会产生错文件。
func locateSheetInMap(fileMap map[string][]byte, wb *Workbook, sheetRef string) (file, name string, index, sheetIdx int, err error) {
	if len(wb.Sheets.Sheet) == 0 {
		return "", "", 0, 0, fmt.Errorf("工作簿中没有工作表")
	}
	idx := -1
	if n, e := strconv.Atoi(sheetRef); e == nil {
		if n >= 1 && n <= len(wb.Sheets.Sheet) {
			idx = n - 1
		} else {
			return "", "", 0, 0, fmt.Errorf("工作表索引 %d 超出范围 (1-%d)", n, len(wb.Sheets.Sheet))
		}
	} else {
		for i, s := range wb.Sheets.Sheet {
			if s.Name == sheetRef {
				idx = i
				break
			}
		}
		if idx == -1 {
			return "", "", 0, 0, fmt.Errorf("找不到名称为 %q 的工作表", sheetRef)
		}
	}
	s := wb.Sheets.Sheet[idx]

	// 唯一权威路径：通过 RID 在 workbook.xml.rels 中解析目标 worksheet 文件。
	// 解析失败（rel s 缺失/RID 不匹配/目标文件不存在）一律直接报错，不臆测。
	file, err = resolveWorksheetFile(fileMap, s.RID)
	if err != nil {
		return "", "", 0, 0, err
	}
	return file, s.Name, idx, idx, nil
}

// resolveWorksheetFile 依据工作表 RID 在 workbook.xml.rels 中查找其 worksheet 文件。
// 这是确定工作表文件位置的唯一权威方式；任何基于 "第 N 个 sheet 即 sheet{N+1}.xml" 的假设
// 在 WPS 删改过的工作簿中都不成立。找不到或目标文件缺失时返回明确错误。
func resolveWorksheetFile(fileMap map[string][]byte, rid string) (string, error) {
	rels := &Relationships{}
	if err := getXMLFromMap(fileMap, "xl/_rels/workbook.xml.rels", rels); err != nil {
		return "", fmt.Errorf("解析 workbook.xml.rels 失败：%w", err)
	}
	for _, rel := range rels.Relationship {
		if rel.ID == rid && rel.Type == worksheetRelType {
			file := resolveTarget("xl", rel.Target)
			if !fileExistsInMap(fileMap, file) {
				return "", fmt.Errorf("工作表关系 %q 指向的文件 %q 不存在于压缩包内", rid, file)
			}
			return file, nil
		}
	}
	return "", fmt.Errorf("在 workbook.xml.rels 中找不到 RID=%q 的 worksheet 关系", rid)
}

// ---------- NewSheet：新建工作表 ----------

// NewSheet 在工作簿末尾新建一个名为 name 的空白工作表，返回其 1 基序号。
// 若同名工作表已存在，name 会自动追加后缀避让。新建过程仅向 workbook.xml、其 rels、
// [Content_Types].xml 追加必要条目，并写入一个最小但合法的空白 worksheet，
// 不影响已有工作表的任何内容。
func NewSheet(filename, name string) (int, error) {
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return 0, err
	}
	file, idx, err := newSheetInMap(fileMap, name)
	if err != nil {
		return 0, err
	}
	if err := writeMapToZip(filename, fileMap); err != nil {
		return 0, err
	}
	_ = file
	return idx, nil
}

// newSheetInMap 在内存 fileMap 上新建名为 name 的空白工作表（不复写磁盘）。
// 返回新 worksheet 的 zip 内路径与 1 基序号。同名自动追加后缀避让。
func newSheetInMap(fileMap map[string][]byte, name string) (file string, idx int, err error) {
	// 新建工作表的统一入口（AddSheet / NewSheet 均走此处），在此集中校验表名
	if err := validateSheetName(name); err != nil {
		return "", 0, err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return "", 0, err
	}
	if len(wb.Sheets.Sheet) == 0 {
		return "", 0, fmt.Errorf("工作簿缺少工作表定义，无法新建")
	}

	// 避让同名
	finalName := name
	for exists := false; ; {
		exists = false
		for _, s := range wb.Sheets.Sheet {
			if s.Name == finalName {
				exists = true
				break
			}
		}
		if !exists {
			break
		}
		finalName = name + "_new"
	}

	// 新 worksheet 编号（避让已用编号）
	newNum := nextFreeNumber(fileMap, "xl/worksheets/sheet", ".xml")
	newSheetFile := fmt.Sprintf("xl/worksheets/sheet%d.xml", newNum)
	rels := &Relationships{}
	if err := getXMLFromMap(fileMap, "xl/_rels/workbook.xml.rels", rels); err != nil {
		return "", 0, err
	}
	newRId := fmt.Sprintf("rId%d", getMaxRId(rels)+1)
	newSheetID := getMaxSheetID(wb) + 1

	// 1) 写入空白 worksheet
	fileMap[newSheetFile] = []byte(blankWorksheet())

	// 2) workbook.xml 追加 <sheet>
	fileMap["xl/workbook.xml"] = insertSheetInWorkbookXML(
		fileMap["xl/workbook.xml"], finalName, newSheetID, newRId)

	// 3) workbook.xml.rels 追加关系
	fileMap["xl/_rels/workbook.xml.rels"] = insertRelationshipInRels(
		fileMap["xl/_rels/workbook.xml.rels"],
		newRId,
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		newSheetFile[len("xl/"):],
	)

	// 4) Content_Types 追加 Override
	fileMap["[Content_Types].xml"] = insertOverrideInContentTypes(
		fileMap["[Content_Types].xml"],
		"/"+newSheetFile,
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
	)

	return newSheetFile, len(wb.Sheets.Sheet) + 1, nil
}

// blankWorksheet 返回最小合法的空白 worksheet XML（含严格必要命名空间）。
func blankWorksheet() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
		`<sheetData></sheetData></worksheet>`
}

// ---------- DeleteSheet：删除工作表 ----------

// DeleteSheet 删除由 sheetRef（名称或 1 基索引）指定的工作表，并清理其 worksheet 文件、
// 工作表级 rels、在 workbook.xml 的关系条目、以及该表专属的 drawing/media 部件（仅当
// 这些媒体不被其他表引用时才删除，避免误删共享图片）。其他工作表完全不受影响。
func DeleteSheet(filename, sheetRef string) error {
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return err
	}
	if err := deleteSheetInMap(fileMap, sheetRef); err != nil {
		return err
	}
	return writeMapToZip(filename, fileMap)
}

// deleteSheetInMap 在内存 fileMap 上删除 sheetRef 指定的工作表并清理其专属部件
// （不复写磁盘）。其他工作表完全不受影响。
func deleteSheetInMap(fileMap map[string][]byte, sheetRef string) error {
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return err
	}
	if len(wb.Sheets.Sheet) <= 1 {
		return fmt.Errorf("工作簿至少需保留一个工作表，不能删除最后一个")
	}

	file, _, index, sheetIdx, err := locateSheetInMap(fileMap, wb, sheetRef)
	if err != nil {
		return err
	}
	rid := wb.Sheets.Sheet[sheetIdx].RID

	// 1) 收集该表专属部件（worksheet 自身、其 rels、其 drawing、其独占 media）
	//    先统计其余工作表引用了哪些 media，避免删除共享图片。
	otherMedia, err := collectReferencedMediaExcept(fileMap, wb, sheetIdx)
	if err != nil {
		return err
	}

	// 待删除的 fileMap key
	toDelete := map[string]bool{file: true}
	relsFile := "xl/worksheets/_rels/" + path.Base(file) + ".rels"
	if fileExistsInMap(fileMap, relsFile) {
		toDelete[relsFile] = true
		// 该表 rels 指向的 drawing / comments 等
		sheetRels := &Relationships{}
		if err := getXMLFromMap(fileMap, relsFile, sheetRels); err == nil {
			for _, rel := range sheetRels.Relationship {
				target := resolveTarget("xl/worksheets", rel.Target)
				addCascade(fileMap, target, toDelete, otherMedia)
			}
		}
	}
	// 工作表直接内嵌图片（WPS）引用的 media 也在 worksheet rels 里，上面已覆盖；
	// 再扫一遍 worksheet 内联 r:embed 以防遗漏
	for _, emb := range extractInlineEmbedIDs(string(fileMap[file])) {
		if sheetRels := loadRels(fileMap, relsFile); sheetRels != nil {
			for _, rel := range sheetRels.Relationship {
				if rel.ID == emb {
					target := resolveTarget("xl/worksheets", rel.Target)
					addCascade(fileMap, target, toDelete, otherMedia)
				}
			}
		}
	}

	// 2) 删除文件
	for k := range toDelete {
		delete(fileMap, k)
	}

	// 3) 从 workbook.xml 移除 <sheet>
	fileMap["xl/workbook.xml"] = removeSheetFromWorkbookXML(fileMap["xl/workbook.xml"], rid)

	// 4) 从 workbook.xml.rels 移除关系
	fileMap["xl/_rels/workbook.xml.rels"] = removeRelationshipFromRels(
		fileMap["xl/_rels/workbook.xml.rels"], rid)

	// 5) 清理该表打印区域（definedName）中 localSheetId 的重映射：
	//    删除后其余表序号前移，需要把 localSheetId > index 的全部减 1（保持引用正确）。
	fileMap["xl/workbook.xml"] = reindexDefinedNamesAfterDelete(fileMap["xl/workbook.xml"], index)

	// 6) Content_Types：删除该 worksheet 的 Override（其他部件若不再被引用也可清，但
	//    为安全仅删除 worksheet 自身；drawing/media 若被共享则保留）。
	fileMap["[Content_Types].xml"] = removeOverrideForPart(
		fileMap["[Content_Types].xml"], "/"+file)

	return nil
}

// loadRels 安全加载关系文件，失败返回 nil。
func loadRels(fileMap map[string][]byte, relsFile string) *Relationships {
	if !fileExistsInMap(fileMap, relsFile) {
		return nil
	}
	r := &Relationships{}
	if err := getXMLFromMap(fileMap, relsFile, r); err != nil {
		return nil
	}
	return r
}

// collectReferencedMediaExcept 返回除 exceptIdx 之外所有工作表引用的 media 绝对路径集合。
// 任一工作表文件无法通过其 RID 在 rels 中确定时，直接返回错误（不跳过、不臆测），
// 以保证共享媒体判定结果唯一确定、不会因「猜不到」而误删被共享的图片。
func collectReferencedMediaExcept(fileMap map[string][]byte, wb *Workbook, exceptIdx int) (map[string]bool, error) {
	seen := make(map[string]bool)
	for i, s := range wb.Sheets.Sheet {
		if i == exceptIdx {
			continue
		}
		// 通过 RID 解析该表真正的工作表文件（与列表顺序解耦，不能用 sheet{i+1}.xml 臆测）
		sf, err := resolveWorksheetFile(fileMap, s.RID)
		if err != nil {
			return nil, fmt.Errorf("无法确定工作表 %q 的文件位置：%w", s.Name, err)
		}
		sr := loadRels(fileMap, "xl/worksheets/_rels/"+path.Base(sf)+".rels")
		if sr == nil {
			continue
		}
		for _, rel := range sr.Relationship {
			t := resolveTarget("xl/worksheets", rel.Target)
			if strings.HasPrefix(t, "xl/media/") {
				seen[t] = true
			}
			// 递归：drawing 的 rels 引用
			dr := loadRels(fileMap, path.Join(path.Dir(t), "_rels", path.Base(t)+".rels"))
			if dr != nil {
				for _, drel := range dr.Relationship {
					mt := resolveTarget(path.Dir(t), drel.Target)
					if strings.HasPrefix(mt, "xl/media/") {
						seen[mt] = true
					}
				}
			}
		}
	}
	return seen, nil
}

// addCascade 把 target 部件加入待删除集合，并级联其 rels 引用的子部件
// （drawing → media），但若某 media 在 otherMedia 中（被其他表引用）则保留。
func addCascade(fileMap map[string][]byte, target string, toDelete, otherMedia map[string]bool) {
	if target == "" || !fileExistsInMap(fileMap, target) {
		return
	}
	if strings.HasPrefix(target, "xl/media/") {
		if otherMedia[target] {
			return // 被其他表引用，保留
		}
		toDelete[target] = true
		return
	}
	toDelete[target] = true
	tr := loadRels(fileMap, path.Join(path.Dir(target), "_rels", path.Base(target)+".rels"))
	if tr == nil {
		return
	}
	for _, rel := range tr.Relationship {
		child := resolveTarget(path.Dir(target), rel.Target)
		addCascade(fileMap, child, toDelete, otherMedia)
	}
}

// ---------- MoveSheet：移动工作表 ----------

// MoveSheet 把 sheetRef 指定的工作表移动到 1 基序号 toIndex 处（1 表示最前）。
// 仅重排 workbook.xml 的 <sheet> 顺序及其 rels 中对应条目顺序，不影响任何工作表内容。
func MoveSheet(filename, sheetRef string, toIndex int) error {
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return err
	}
	if toIndex < 1 || toIndex > len(wb.Sheets.Sheet) {
		return fmt.Errorf("目标位置 %d 超出范围 (1-%d)", toIndex, len(wb.Sheets.Sheet))
	}
	_, _, _, sheetIdx, err := locateSheetInMap(fileMap, wb, sheetRef)
	if err != nil {
		return err
	}
	fileMap["xl/workbook.xml"] = moveSheetInWorkbookXML(fileMap["xl/workbook.xml"], sheetIdx, toIndex-1)
	if err := writeMapToZip(filename, fileMap); err != nil {
		return err
	}
	return nil
}

// ---------- RenameSheet：改名 ----------

// RenameSheet 把名为 oldName 的工作表改名为 newName。workbook.xml 的 <sheet name="...">
// 是唯一需要修改处；若同名则追加后缀避让。其他内容不受影响。
func RenameSheet(filename, oldName, newName string) error {
	// 与 (*Book).RenameSheet 保持一致地校验表名（31 字符上限 / 非法字符）
	if err := validateSheetName(newName); err != nil {
		return err
	}
	fileMap, err := readZipToMap(filename)
	if err != nil {
		return err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return err
	}
	found := false
	for _, s := range wb.Sheets.Sheet {
		if s.Name == oldName {
			found = true
			break
		}
	}
	if !found {
		return fmt.Errorf("找不到名为 %q 的工作表", oldName)
	}
	// 避让同名：逐个加序号，而非固定后缀（固定 "_ren" 遇同名会死循环）
	finalNew := newName
	for i := 1; indexOfSheetName(wb.Sheets.Sheet, finalNew) >= 0; i++ {
		finalNew = fmt.Sprintf("%s_%d", newName, i)
	}
	if err := validateSheetName(finalNew); err != nil {
		return err
	}
	fileMap["xl/workbook.xml"] = renameSheetInWorkbookXML(fileMap["xl/workbook.xml"], oldName, finalNew)
	if err := writeMapToZip(filename, fileMap); err != nil {
		return err
	}
	return nil
}

// SetSheetVisible 设置 filename 中 sheetRef 工作表的可见性状态（见 Book.SetSheetVisible）。
// state 取 SheetStateVisible / SheetStateHidden / SheetStateVeryHidden。
func SetSheetVisible(filename, sheetRef, state string) error {
	b, err := Open(filename)
	if err != nil {
		return err
	}
	if err := b.SetSheetVisible(sheetRef, state); err != nil {
		return err
	}
	return b.Save()
}

// GetSheetVisible 读取 filename 中 sheetRef 工作表的可见性状态。
func GetSheetVisible(filename, sheetRef string) (string, error) {
	b, err := Open(filename)
	if err != nil {
		return "", err
	}
	return b.GetSheetVisible(sheetRef)
}

// ---------- workbook.xml 字节层操作 ----------

func removeSheetFromWorkbookXML(wbXML []byte, rid string) []byte {
	content := string(wbXML)
	// 匹配 <sheet ...> 中含 r:id="rid" 的整个标签（属性任意顺序）
	re := regexp.MustCompile(`<sheet\b[^>]*\br:id="` + regexp.QuoteMeta(rid) + `"[^>]*/?>`)
	newContent := re.ReplaceAllString(content, "")
	return []byte(newContent)
}

func removeRelationshipFromRels(relsXML []byte, rid string) []byte {
	content := string(relsXML)
	re := regexp.MustCompile(`(?s)<Relationship\b[^>]*\bId="` + regexp.QuoteMeta(rid) + `"[^>]*/?>`)
	return []byte(re.ReplaceAllString(content, ""))
}

// reindexDefinedNamesAfterDelete 在删除序号为 delIdx（0 基）的工作表后，修正所有表级
// （带 localSheetId）命名区域的引用，保持引用正确：
//   - 属于被删表的（localSheetId==delIdx，如该表专属的 _xlnm.Print_Area / _xlnm.Print_Titles
//     及用户自建的表级命名区域）直接移除；
//   - localSheetId>delIdx 的全部减 1（删除后其余表序号前移）；
//   - 工作簿级（无 localSheetId）命名区域不受影响。
func reindexDefinedNamesAfterDelete(wbXML []byte, delIdx int) []byte {
	content := string(wbXML)
	re := regexp.MustCompile(`(?s)<definedName\b([^>]*)>(.*?)</definedName>`)
	out := re.ReplaceAllStringFunc(content, func(m string) string {
		sub := re.FindStringSubmatch(m)
		if sub == nil {
			return m
		}
		attrs := sub[1]
		inner := sub[2]
		// 仅处理带 localSheetId 的表级命名区域
		lsRe := regexp.MustCompile(`localSheetId="(\d+)"`)
		lsM := lsRe.FindStringSubmatch(attrs)
		if lsM == nil {
			return m
		}
		id, _ := strconv.Atoi(lsM[1])
		if id == delIdx {
			return "" // 属于被删表，移除该命名区域（含其打印区域/重复打印标题）
		}
		if id > delIdx {
			id--
		}
		newAttrs := lsRe.ReplaceAllString(attrs, fmt.Sprintf(`localSheetId="%d"`, id))
		return fmt.Sprintf(`<definedName%s>%s</definedName>`, newAttrs, inner)
	})
	return []byte(out)
}

// moveSheetInWorkbookXML 把第 fromIdx（0 基）个 <sheet> 移动到 toIdx（0 基）位置。
func moveSheetInWorkbookXML(wbXML []byte, fromIdx, toIdx int) []byte {
	sheets := extractSheetTags(wbXML)
	if fromIdx < 0 || fromIdx >= len(sheets) || toIdx < 0 || toIdx >= len(sheets) {
		return []byte(wbXML)
	}
	moved := sheets[fromIdx]
	// 从切片移除
	sheets = append(sheets[:fromIdx], sheets[fromIdx+1:]...)
	// 插入到目标位置
	if toIdx >= len(sheets) {
		sheets = append(sheets, moved)
	} else {
		sheets = append(sheets[:toIdx], append([]string{moved}, sheets[toIdx:]...)...)
	}
	// 重建 <sheets> 内容
	return []byte(replaceSheetsInner(wbXML, sheets))
}

// renameSheetInWorkbookXML 把 <sheet name="oldName"> 改为新名。
// oldName/newName 都是「用户视角」的原始表名；匹配与写入时都必须做 XML 属性转义，
// 否则表名含 & < > " 时会因文件里存的是 A&amp;B 而匹配不到 A&B，导致重命名静默失效。
func renameSheetInWorkbookXML(wbXML []byte, oldName, newName string) []byte {
	content := string(wbXML)
	re := regexp.MustCompile(`(<sheet\b[^>]*\bname=")(` + regexp.QuoteMeta(safeAttr(oldName)) + `)(")`)
	return []byte(re.ReplaceAllString(content, `${1}`+safeAttr(newName)+`${3}`))
}

// ---------- sheet 标签提取/重建 ----------

// extractSheetTags 提取 <sheets> 内部所有 <sheet .../> 标签原文（按顺序）。
func extractSheetTags(wbXML []byte) []string {
	content := string(wbXML)
	open := strings.Index(content, "<sheets>")
	close := strings.Index(content, "</sheets>")
	if open == -1 || close == -1 || close < open {
		return nil
	}
	inner := content[open+len("<sheets>") : close]
	re := regexp.MustCompile(`<sheet\b[^>]*/?>`)
	return re.FindAllString(inner, -1)
}

// replaceSheetsInner 用新的 sheet 标签列表替换 <sheets> 内部内容，其余不动。
func replaceSheetsInner(wbXML []byte, sheets []string) []byte {
	content := string(wbXML)
	open := strings.Index(content, "<sheets>")
	close := strings.Index(content, "</sheets>")
	if open == -1 || close == -1 || close < open {
		return wbXML
	}
	openEnd := open + len("<sheets>")
	return []byte(content[:openEnd] + strings.Join(sheets, "") + content[close:])
}

// removeOverrideForPart 从 [Content_Types].xml 删除指定 PartName 的 Override 条目。
func removeOverrideForPart(ctXML []byte, partName string) []byte {
	content := string(ctXML)
	re := regexp.MustCompile(`<Override\b[^>]*\bPartName="` + regexp.QuoteMeta(partName) + `"[^>]*/?>`)
	return []byte(re.ReplaceAllString(content, ""))
}
