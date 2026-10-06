package excelgo

// merge.go 提供跨工作簿合并工作表的能力：把多个源工作簿中指定的工作表，
// 作为独立工作表搬移到同一个目标工作簿中（合并后各表仍彼此独立）。
//
// 与 CopySheet（同工作簿内复制）的区别：
//   - 源与目标分属不同工作簿（不同 fileMap），各自有独立的 styles.xml；
//   - 必须合并 styles 并对单元格 s 索引做重映射（见 styles.go），否则格式错乱；
//   - 部件（drawing/media/comments/charts/打印区域等）随工作表一起搬移并重编号。
//
// 设计目标（与 CopySheet 一致）：纯 Go、输出符合 OOXML、可被 Excel/WPS 打开、
// 保留页面布局/图片/形状/分页符/打印属性，样式合并去重避免数量溢出。

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"path"
	"regexp"
	"strconv"
	"strings"
)

// MergeOptions 控制合并行为。
type MergeOptions struct {
	// Rename 决定搬入目标后工作表的命名方式。
	// 若为空，则沿用源表原名；若目标已存在同名，则追加后缀避免冲突。
	Rename string
	// Suffix 当发生命名冲突时追加的后缀，默认 "_merge"。
	Suffix string
}

// DefaultMergeOptions 返回默认合并配置。
func DefaultMergeOptions() MergeOptions {
	return MergeOptions{Suffix: "_merge"}
}

// MergeOption 是函数式选项。
type MergeOption func(*MergeOptions)

// WithRename 指定合并后使用的工作表名（作用于所有源表）。
// 同名时按 WithMergeSuffix 后缀避让。
// 注意：CopySheet/CopySheetTo 的表名请直接用其 newName 参数，与本选项无关。
func WithRename(name string) MergeOption {
	return func(o *MergeOptions) { o.Rename = name }
}

// WithMergeSuffix 指定命名冲突时追加的后缀（默认 "_merge"）。
func WithMergeSuffix(s string) MergeOption {
	return func(o *MergeOptions) { o.Suffix = s }
}

// SourceRef 描述一个待合并的源工作表：来自哪个工作簿、哪个工作表。
type SourceRef struct {
	// Workbook 源工作簿路径。
	Workbook string
	// Sheet 源工作表名称或 1 基索引。
	Sheet string
}

// mergeIntoMap 把多个源工作表合并进 dstMap（目标工作簿的内存 map），dstMap 必须已含
// xl/styles.xml 与 xl/workbook.xml（作为合并容器）。不读写磁盘，仅修改 dstMap。
func mergeIntoMap(dstMap map[string][]byte, sources []SourceRef, opts ...MergeOption) error {
	o := DefaultMergeOptions()
	for _, fn := range opts {
		fn(&o)
	}
	if o.Suffix == "" {
		o.Suffix = "_merge"
	}

	if _, ok := dstMap["xl/styles.xml"]; !ok {
		return fmt.Errorf("目标工作簿缺少 xl/styles.xml，无法合并样式")
	}
	if _, ok := dstMap["xl/workbook.xml"]; !ok {
		return fmt.Errorf("目标工作簿缺少 xl/workbook.xml")
	}

	// 解析目标 workbook.xml / rels，用于追加 sheet 与关系
	dstWB := &Workbook{}
	if err := getXMLFromMap(dstMap, "xl/workbook.xml", dstWB); err != nil {
		return err
	}
	dstWBrels := &Relationships{}
	if err := getXMLFromMap(dstMap, "xl/_rels/workbook.xml.rels", dstWBrels); err != nil {
		return err
	}

	var mergeErrors []string
	var printAreas []*printAreaInfo

	for _, src := range sources {
		pas, err := mergeOneSheet(dstMap, dstWB, dstWBrels, src, o)
		if err != nil {
			mergeErrors = append(mergeErrors, fmt.Sprintf("合并 %s!%s 失败: %v", src.Workbook, src.Sheet, err))
			continue
		}
		if len(pas) > 0 {
			printAreas = append(printAreas, pas...)
		}
	}

	// 统一把打印区域写入 workbook.xml 的 <definedNames> 块（位于 </sheets> 之后）。
	// 不在循环内逐表插入，避免把 definedNames 误塞进 <sheets> 内部破坏结构。
	if len(printAreas) > 0 {
		dstMap["xl/workbook.xml"] = appendDefinedNames(dstMap["xl/workbook.xml"], printAreas)
	}

	// 序列化更新后的 workbook.xml.rels 回写。workbook.xml 字节已在 mergeOneSheet 中
	// 就地更新为最终版本；rels 必须在此统一回写，否则即便部分合并失败，已成功合并的
	// 工作表也会因缺少关系条目而落盘为损坏文件。
	dstMap["xl/_rels/workbook.xml.rels"] = dstWBrelsBytes(dstWBrels)

	// 至此，已成功合并的部分（workbook.xml / rels / 打印区域）均已回写，即便存在
	// 局部合并错误也尽量写出可用结果；仍返回首个错误以便调用方察觉。
	if len(mergeErrors) > 0 {
		return fmt.Errorf("合并完成但有错误: %s", strings.Join(mergeErrors, "; "))
	}

	return nil
}

// Merge 把多个源工作表合并进本工作簿（excelize 式对象 API，目标容器即本工作簿）。
// 合并后每个源表在本工作簿中以独立工作表存在；本工作簿原有工作表保持不变。
//
// 操作基于内存 fileMap，不直接写盘；如需落盘请调用 Save/SaveAs。
//
//	f, _ := excelgo.Open(dst)
//	f.Merge([]excelgo.SourceRef{{Workbook: src, Sheet: "数据A"}})
//	f.Save()
func (b *Book) Merge(sources []SourceRef, opts ...MergeOption) error {
	return mergeIntoMap(b.fileMap, sources, opts...)
}

// MergeWorkbook 包级便捷函数：把多个源工作表合并进 dst 工作簿（dst 必须已存在，作为目标容器），
// 合并后每个源表在 dst 中以独立工作表存在；dst 中原有工作表保持不变。等价于
// excelgo.Open(dst) 后调用 (*Book).Merge(sources) 并 Save。
//
// 适合「单文件一次性操作」场景；若需要在同一工作簿上连续多次操作，推荐改用对象式 API：
//
//	f, _ := excelgo.Open(dst)
//	f.Merge(sources)
//	f.Save()
func MergeWorkbook(dst string, sources []SourceRef, opts ...MergeOption) error {
	b, err := Open(dst)
	if err != nil {
		return err
	}
	merr := b.Merge(sources, opts...)
	// 即使存在局部合并错误（原逻辑也会尽量写回），这里同样写回再返回首个错误。
	if serr := b.Save(); serr != nil {
		if merr != nil {
			return merr
		}
		return serr
	}
	return merr
}

// printAreaInfo 记录合并后某源表打印区域在目标工作簿中应写入的位置与范围。
type printAreaInfo struct {
	localSheetId int
	newRange     string
	// field 为 definedName 名（如 "_xlnm.Print_Area" / "_xlnm.Print_Titles"）。
	// 为空时按 _xlnm.Print_Area 处理，保持既有调用兼容。
	field string
}

// mergeOneSheet 把单个源工作表搬入 dstMap（目标工作簿的内存 map）。
// 该函数会就地修改 dstMap（追加部件、更新 styles、Content_Types、workbook 关系），
// 但 workbook.xml / workbook.xml.rels 的结构体改动在循环外统一序列化回写。
// 返回该源表打印设置（打印区域/重复打印行）在目标侧的映射信息（若有），供调用方统一写入 definedNames。
func mergeOneSheet(dstMap map[string][]byte, dstWB *Workbook, dstWBrels *Relationships, src SourceRef, o MergeOptions) ([]*printAreaInfo, error) {
	// 1. 读取源工作簿
	srcMap, err := readZipToMap(src.Workbook)
	if err != nil {
		return nil, err
	}
	srcWB := &Workbook{}
	if err := getXMLFromMap(srcMap, "xl/workbook.xml", srcWB); err != nil {
		return nil, err
	}

	// 2. 定位源工作表文件
	srcSheetFile, srcName, srcIndex, err := locateSheet(srcMap, srcWB, src.Sheet)
	if err != nil {
		return nil, err
	}

	// 3. 计算目标侧新编号（避让目标已用编号）
	newSheetNum := nextFreeNumber(dstMap, "xl/worksheets/sheet", ".xml")
	newSheetFile := fmt.Sprintf("xl/worksheets/sheet%d.xml", newSheetNum)

	// 4. 收集源 worksheet 中用到的 s 索引，并合并 styles（累积到 dstMap 的 styles.xml）
	srcWSBytes := srcMap[srcSheetFile]
	usedStyles := collectUsedStyleIndexes(string(srcWSBytes))
	dstStyles := dstMap["xl/styles.xml"]
	merger, err := NewStylesMerger(dstStyles, srcMap["xl/styles.xml"])
	if err != nil {
		return nil, err
	}
	styleRemap, err := merger.Merge(srcMap["xl/styles.xml"], usedStyles)
	if err != nil {
		return nil, err
	}
	stylesAfterCellXf, err := merger.Build()
	if err != nil {
		return nil, err
	}
	dstMap["xl/styles.xml"] = stylesAfterCellXf

	// 5. 改写源 worksheet 的 s 索引为新索引（核心：避免格式错乱）
	newWS := remapStyleIndexes(string(srcWSBytes), styleRemap)
	// 5b. 搬运差分样式（dxf）并重映射 dxfId —— 与 copySheetAcrossMaps 共用同一实现，
	// 避免"只在一条路径上补"的疏漏（曾漏过 CopySheetTo，导致同样的悬空 dxfId）。
	newWS = mergeDxfsAndRemapSheet(dstMap, srcMap, newWS, merger.StyleXfRemap())

	// 5.5 将源 worksheet 的共享字符串单元格（t="s"）转换为内联字符串（t="inlineStr"），
	// 使合并表自包含，避免目标工作簿缺失/不一致的 sharedStrings.xml 导致索引越界或串文。
	if sst, ok := srcMap["xl/sharedStrings.xml"]; ok {
		newWS = convertSharedStringsToInline(newWS, string(sst))
	}

	// 6. 复制源工作表的关联部件（drawing/media/comments/charts 等）到目标编号空间
	//    复用 copy.go 的 handleTarget 思路，但源在 srcMap、目标在 dstMap。
	// 搬运关联部件并重写工作表里的 rId 引用。
	// 单一写入点：s 索引 / dxfId / 共享字符串 / rId 全部改写完成后才落盘。
	rewritten, err := copySheetPartsToTarget(srcMap, dstMap,
		srcSheetFile, newSheetFile, newWS)
	if err != nil {
		return nil, err
	}
	dstMap[newSheetFile] = []byte(rewritten)

	// 7. 决定目标工作表名（避让同名）
	finalName := o.Rename
	if finalName == "" {
		finalName = srcName
	}
	finalName = uniqueSheetName(dstWB, finalName, o.Suffix)

	// 8. 追加 workbook.xml 的 <sheet>（字节层面，保留 r: 前缀）与
	//    workbook.xml.rels 的 <Relationship>
	newSheetID := getMaxSheetID(dstWB) + 1
	newRId := fmt.Sprintf("rId%d", getMaxRId(dstWBrels)+1)
	dstWB.Sheets.Sheet = append(dstWB.Sheets.Sheet, Sheet{
		Name:    finalName,
		SheetID: newSheetID,
		RID:     newRId,
	})
	dstWBrels.Relationship = append(dstWBrels.Relationship, Relationship{
		ID:     newRId,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		Target: strings.TrimPrefix(newSheetFile, "xl/"),
	})
	// 字节层面把 <sheet> 追加进 workbook.xml，使下一轮循环 readZipToMap 后能读到
	dstMap["xl/workbook.xml"] = insertSheetXML(dstMap["xl/workbook.xml"], finalName, newSheetID, newRId)

	// 9. 提取源工作表的打印设置（打印区域 + 重复打印行），重映射 localSheetId 与表名，
	//    返回给调用方统一写入。目标新表（追加在末尾）的 0 基序号等于当前表数量 - 1。
	newLocalSheetId := len(dstWB.Sheets.Sheet) - 1
	pas := extractPrintFields(srcMap, srcIndex, srcName, newLocalSheetId, finalName)

	// 10. 为新工作表及新复制的部件补充 [Content_Types].xml Override（幂等）
	ctData := dstMap["[Content_Types].xml"]
	ctData = insertOverrideInContentTypes(ctData, "/"+newSheetFile,
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
	for name := range dstMap {
		if strings.HasPrefix(name, "xl/drawings/") && strings.HasSuffix(name, ".xml") {
			ctData = insertOverrideInContentTypes(ctData, "/"+name,
				"application/vnd.openxmlformats-officedocument.drawing+xml")
		} else if strings.HasPrefix(name, "xl/charts/") && strings.HasSuffix(name, ".xml") {
			ctData = insertOverrideInContentTypes(ctData, "/"+name,
				"application/vnd.openxmlformats-officedocument.chart+xml")
		} else if strings.HasPrefix(name, "xl/comments/") && strings.HasSuffix(name, ".xml") {
			ctData = insertOverrideInContentTypes(ctData, "/"+name,
				"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
		}
	}
	dstMap["[Content_Types].xml"] = ctData

	return pas, nil
}

// ---------- 辅助 ----------

// locateSheet 返回源工作表文件路径、名称、0 基索引。
func locateSheet(srcMap map[string][]byte, srcWB *Workbook, sheetRef string) (string, string, int, error) {
	idx := -1
	name := ""
	if n, err := strconv.Atoi(sheetRef); err == nil {
		if n >= 1 && n <= len(srcWB.Sheets.Sheet) {
			idx = n - 1
			name = srcWB.Sheets.Sheet[idx].Name
		} else {
			return "", "", -1, fmt.Errorf("工作表索引 %d 超出范围", n)
		}
	} else {
		for i, s := range srcWB.Sheets.Sheet {
			if s.Name == sheetRef {
				idx = i
				name = s.Name
				break
			}
		}
		if idx == -1 {
			return "", "", -1, fmt.Errorf("找不到工作表 %q", sheetRef)
		}
	}
	f, ferr := resolveWorksheetFile(srcMap, srcWB.Sheets.Sheet[idx].RID)
	if ferr != nil {
		return "", "", -1, ferr
	}
	return f, name, idx, nil
}

// nextFreeNumber 在 fileMap 中查找前缀 prefix+数字+suffix 形式里未占用的最小正整数编号。
func nextFreeNumber(fileMap map[string][]byte, prefix, suffix string) int {
	used := make(map[int]bool)
	re := regexp.MustCompile(regexp.QuoteMeta(prefix) + `(\d+)` + regexp.QuoteMeta(suffix) + `$`)
	for name := range fileMap {
		if m := re.FindStringSubmatch(name); m != nil {
			if n, err := strconv.Atoi(m[1]); err == nil {
				used[n] = true
			}
		}
	}
	n := 1
	for used[n] {
		n++
	}
	return n
}

// collectUsedStyleIndexes 从 worksheet XML 中提取所有被引用的 s 索引（去重）。
// 单元格形如 <c r="A1" s="3" t="s"/>；也处理 <c s="0"/>（无 s 时默认 0）。
func collectUsedStyleIndexes(wsXML string) []int {
	seen := make(map[int]bool)
	// 匹配 <c ... s="N" ...> 中的 s 属性（任意属性顺序，容错）
	re := regexp.MustCompile(`<c\b[^>]*\bs="(\d+)"`)
	for _, m := range re.FindAllStringSubmatch(wsXML, -1) {
		if n, err := strconv.Atoi(m[1]); err == nil {
			if !seen[n] {
				seen[n] = true
			}
		}
	}
	// 也扫描可能的无 s 单元格（默认样式 0），但只有真的出现才计入
	if !seen[0] {
		// 仅当 worksheet 存在 <c> 且无 s 属性时，默认样式 0 才可能被使用
		cRe := regexp.MustCompile(`<c\b[^>]*>`)
		for _, m := range cRe.FindAllStringSubmatch(wsXML, -1) {
			if !strings.Contains(m[0], `s="`) {
				seen[0] = true
				break
			}
		}
	}
	out := make([]int, 0, len(seen))
	for k := range seen {
		out = append(out, k)
	}
	return out
}

// remapStyleIndexes 把 worksheet XML 中所有 <c ... s="old"> 的 s 属性按 remap 替换为新值。
// 仅替换 remap 中存在的 old 值，未出现的保持不变。
func remapStyleIndexes(wsXML string, remap map[int]int) string {
	re := regexp.MustCompile(`(<c\b[^>]*\bs=")(\d+)(")`)
	return re.ReplaceAllStringFunc(wsXML, func(match string) string {
		sub := re.FindStringSubmatch(match)
		if sub == nil {
			return match
		}
		old, err := strconv.Atoi(sub[2])
		if err != nil {
			return match
		}
		newer, ok := remap[old]
		if !ok {
			return match
		}
		return sub[1] + strconv.Itoa(newer) + sub[3]
	})
}

// convertSharedStringsToInline 将 worksheet 中引用共享字符串的单元格（t="s"）就地
// 转换为内联字符串单元格（t="inlineStr"）。转换后会话表自包含，不再依赖工作簿的
// sharedStrings.xml，从而避免跨工作簿合并时共享字符串索引错位/越界。
//
// 源 worksheet 的共享字符串定义在 sstXML（源工作簿的 xl/sharedStrings.xml）中，
// 按单元格 <v> 内的索引解析出文本后内联写入 <is><t>。库自身即以 inlineStr 写单元格，
// 此转换与既有风格一致；仅当源 worksheet 确实含 t="s" 单元格时才发生改写。
func convertSharedStringsToInline(wsXML, sstXML string) string {
	if !regexp.MustCompile(`t="s"`).MatchString(wsXML) {
		return wsXML
	}
	re := regexp.MustCompile(`(?s)(<c\b[^>]*\bt="s"[^>]*>)\s*<v>(\d+)</v>\s*(</c>)`)
	return re.ReplaceAllStringFunc(wsXML, func(m string) string {
		sm := re.FindStringSubmatch(m)
		if sm == nil {
			return m
		}
		openTag := regexp.MustCompile(`\bt="s"`).ReplaceAllString(sm[1], `t="inlineStr"`)
		idx, err := strconv.Atoi(sm[2])
		if err != nil {
			return m
		}
		text := resolveSharedString(sstXML, idx)
		return fmt.Sprintf(`%s<is><t xml:space="preserve">%s</t></is>%s`, openTag, safeText(text), sm[3])
	})
}

// resolveSharedString 按索引从 sharedStrings.xml 中取出第 idx 个 <si> 的文本内容。
// 支持富文本（多个 <r><t>...</t></r>），拼接所有 <t> 文本节点。
func resolveSharedString(sstXML string, idx int) string {
	siRe := regexp.MustCompile(`(?s)<si\b[^>]*>(.*?)</si>`)
	sis := siRe.FindAllStringSubmatch(sstXML, -1)
	if idx < 0 || idx >= len(sis) {
		return ""
	}
	tRe := regexp.MustCompile(`(?s)<t\b[^>]*>(.*?)</t>`)
	ts := tRe.FindAllStringSubmatch(sis[idx][1], -1)
	var sb strings.Builder
	for _, t := range ts {
		sb.WriteString(t[1])
	}
	return sb.String()
}

// uniqueSheetName 在目标工作簿中生成一个不冲突的工作表名。
func uniqueSheetName(wb *Workbook, base, suffix string) string {
	exists := func(n string) bool {
		for _, s := range wb.Sheets.Sheet {
			if s.Name == n {
				return true
			}
		}
		return false
	}
	if !exists(base) {
		return base
	}
	i := 1
	for {
		cand := base + suffix + strconv.Itoa(i)
		if !exists(cand) {
			return cand
		}
		i++
	}
}

// extractPrintAreaRange 从源工作簿提取源表（0 基 srcIndex，名 srcName）的打印区域引用范围。
// 找不到则返回空字符串。
// extractPrintFields 提取源工作表（0 基 srcIndex，名 srcName）上所有打印设置类
// definedName（打印区域 + 重复打印行/列），返回待写入目标侧的 printAreaInfo 列表。
// newLocalSheetId 为目标工作簿中该表 append 后的 0 基序号，newName 为目标表名。
// 源表没有打印设置时返回 nil。
func extractPrintFields(srcMap map[string][]byte, srcIndex int, srcName string, newLocalSheetId int, newName string) []*printAreaInfo {
	srcWBdata := string(srcMap["xl/workbook.xml"])
	var out []*printAreaInfo
	for _, field := range printSheetFields {
		re := regexp.MustCompile(`(?s)<definedName\s+name="` + regexp.QuoteMeta(field) +
			`"\s+localSheetId="` + strconv.Itoa(srcIndex) + `"[^>]*>([^<]*)</definedName>`)
		m := re.FindStringSubmatch(srcWBdata)
		if m == nil {
			continue
		}
		out = append(out, &printAreaInfo{
			localSheetId: newLocalSheetId,
			newRange:     replaceSheetNameInRange(m[1], srcName, newName),
			field:        field,
		})
	}
	return out
}

// extractPrintAreaRange 提取源工作表的打印区域引用范围；没有则返回空字符串。
func extractPrintAreaRange(srcMap map[string][]byte, srcWB *Workbook, srcIndex int, srcName string) string {
	for _, pa := range extractPrintFields(srcMap, srcIndex, srcName, srcIndex, srcName) {
		if pa.field == "_xlnm.Print_Area" {
			return pa.newRange
		}
	}
	return ""
}

// appendDefinedNames 把若干打印区域 definedName 统一写入目标 workbook.xml。
// 规则：若已有 <definedNames> 块，则追加到其闭合标签前；否则在 </sheets> 之后新建块。
// 保证 definedNames 始终位于 <sheets> 之外，避免破坏 workbook 结构。
func appendDefinedNames(wbXML []byte, pas []*printAreaInfo) []byte {
	content := string(wbXML)
	var tags string
	for _, pa := range pas {
		field := pa.field
		if field == "" {
			field = "_xlnm.Print_Area"
		}
		tags += fmt.Sprintf(`<definedName name="%s" localSheetId="%d">%s</definedName>`,
			safeText(field), pa.localSheetId, pa.newRange)
	}
	// 先移除可能存在的自闭合空占位 <definedNames />（openpyxl 等会在无打印区域时写入），
	// 否则下方“没有 </definedNames> 闭合标签”分支会再新建一个块，导致出现两个块、解析歧义。
	content = strings.ReplaceAll(content, "<definedNames />", "")
	if idx := strings.LastIndex(content, "</definedNames>"); idx != -1 {
		return []byte(content[:idx] + tags + content[idx:])
	}
	// 没有 definedNames 块：插在 </sheets> 闭合标签「之后」（不能在 </sheets> 之前，
	// 否则 definedNames 会被误认为 sheets 的子元素，破坏 workbook 结构）。
	if idx := strings.LastIndex(content, "</sheets>"); idx != -1 {
		closeEnd := idx + len("</sheets>")
		return []byte(content[:closeEnd] + `<definedNames>` + tags + `</definedNames>` + content[closeEnd:])
	}
	// 兜底：追加到末尾
	return []byte(content + `<definedNames>` + tags + `</definedNames>`)
}

// ---------- 序列化回写 workbook ----------

func dstWBrelsBytes(rels *Relationships) []byte {
	data, err := xml.MarshalIndent(rels, "", "  ")
	if err != nil {
		return nil
	}
	out := append([]byte(xml.Header), data...)
	out = bytes.Replace(out, []byte("<Relationships>"),
		[]byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`), 1)
	return out
}

// insertSheetXML 把 <sheet> 标签追加进 workbook.xml 的 <sheets> 末尾。
// （保留 r: 前缀，参见 copy.go 中 insertSheetInWorkbookXML 的说明）
func insertSheetXML(wbXML []byte, sheetName string, sheetID int, rid string) []byte {
	return insertSheetInWorkbookXML(wbXML, sheetName, sheetID, rid)
}

// copySheetPartsToTarget 把源工作表 srcSheetFile（位于 srcMap）的所有关联部件
// （drawing/media/comments/charts 等）复制到目标 fileMap 的 newSheetFile 编号空间下，
// 并重写 worksheet 中的 <drawing r:id> 与内联 r:embed 引用到新 rId。
//
// 跨工作簿场景下媒体必须独立复制（源与目标分属不同 zip，无法共享），因此媒体一律
// 复制为副本并改写引用。绘制对象（drawing）及其依赖的 media 一并迁移，确保浮动图片、
// 形状、图表等随工作表完整保留。
// sheetXML 是待写回的工作表内容（本函数会重写其中的 rId 引用并返回新内容）。
//
// 早期实现里本函数直接写 dstMap[newSheetFile]，而调用方随后又用自己的
// newWS 覆盖一次 —— 外层那次覆盖把这里做完的 rId 重写抹掉了。
// 症状：sheet XML 仍引用源文件的旧 rId（如 rId1），而 rels 里已换成新ID
// （rId5 起），openpyxl 报 "Unknown relationship: rId1"，Excel 报文件损坏。
// 修法：让本函数**返回**改写后的内容，由调用方统一写回，单一写入点。
func copySheetPartsToTarget(srcMap, dstMap map[string][]byte,
	srcSheetFile, newSheetFile string, sheetXML string) (string, error) {
	srcRelsFile := "xl/worksheets/_rels/" + path.Base(srcSheetFile) + ".rels"
	if !fileExistsInMap(srcMap, srcRelsFile) {
		return sheetXML, nil // 源表无关联部件
	}
	srcRels := &Relationships{}
	if err := getXMLFromMap(srcMap, srcRelsFile, srcRels); err != nil {
		return sheetXML, err
	}

	newRelsFile := "xl/worksheets/_rels/" + path.Base(newSheetFile) + ".rels"
	newRels := &Relationships{XMLName: xml.Name{Local: "Relationships"}}
	newWS := sheetXML
	nextRId := getMaxRId(srcRels) + 1
	ridMap := make(map[string]string)

	newRIdFor := func() string {
		id := fmt.Sprintf("rId%d", nextRId)
		nextRId++
		return id
	}

	// copyTarget 把 srcMap 中 rel.Target 指向的部件复制到 dstMap（避让目标编号），
	// 返回新部件相对新工作表目录的路径。媒体与子部件递归复制。
	var copyTarget func(relTarget, baseDir string) (string, error)
	copyTarget = func(relTarget, baseDir string) (string, error) {
		targetPath := resolveTarget(baseDir, relTarget)
		if !fileExistsInMap(srcMap, targetPath) {
			return relTarget, nil // 外部或找不到，保持原样
		}
		targetDir := path.Dir(targetPath)
		base := strings.TrimSuffix(path.Base(targetPath), path.Ext(targetPath))
		ext := path.Ext(targetPath)
		newTargetPath := generateUniqueNameIn(dstMap, targetDir, base, ext)
		dstMap[newTargetPath] = srcMap[targetPath]

		// 表格部件带**工作簿级唯一约束**（id 与 displayName 都必须唯一），
		// 字节级原样搬运会与目标簿里已有的表冲突 —— openpyxl 报
		// "Table with name X already exists"，Excel 同样判损坏。
		// 故搬完立即重映射 id 与 displayName。
		if strings.HasPrefix(newTargetPath, "xl/tables/") && strings.HasSuffix(newTargetPath, ".xml") {
			dstMap[newTargetPath] = reidentifyTable(dstMap, newTargetPath)
		}

		// 若该部件自身还有 rels（drawing -> media/chart），递归复制并重写。
		//
		// 关系链是多级的：sheet -> drawing -> chart。只处理一层时 chartN.xml
		// 从不被复制，而 rels 仍指向它 —— 悬空关系，Excel 判文件损坏。
		targetRelsFile := path.Join(path.Dir(targetPath), "_rels", path.Base(targetPath)+".rels")
		if fileExistsInMap(srcMap, targetRelsFile) {
			newTargetRelsFile := path.Join(path.Dir(newTargetPath), "_rels", path.Base(newTargetPath)+".rels")
			relsData := srcMap[targetRelsFile]
			// 重写该 rels 中指向 media 的 Target（复制媒体副本）
			relsData = rewriteMediaTargetsInRelsCross(srcMap, dstMap, relsData)
			// 递归搬运下游部件（chart 等），并把它们放到目标 map
			var err error
			relsData, err = copyDownstreamPartsCross(srcMap, dstMap, relsData,
				path.Dir(targetPath), path.Dir(newTargetPath), map[string]bool{})
			if err != nil {
				return "", err
			}
			dstMap[newTargetRelsFile] = relsData
		}
		newRel, err := zipPathRel("xl/worksheets", newTargetPath)
		if err != nil {
			return "", err
		}
		return newRel, nil
	}

	for _, rel := range srcRels.Relationship {
		newTargetRel, err := copyTarget(rel.Target, "xl/worksheets")
		if err != nil {
			return sheetXML, err
		}
		nid := newRIdFor()
		ridMap[rel.ID] = nid
		newRels.Relationship = append(newRels.Relationship, Relationship{
			ID:         nid,
			Type:       rel.Type,
			Target:     newTargetRel,
			TargetMode: rel.TargetMode,
		})
	}

	// 处理工作表内联 r:embed（WPS 单元格内嵌图片等）：复用上面已生成的 ridMap
	// （内联引用对应的关系 ID 也在 srcRels 中），统一重写。
	for oldID, newID := range ridMap {
		newWS = rewriteRIdRef(newWS, oldID, newID)
		newWS = rewriteInlineEmbed(newWS, oldID, newID)
	}

	// 内联图片命名空间补全
	if strings.Contains(newWS, "x14:picture") || strings.Contains(newWS, "AlternateContent") {
		newWS = ensureWorksheetNamespaces(newWS)
	}

	// 序列化新 worksheet rels
	relsData, err := xml.MarshalIndent(newRels, "", "  ")
	if err != nil {
		return sheetXML, err
	}
	relsData = append([]byte(xml.Header), relsData...)
	relsData = bytes.Replace(relsData, []byte("<Relationships>"),
		[]byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`), 1)
	dstMap[newRelsFile] = relsData

	// 返回改写后的工作表内容，由调用方统一写回（单一写入点）
	return newWS, nil
}

// generateUniqueNameIn 在指定 fileMap 中寻找空闲名（同 copy.go 的 generateUniqueName，
// 但作用于任意 map 而非强制 xl/ 前缀约束）。
func generateUniqueNameIn(fileMap map[string][]byte, baseDir, baseName, ext string) string {
	return generateUniqueName(fileMap, baseDir, baseName, ext)
}

// rewriteMediaTargetsInRelsCross 跨 fileMap 版：把部件 rels 中指向 media 的 Target
// 复制为 dstMap 中的独立副本，并重写 Target。源在 srcMap、目标在 dstMap。
func rewriteMediaTargetsInRelsCross(srcMap, dstMap map[string][]byte, relsData []byte) []byte {
	relRels := &Relationships{}
	if err := xml.Unmarshal(relsData, relRels); err != nil {
		return relsData
	}
	changed := false
	for i := range relRels.Relationship {
		rel := &relRels.Relationship[i]
		if rel.TargetMode == "External" {
			continue
		}
		if !strings.Contains(rel.Target, "../media/") && !strings.HasPrefix(rel.Target, "/xl/media/") && !strings.HasPrefix(rel.Target, "xl/media/") {
			continue
		}
		mediaAbs := resolveTarget("xl/drawings", rel.Target)
		if !fileExistsInMap(srcMap, mediaAbs) {
			continue
		}
		dir := path.Dir(mediaAbs)
		base := strings.TrimSuffix(path.Base(mediaAbs), path.Ext(mediaAbs))
		ext := path.Ext(mediaAbs)
		newMedia := generateUniqueNameIn(dstMap, dir, base, ext)
		dstMap[newMedia] = srcMap[mediaAbs]
		newRel, err := zipPathRel("xl/drawings", newMedia)
		if err == nil {
			rel.Target = newRel
			changed = true
		}
	}
	if !changed {
		return relsData
	}
	out, err := xml.MarshalIndent(relRels, "", "  ")
	if err != nil {
		return relsData
	}
	out = append([]byte(xml.Header), out...)
	out = bytes.Replace(out, []byte("<Relationships>"),
		[]byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`), 1)
	return out
}

// reidentifyTable 为刚搬进目标簿的表格部件重新分配 id 与 displayName。
//
// OOXML 里 table 的 id 与 displayName 都是**工作簿级唯一**的：
//   - id 重复：消费方按 id 索引表，冲突即错乱；
//   - displayName 重复：openpyxl 直接抛 "Table with name X already exists"，
//     Excel 也会判文件损坏。
//
// 部件搬运是字节级的（原样复制 srcMap 的内容），因此必须显式重映射。
// 名字冲突时追加 _1 / _2 序号，与工作表重名的避让策略一致。
func reidentifyTable(fm map[string][]byte, partName string) []byte {
	s := string(fm[partName])

	// 1) 重新分配 id = 目标簿内现有最大 id + 1
	newID := nextTableContentID(fm)
	s = regexp.MustCompile(`(<table\b[^>]*\bid=")\d+(")`).
		ReplaceAllString(s, "${1}"+strconv.Itoa(newID)+"${2}")

	// 2) name / displayName 冲突时改名。
	//
	//    两个属性都要改，且 openpyxl 判定重名用的是 **name**（不是 displayName）——
	//    只改 displayName 仍会抛 "Table with name X already exists"。
	//    这里的判据是"目标簿里是否已有别的表用了这个名字"。
	cur := tableAttr(s, tagName("name"))
	if cur != "" && tableNameTaken(fm, cur, partName) {
		unique := cur
		for i := 1; ; i++ {
			cand := cur + "_" + strconv.Itoa(i)
			if !tableNameTaken(fm, cand, partName) {
				unique = cand
				break
			}
		}
		esc := safeAttr(unique)
		s = regexp.MustCompile(`(<table\b[^>]*\bname=")[^"]*(")`).
			ReplaceAllString(s, "${1}"+esc+"${2}")
		s = regexp.MustCompile(`(<table\b[^>]*\bdisplayName=")[^"]*(")`).
			ReplaceAllString(s, "${1}"+esc+"${2}")
	}
	return []byte(s)
}

// tableAttrNames 是 tableAttr 允许查询的属性名白名单。
// 这些是 OOXML 里 table 的固定属性名，由库内常量传入，绝不接受外部数据 ——
// 它们会拼进正则，因此必须收口在此白名单内。
var tableAttrNames = map[string]bool{"name": true, "displayName": true, "id": true}

// tableAttr 取出 table 元素上指定属性的值。attr 非白名单值时返回空串。
func tableAttr(tableXML string, attr xmlTagName) string {
	if !tableAttrNames[string(attr)] {
		return ""
	}
	re := regexp.MustCompile(`<table\b[^>]*\b` + regexp.QuoteMeta(string(attr)) + `="([^"]*)"`)
	if m := re.FindStringSubmatch(tableXML); m != nil {
		return m[1]
	}
	return ""
}

// tableNameTaken 判断除 skipPart 之外是否已有表格用了该名字。
func tableNameTaken(fm map[string][]byte, name, skipPart string) bool {
	for p, data := range fm {
		if p == skipPart || !strings.HasPrefix(p, "xl/tables/") {
			continue
		}
		// name 与 displayName 都算：Excel 两者都不允许与他表重复
		if tableAttr(string(data), tagName("name")) == name ||
			tableAttr(string(data), tagName("displayName")) == name {
			return true
		}
	}
	return false
}

// copyDownstreamPartsCross 是 copyDownstreamParts 的跨簿版本：从 srcMap 读、
// 往 dstMap 写。用于 CopySheetTo / MergeWorkbook 这类跨工作簿的部件搬运。
//
// 与同簿版的差异只在"源与目标是两个 map"，语义要求完全一致：
// 关系链必须递归搬运（sheet -> drawing -> chart），且下游部件要放到
// 目标簿里的**同名规范目录**（chart 进 xl/charts/，不能放到引用方目录下）。
func copyDownstreamPartsCross(srcMap, dstMap map[string][]byte, relsData []byte,
	srcDir, newRefDir string, visited map[string]bool) ([]byte, error) {
	rels := &Relationships{}
	if err := xml.Unmarshal(relsData, rels); err != nil {
		return relsData, nil
	}
	changed := false
	for i := range rels.Relationship {
		rel := &rels.Relationship[i]
		if rel.TargetMode == "External" || rel.Target == "" {
			continue
		}
		srcPath := resolveTarget(srcDir, rel.Target)
		if !fileExistsInMap(srcMap, srcPath) || visited[srcPath] {
			continue
		}
		// media 已由 rewriteMediaTargetsInRelsCross 处理，不重复复制
		if strings.HasPrefix(srcPath, "xl/media/") {
			continue
		}
		visited[srcPath] = true

		// 下游部件保留自身规范目录
		dstDir := path.Dir(srcPath)
		base := strings.TrimSuffix(path.Base(srcPath), path.Ext(srcPath))
		ext := path.Ext(srcPath)
		dstPath := generateUniqueNameIn(dstMap, dstDir, base, ext)
		dstMap[dstPath] = srcMap[srcPath]
		if strings.HasPrefix(dstPath, "xl/tables/") && strings.HasSuffix(dstPath, ".xml") {
			dstMap[dstPath] = reidentifyTable(dstMap, dstPath)
		}

		// 递归该部件自己的 rels
		subRels := path.Join(dstDir, "_rels", path.Base(srcPath)+".rels")
		if fileExistsInMap(srcMap, subRels) {
			newSub, err := copyDownstreamPartsCross(srcMap, dstMap, srcMap[subRels],
				dstDir, dstDir, visited)
			if err != nil {
				return nil, err
			}
			dstMap[path.Join(dstDir, "_rels", path.Base(dstPath)+".rels")] = newSub
		}

		// 重算相对"引用方新目录"的 Target
		newRel, err := zipPathRel(newRefDir, dstPath)
		if err != nil {
			return nil, err
		}
		if newRel != rel.Target {
			rel.Target = newRel
			changed = true
		}
	}
	if !changed {
		return relsData, nil
	}
	out, err := xml.MarshalIndent(rels, "", "  ")
	if err != nil {
		return relsData, nil
	}
	out = append([]byte(xml.Header), out...)
	out = bytes.Replace(out, []byte("<Relationships>"),
		[]byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`), 1)
	return out, nil
}
