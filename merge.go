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

// WithRename 指定搬入后的统一工作表名（所有源表同名会触发后缀避让）。
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

// MergeWorkbook 把多个源工作表合并进 dst 工作簿（dst 必须已存在，作为目标容器）。
// 合并后每个源表在 dst 中以独立工作表存在；dst 中原有工作表保持不变。
//
//	sources 为源工作表列表；dst 为目标工作簿路径（会被覆盖写入）。
func MergeWorkbook(dst string, sources []SourceRef, opts ...MergeOption) error {
	o := DefaultMergeOptions()
	for _, fn := range opts {
		fn(&o)
	}
	if o.Suffix == "" {
		o.Suffix = "_merge"
	}

	// 读取目标工作簿
	dstMap, err := readZipToMap(dst)
	if err != nil {
		return err
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
		pa, err := mergeOneSheet(dstMap, dstWB, dstWBrels, src, o)
		if err != nil {
			mergeErrors = append(mergeErrors, fmt.Sprintf("合并 %s!%s 失败: %v", src.Workbook, src.Sheet, err))
			continue
		}
		if pa != nil {
			printAreas = append(printAreas, pa)
		}
	}

	if len(mergeErrors) > 0 {
		// 仍尽量写出已成功合并的部分，但返回首个错误以便察觉
		_ = writeMapToZip(dst, dstMap)
		return fmt.Errorf("合并完成但有错误: %s", strings.Join(mergeErrors, "; "))
	}

	// 统一把打印区域写入 workbook.xml 的 <definedNames> 块（位于 </sheets> 之后）。
	// 不在循环内逐表插入，避免把 definedNames 误塞进 <sheets> 内部破坏结构。
	if len(printAreas) > 0 {
		dstMap["xl/workbook.xml"] = appendDefinedNames(dstMap["xl/workbook.xml"], printAreas)
	}

	// 序列化更新后的 workbook.xml.rels 回写（workbook.xml 字节已在 mergeOneSheet 中
	// 就地更新为最终版本）
	dstMap["xl/_rels/workbook.xml.rels"] = dstWBrelsBytes(dstWBrels)

	if err := writeMapToZip(dst, dstMap); err != nil {
		return err
	}
	return nil
}

// printAreaInfo 记录合并后某源表打印区域在目标工作簿中应写入的位置与范围。
type printAreaInfo struct {
	localSheetId int
	newRange     string
}

// mergeOneSheet 把单个源工作表搬入 dstMap（目标工作簿的内存 map）。
// 该函数会就地修改 dstMap（追加部件、更新 styles、Content_Types、workbook 关系），
// 但 workbook.xml / workbook.xml.rels 的结构体改动在循环外统一序列化回写。
// 返回该源表打印区域在目标侧的映射信息（若有），供调用方统一写入 definedNames。
func mergeOneSheet(dstMap map[string][]byte, dstWB *Workbook, dstWBrels *Relationships, src SourceRef, o MergeOptions) (*printAreaInfo, error) {
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
	newStyles, err := merger.Build()
	if err != nil {
		return nil, err
	}
	dstMap["xl/styles.xml"] = newStyles

	// 5. 改写源 worksheet 的 s 索引为新索引（核心：避免格式错乱）
	newWS := remapStyleIndexes(string(srcWSBytes), styleRemap)

	// 6. 复制源工作表的关联部件（drawing/media/comments/charts 等）到目标编号空间
	//    复用 copy.go 的 handleTarget 思路，但源在 srcMap、目标在 dstMap。
	if err := copySheetPartsToTarget(srcMap, dstMap, srcSheetFile, newSheetFile); err != nil {
		return nil, err
	}

	// 写回改写后的 worksheet
	dstMap[newSheetFile] = []byte(newWS)

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

	// 9. 提取源工作表的打印区域，重映射 localSheetId 与表名，返回给调用方统一写入。
	//    目标新表（追加在末尾）的 0 基序号等于当前工作表数量 - 1（因结构体已 append）
	newLocalSheetId := len(dstWB.Sheets.Sheet) - 1
	var pa *printAreaInfo
	if rng := extractPrintAreaRange(srcMap, srcWB, srcIndex, srcName); rng != "" {
		pa = &printAreaInfo{
			localSheetId: newLocalSheetId,
			newRange:     replaceSheetNameInRange(rng, srcName, finalName),
		}
	}

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

	return pa, nil
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
	file := fmt.Sprintf("xl/worksheets/sheet%d.xml", idx+1)
	if !fileExistsInMap(srcMap, file) {
		rels := &Relationships{}
		if err := getXMLFromMap(srcMap, "xl/_rels/workbook.xml.rels", rels); err != nil {
			return "", "", -1, err
		}
		for _, rel := range rels.Relationship {
			if rel.ID == srcWB.Sheets.Sheet[idx].RID &&
				rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" {
				file = resolveTarget("xl", rel.Target)
				break
			}
		}
	}
	if !fileExistsInMap(srcMap, file) {
		return "", "", -1, fmt.Errorf("找不到源工作表文件")
	}
	return file, name, idx, nil
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
func extractPrintAreaRange(srcMap map[string][]byte, srcWB *Workbook, srcIndex int, srcName string) string {
	srcWBdata := srcMap["xl/workbook.xml"]
	re := regexp.MustCompile(`(?s)<definedName\s+name="_xlnm\.Print_Area"\s+localSheetId="` + strconv.Itoa(srcIndex) + `"[^>]*>([^<]*)</definedName>`)
	m := re.FindStringSubmatch(string(srcWBdata))
	if m == nil {
		return ""
	}
	return m[1]
}

// appendDefinedNames 把若干打印区域 definedName 统一写入目标 workbook.xml。
// 规则：若已有 <definedNames> 块，则追加到其闭合标签前；否则在 </sheets> 之后新建块。
// 保证 definedNames 始终位于 <sheets> 之外，避免破坏 workbook 结构。
func appendDefinedNames(wbXML []byte, pas []*printAreaInfo) []byte {
	content := string(wbXML)
	var tags string
	for _, pa := range pas {
		tags += fmt.Sprintf(`<definedName name="_xlnm.Print_Area" localSheetId="%d">%s</definedName>`, pa.localSheetId, pa.newRange)
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
func copySheetPartsToTarget(srcMap, dstMap map[string][]byte, srcSheetFile, newSheetFile string) error {
	srcRelsFile := "xl/worksheets/_rels/" + path.Base(srcSheetFile) + ".rels"
	if !fileExistsInMap(srcMap, srcRelsFile) {
		return nil // 源表无关联部件
	}
	srcRels := &Relationships{}
	if err := getXMLFromMap(srcMap, srcRelsFile, srcRels); err != nil {
		return err
	}

	newRelsFile := "xl/worksheets/_rels/" + path.Base(newSheetFile) + ".rels"
	newRels := &Relationships{XMLName: xml.Name{Local: "Relationships"}}
	newWS := string(dstMap[newSheetFile])
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

		// 若该部件自身还有 rels（如 drawing 指向 media），递归复制并重写
		targetRelsFile := path.Join(path.Dir(targetPath), "_rels", path.Base(targetPath)+".rels")
		if fileExistsInMap(srcMap, targetRelsFile) {
			newTargetRelsFile := path.Join(path.Dir(newTargetPath), "_rels", path.Base(newTargetPath)+".rels")
			relsData := srcMap[targetRelsFile]
			// 重写该 rels 中指向 media 的 Target（复制媒体副本）
			relsData = rewriteMediaTargetsInRelsCross(srcMap, dstMap, relsData)
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
			return err
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
	dstMap[newSheetFile] = []byte(newWS)

	// 序列化新 worksheet rels
	relsData, err := xml.MarshalIndent(newRels, "", "  ")
	if err != nil {
		return err
	}
	relsData = append([]byte(xml.Header), relsData...)
	relsData = bytes.Replace(relsData, []byte("<Relationships>"),
		[]byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`), 1)
	dstMap[newRelsFile] = relsData
	return nil
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
