// Package excelgo 提供在不依赖 Excel/WPS 的前提下，操作 .xlsx 工作簿的能力，
// 目前已实现工作表复制、跨工作簿合并（保留页面布局、图片、形状、分页符、打印属性等）。
//
// 后续将逐步扩展为覆盖读写、样式、行列、图表等常见场景的纯 Go Excel 工具库。
//
// 设计目标：
//   - 纯 Go 实现，不调用任何外部 Office 组件；
//   - 输出是符合 OOXML 规范的 .xlsx，可被 Excel / WPS 正常打开；
//   - 通过 MediaStrategy 控制复制体与其图片媒体的耦合方式。
package excelgo

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
)

// MediaStrategy 决定复制体与其图片/媒体部件的耦合方式。
type MediaStrategy int

const (
	// MediaShared 复制体与原表共享同一媒体文件（xl/media/imageN.png）。
	// 优点：文件体积最小；缺点：两表强耦合，删除一方媒体会使另一方失图。
	// 这是 WPS「图片嵌入单元格」场景下的轻量语义。
	MediaShared MediaStrategy = iota

	// MediaIndependent 复制体拥有独立的媒体副本（如 imageN+1.png），
	// 两个工作表完全解耦，复制体可单独编辑、删除而不影响原表。
	MediaIndependent
)

// Options 控制复制行为。使用 NewDefaultOptions 获得默认值，再用函数式选项调整。
type Options struct {
	// Media 控制媒体（图片等）的复制策略，默认 MediaShared。
	Media MediaStrategy
	// Suffix 追加到新工作表名后的后缀，默认 "_copy"。
	Suffix string
}

// DefaultOptions 返回推荐默认配置：共享媒体、后缀 "_copy"。
func DefaultOptions() Options {
	return Options{
		Media:  MediaShared,
		Suffix: "_copy",
	}
}

// Option 是函数式选项，用于调整 Options。
type Option func(*Options)

// WithMedia 设置媒体复制策略。
func WithMedia(s MediaStrategy) Option {
	return func(o *Options) { o.Media = s }
}

// WithSuffix 设置新工作表名后缀（默认 "_copy"）。
func WithSuffix(s string) Option {
	return func(o *Options) { o.Suffix = s }
}

// CopySheet 将 src 工作簿中由 sheetRef（名称或 1 基索引）指定的工作表，
// 复制为同名 + Suffix 的新工作表，写入 dst 路径。
//
//	sheetRef 可以是工作表名称（如 "Sheet1"）或 1 基数字索引（如 "1"）。
//
// 默认行为见 DefaultOptions；可通过 opts 调整媒体策略等。
func CopySheet(src, dst, sheetRef string, opts ...Option) error {
	o := DefaultOptions()
	for _, fn := range opts {
		fn(&o)
	}
	if o.Suffix == "" {
		o.Suffix = "_copy"
	}

	fileMap, err := readZipToMap(src)
	if err != nil {
		return err
	}

	// 解析 workbook.xml（仅读取信息）
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return err
	}

	// 查找源工作表
	var srcSheet *Sheet
	srcIndex := -1
	if idx, err := strconv.Atoi(sheetRef); err == nil {
		if idx >= 1 && idx <= len(wb.Sheets.Sheet) {
			srcSheet = &wb.Sheets.Sheet[idx-1]
			srcIndex = idx - 1
		} else {
			return fmt.Errorf("工作表索引 %d 超出范围 (1-%d)", idx, len(wb.Sheets.Sheet))
		}
	} else {
		for i, s := range wb.Sheets.Sheet {
			if s.Name == sheetRef {
				srcSheet = &wb.Sheets.Sheet[i]
				srcIndex = i
				break
			}
		}
		if srcSheet == nil {
			return fmt.Errorf("找不到名称为 %q 的工作表", sheetRef)
		}
	}

	// 获取源工作表文件路径：以 RID→rels 解析为唯一权威依据（WPS 删改后文件名与顺序解耦）。
	// 解析失败（rel s 缺失/RID 不匹配/目标文件不存在）直接报错，不按位置命名臆测/兜底。
	srcSheetFile, rerr := resolveWorksheetFile(fileMap, srcSheet.RID)
	if rerr != nil {
		return rerr
	}

	// 复制工作表
	newSheetNumber := getMaxSheetNumber(fileMap) + 1
	newSheetFile, err := copyWorksheetWithRelationships(fileMap, srcSheetFile, newSheetNumber, o.Media)
	if err != nil {
		return err
	}

	// 生成新 ID
	newSheetID := getMaxSheetID(wb) + 1
	wbRels := &Relationships{}
	if err := getXMLFromMap(fileMap, "xl/_rels/workbook.xml.rels", wbRels); err != nil {
		return err
	}
	newRId := fmt.Sprintf("rId%d", getMaxRId(wbRels)+1)

	newSheetName := srcSheet.Name + o.Suffix

	// 更新 workbook.xml
	wbData := fileMap["xl/workbook.xml"]
	wbData = insertSheetInWorkbookXML(wbData, newSheetName, newSheetID, newRId)
	fileMap["xl/workbook.xml"] = wbData

	// 更新 workbook.xml.rels
	wbRelsData := fileMap["xl/_rels/workbook.xml.rels"]
	wbRelsData = insertRelationshipInRels(
		wbRelsData,
		newRId,
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		strings.TrimPrefix(newSheetFile, "xl/"),
	)
	fileMap["xl/_rels/workbook.xml.rels"] = wbRelsData

	// 更新 [Content_Types].xml
	ctData := fileMap["[Content_Types].xml"]
	ctData = insertOverrideInContentTypes(
		ctData,
		"/"+newSheetFile,
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
	)
	// 为新复制的部件添加 Override（根据常见路径）—— insertOverrideInContentTypes 幂等，安全。
	for name := range fileMap {
		if strings.HasPrefix(name, "xl/drawings/") && strings.HasSuffix(name, ".xml") {
			ctData = insertOverrideInContentTypes(ctData, "/"+name, "application/vnd.openxmlformats-officedocument.drawing+xml")
		} else if strings.HasPrefix(name, "xl/charts/") && strings.HasSuffix(name, ".xml") {
			ctData = insertOverrideInContentTypes(ctData, "/"+name, "application/vnd.openxmlformats-officedocument.chart+xml")
		} else if strings.HasPrefix(name, "xl/comments/") && strings.HasSuffix(name, ".xml") {
			ctData = insertOverrideInContentTypes(ctData, "/"+name, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
		}
	}
	fileMap["[Content_Types].xml"] = ctData

	// 复制源工作表的打印区域（_xlnm.Print_Area）。
	// 注意：insertSheetInWorkbookXML 只改写字节，wb 结构体仍是插入前状态，
	// 因此新表（被追加在末尾）的 0 基序号等于插入前的工作表数量。
	newSheetIndex := len(wb.Sheets.Sheet)
	wbData = duplicatePrintArea(wbData, srcIndex, newSheetIndex, srcSheet.Name, newSheetName)
	fileMap["xl/workbook.xml"] = wbData

	// 写出新文件
	if err := writeMapToZip(dst, fileMap); err != nil {
		return err
	}

	return nil
}

// ---------- XML 辅助结构 ----------

type Workbook struct {
	XMLName xml.Name `xml:"workbook"`
	Sheets  Sheets   `xml:"sheets"`
}

type Sheets struct {
	Sheet []Sheet `xml:"sheet"`
}

type Sheet struct {
	Name    string `xml:"name,attr"`
	SheetID int    `xml:"sheetId,attr"`
	RID     string `xml:"id,attr"`
}

type Relationships struct {
	XMLName      xml.Name       `xml:"Relationships"`
	Relationship []Relationship `xml:"Relationship"`
}

type Relationship struct {
	ID         string `xml:"Id,attr"`
	Type       string `xml:"Type,attr"`
	Target     string `xml:"Target,attr"`
	TargetMode string `xml:"TargetMode,attr,omitempty"`
}

type ContentTypes struct {
	XMLName   xml.Name   `xml:"Types"`
	Overrides []Override `xml:"Override"`
	Defaults  []Default  `xml:"Default"`
}

type Override struct {
	PartName    string `xml:"PartName,attr"`
	ContentType string `xml:"ContentType,attr"`
}

type Default struct {
	Extension   string `xml:"Extension,attr"`
	ContentType string `xml:"ContentType,attr"`
}

// ---------- 基础 ZIP 读写 ----------

func readZipToMap(filename string) (map[string][]byte, error) {
	reader, err := zip.OpenReader(filename)
	if err != nil {
		return nil, err
	}
	defer reader.Close()

	fileMap := make(map[string][]byte)
	for _, f := range reader.File {
		if f.FileInfo().IsDir() {
			continue
		}
		rc, err := f.Open()
		if err != nil {
			return nil, err
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return nil, err
		}
		fileMap[f.Name] = data
	}
	return fileMap, nil
}

func writeMapToZip(filename string, fileMap map[string][]byte) error {
	outFile, err := os.Create(filename)
	if err != nil {
		return err
	}
	defer outFile.Close()

	writer := zip.NewWriter(outFile)
	defer writer.Close()

	for name, data := range fileMap {
		w, err := writer.Create(name)
		if err != nil {
			return err
		}
		_, err = w.Write(data)
		if err != nil {
			return err
		}
	}
	return nil
}

func fileExistsInMap(fileMap map[string][]byte, filename string) bool {
	_, ok := fileMap[filename]
	return ok
}

// ---------- 解析工具 ----------

func parseXML(data []byte, v interface{}) error {
	return xml.Unmarshal(data, v)
}

func getXMLFromMap(fileMap map[string][]byte, filename string, v interface{}) error {
	data, ok := fileMap[filename]
	if !ok {
		return fmt.Errorf("文件 %s 不存在", filename)
	}
	return parseXML(data, v)
}

// ---------- 字符串插入函数（保留命名空间） ----------

func insertSheetInWorkbookXML(workbookXML []byte, sheetName string, sheetID int, rid string) []byte {
	content := string(workbookXML)
	// workbook.xml 中工作表关系属性使用带 r: 命名空间前缀的 r:id，必须保留前缀，
	// 否则 Excel 无法解析工作表与关系的对应（文件判为损坏）。
	sheetTag := fmt.Sprintf(`<sheet name="%s" sheetId="%d" r:id="%s"/>`, sheetName, sheetID, rid)
	insertPos := strings.LastIndex(content, "</sheets>")
	if insertPos == -1 {
		insertPos = strings.Index(content, "<sheets>")
		if insertPos == -1 {
			return workbookXML
		}
		insertPos += len("<sheets>")
		return []byte(content[:insertPos] + sheetTag + content[insertPos:])
	}
	return []byte(content[:insertPos] + sheetTag + content[insertPos:])
}

func insertRelationshipInRels(relsXML []byte, id, relType, target string) []byte {
	content := string(relsXML)
	relTag := fmt.Sprintf(`<Relationship Id="%s" Type="%s" Target="%s"/>`, id, relType, target)
	insertPos := strings.LastIndex(content, "</Relationships>")
	if insertPos == -1 {
		return relsXML
	}
	return []byte(content[:insertPos] + relTag + content[insertPos:])
}

func insertOverrideInContentTypes(ctXML []byte, partName, contentType string) []byte {
	content := string(ctXML)
	// 幂等：若已存在相同 PartName 的 Override，直接返回，避免产生重复条目导致损坏
	if strings.Contains(content, fmt.Sprintf(`PartName="%s"`, partName)) {
		return ctXML
	}
	overrideTag := fmt.Sprintf(`<Override PartName="%s" ContentType="%s"/>`, partName, contentType)
	insertPos := strings.LastIndex(content, "</Types>")
	if insertPos == -1 {
		return ctXML
	}
	return []byte(content[:insertPos] + overrideTag + content[insertPos:])
}

// 复制源工作表的打印区域（_xlnm.Print_Area）到新工作表。
// 打印区域在 workbook.xml 中以 definedName 形式存在，localSheetId 为 0 基工作表序号，
// 引用范围形如 "Sheet1!$A$1:$J$37"。复制时需把 localSheetId 改为新表序号、把表名改为新表名。
// 若源表无打印区域则原样返回。
func duplicatePrintArea(wbXML []byte, srcIndex, newIndex int, srcName, newName string) []byte {
	content := string(wbXML)
	// 匹配源表的打印区域 definedName
	re := regexp.MustCompile(`(?s)<definedName\s+name="_xlnm\.Print_Area"\s+localSheetId="` + strconv.Itoa(srcIndex) + `"[^>]*>([^<]*)</definedName>`)
	m := re.FindStringSubmatch(content)
	if m == nil {
		return wbXML
	}
	srcRange := m[1]
	// 把范围中的源表名替换为新表名（表名后跟 !）
	newRange := replaceSheetNameInRange(srcRange, srcName, newName)
	newTag := fmt.Sprintf(`<definedName name="_xlnm.Print_Area" localSheetId="%d">%s</definedName>`, newIndex, newRange)
	// 插在源打印区域 definedName 之后
	srcTagEnd := strings.Index(content, m[0]) + len(m[0])
	return []byte(content[:srcTagEnd] + newTag + content[srcTagEnd:])
}

// 将引用范围（如 "Sheet1!$A$1:$J$37"）中的工作表名替换为新名。
// 仅当范围以 srcName 开头并紧跟 ! 时才替换，避免误伤普通单元格引用。
func replaceSheetNameInRange(rng, srcName, newName string) string {
	prefix := srcName + "!"
	if strings.HasPrefix(rng, prefix) {
		return newName + "!" + rng[len(prefix):]
	}
	return rng
}

// ---------- 路径处理 ----------

// 计算 zip 内两个路径的相对路径（统一使用正斜杠）
func zipPathRel(base, target string) (string, error) {
	// 转换为系统路径格式进行计算
	baseSys := filepath.FromSlash(base)
	targetSys := filepath.FromSlash(target)
	rel, err := filepath.Rel(baseSys, targetSys)
	if err != nil {
		return "", err
	}
	// 转换回正斜杠
	return filepath.ToSlash(rel), nil
}

// 解析相对路径，返回 zip 内的绝对路径
func resolveTarget(baseDir, target string) string {
	target = strings.ReplaceAll(target, "\\", "/")
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(target, "/")
	}
	return path.Clean(path.Join(baseDir, target))
}

// 生成不冲突的文件名
// 正确处理形如 "drawing1" 的带数字序号名称：剥离末尾连续数字作为索引基，
// 从 max+1 递增寻找空闲名，避免把 "drawing1" 变成 "drawing11"。
func generateUniqueName(fileMap map[string][]byte, baseDir, baseName, ext string) string {
	candidate := path.Join(baseDir, baseName+ext)
	if !fileExistsInMap(fileMap, candidate) {
		return candidate
	}
	// 剥离 baseName 末尾的连续数字，得到前缀与起始序号
	prefix := baseName
	num := 0
	i := len(baseName)
	for i > 0 && baseName[i-1] >= '0' && baseName[i-1] <= '9' {
		i--
	}
	if i > 0 && i < len(baseName) {
		prefix = baseName[:i]
		if n, err := strconv.Atoi(baseName[i:]); err == nil {
			num = n
		}
	}
	// 从 num+1 开始寻找空闲名
	for next := num + 1; ; next++ {
		candidate = path.Join(baseDir, fmt.Sprintf("%s%d%s", prefix, next, ext))
		if !fileExistsInMap(fileMap, candidate) {
			return candidate
		}
	}
}

// ---------- 核心复制逻辑 ----------

// 提取工作表 XML 中所有内联 r:embed="rIdX" 引用的关系 ID（WPS 单元格内嵌图片、
// 以及 xdr:blip 等直接写在工作表内部的图片均通过此方式引用，不在 _rels 文件里）。
func extractInlineEmbedIDs(wsXML string) []string {
	re := regexp.MustCompile(`r:embed="(rId\d+)"`)
	seen := make(map[string]bool)
	var ids []string
	for _, m := range re.FindAllStringSubmatch(wsXML, -1) {
		id := m[1]
		if !seen[id] {
			seen[id] = true
			ids = append(ids, id)
		}
	}
	return ids
}

// 将工作表 XML 中的 r:embed="oldRId" 整体改写为 r:embed="newRId"
func rewriteInlineEmbed(wsXML, oldRId, newRId string) string {
	return strings.ReplaceAll(wsXML, `r:embed="`+oldRId+`"`, `r:embed="`+newRId+`"`)
}

// 将工作表 XML 中的 r:id="oldRId" 整体改写为 r:id="newRId"。
// 覆盖浮动图片 <drawing r:id="..."/>、超链接等通过关系引用工作表部件的场景，
// 复制时必须随关系重映射一并改写，否则 rId 悬空、图片/链接丢失。
func rewriteRIdRef(wsXML, oldRId, newRId string) string {
	return strings.ReplaceAll(wsXML, `r:id="`+oldRId+`"`, `r:id="`+newRId+`"`)
}

// 确保工作表根元素声明内联图片块所需的命名空间前缀。
// WPS 单元格内嵌图片（mc:AlternateContent / x14:picture / a:extLst）依赖这些前缀，
// 缺失会导致 "unbound prefix" 损坏。
func ensureWorksheetNamespaces(wsXML string) string {
	required := []struct {
		prefix string
		uri    string
	}{
		{"r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships"},
		{"mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"},
		{"x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"},
		{"a", "http://schemas.openxmlformats.org/drawingml/2006/main"},
		{"xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"},
	}
	// 定位真正的 <worksheet ...> 起始标签，跳过可能的 <?xml ...?> 声明
	wsStart := strings.Index(wsXML, "<worksheet")
	if wsStart == -1 {
		return wsXML
	}
	rootEnd := strings.Index(wsXML[wsStart:], ">")
	if rootEnd == -1 {
		return wsXML
	}
	rootEnd += wsStart
	root := wsXML[:rootEnd+1]
	modified := false
	for _, ns := range required {
		if !strings.Contains(root, "xmlns:"+ns.prefix+"=") {
			root = root[:len(root)-1] + fmt.Sprintf(` xmlns:%s="%s"`, ns.prefix, ns.uri) + ">"
			modified = true
		}
	}
	if !modified {
		return wsXML
	}
	return root + wsXML[rootEnd+1:]
}

// copyWorksheetWithRelationships 复制单个工作表及其所有关系部件，返回新工作表文件路径。
// media 参数控制图片/媒体部件的复制策略：
//   - MediaShared：复制体与原表共享同一媒体文件（不复制 xl/media/*）；
//   - MediaIndependent：复制体拥有独立媒体副本（复制 xl/media/* 并重写引用）。
func copyWorksheetWithRelationships(fileMap map[string][]byte, srcSheetFile string, newSheetNumber int, media MediaStrategy) (string, error) {
	newSheetFile := fmt.Sprintf("xl/worksheets/sheet%d.xml", newSheetNumber)
	if fileExistsInMap(fileMap, newSheetFile) {
		return "", fmt.Errorf("目标工作表文件 %s 已存在", newSheetFile)
	}

	// 1. 复制工作表 XML（保留源文件命名空间声明，避免内联 x14:picture 等出现 unbound prefix）
	srcWS := string(fileMap[srcSheetFile])
	newWS := srcWS

	// 2. 处理工作表关系文件
	srcRelsFile := "xl/worksheets/_rels/" + path.Base(srcSheetFile) + ".rels"
	hasSrcRels := fileExistsInMap(fileMap, srcRelsFile)

	srcRels := &Relationships{}
	if hasSrcRels {
		if err := getXMLFromMap(fileMap, srcRelsFile, srcRels); err != nil {
			return "", err
		}
	}

	newRelsFile := "xl/worksheets/_rels/" + path.Base(newSheetFile) + ".rels"
	newRels := &Relationships{
		XMLName: xml.Name{Local: "Relationships"},
	}

	// 为每个关系生成唯一的新 rId（rId 作用域限定在单个工作表）
	nextRId := getMaxRId(srcRels) + 1
	newRIdFor := func(oldID string) string {
		id := fmt.Sprintf("rId%d", nextRId)
		nextRId++
		return id
	}
	// 记录旧 rId -> 新 rId 的映射（同时覆盖 rels 与内联引用）
	ridMap := make(map[string]string)

	// handleTarget 复制（或共享）一个关系目标部件，返回新工作表 rels 中应有的 Target 值。
	// 在 MediaIndependent 模式下，会递归复制目标部件自身 rels 所引用的媒体（例如浮动
	// drawing 的 rels 指向的 xl/media/*），使复制体彻底脱离原表、可独立编辑/删除。
	handleTarget := func(rel Relationship) (string, error) {
		targetPath := resolveTarget("xl/worksheets", rel.Target)
		if rel.TargetMode == "External" || !fileExistsInMap(fileMap, targetPath) {
			// 外部引用：保持原样，仅换 rId
			return rel.Target, nil
		}
		// 媒体文件（图片等）按策略处理
		isMedia := strings.HasPrefix(targetPath, "xl/media/")
		if isMedia && media == MediaShared {
			// 共享同一媒体：复用相同 Target，不复制文件
			return rel.Target, nil
		}
		// 独立媒体或一般部件：复制为目标目录中的唯一命名副本
		targetDir := path.Dir(targetPath)
		targetBase := strings.TrimSuffix(path.Base(targetPath), path.Ext(targetPath))
		targetExt := path.Ext(targetPath)
		newTargetPath := generateUniqueName(fileMap, targetDir, targetBase, targetExt)
		fileMap[newTargetPath] = fileMap[targetPath]

		// 复制目标部件的关系文件（如果有），并按策略处理其引用的媒体
		targetRelsFile := path.Join(path.Dir(targetPath), "_rels", path.Base(targetPath)+".rels")
		if fileExistsInMap(fileMap, targetRelsFile) {
			newTargetRelsFile := path.Join(path.Dir(newTargetPath), "_rels", path.Base(newTargetPath)+".rels")
			relsData := fileMap[targetRelsFile]
			if media == MediaIndependent {
				// 重写目标部件 rels 中指向媒体的 Target，使其指向独立副本
				relsData = rewriteMediaTargetsInRels(relsData, fileMap)
			}
			fileMap[newTargetRelsFile] = relsData
		}

		// 计算新 Target 相对于工作表目录的路径
		newTargetRel, err := zipPathRel("xl/worksheets", newTargetPath)
		if err != nil {
			return "", err
		}
		return newTargetRel, nil
	}

	for _, rel := range srcRels.Relationship {
		newTargetRel, err := handleTarget(rel)
		if err != nil {
			return "", err
		}
		newID := newRIdFor(rel.ID)
		ridMap[rel.ID] = newID
		newRels.Relationship = append(newRels.Relationship, Relationship{
			ID:         newID,
			Type:       rel.Type,
			Target:     newTargetRel,
			TargetMode: rel.TargetMode,
		})
	}

	// 2b. 处理工作表内联引用（WPS 单元格内嵌图片 / xdr:blip）：这些 r:embed 直接写在工作表
	// XML 内部，对应源 rels 文件里的 image 关系。
	// 按媒体策略：共享则复用 Target；独立则复制媒体为副本并重写引用。
	if hasSrcRels {
		relsByID := make(map[string]Relationship)
		for _, r := range srcRels.Relationship {
			relsByID[r.ID] = r
		}
		for _, embID := range extractInlineEmbedIDs(srcWS) {
			srcRel, ok := relsByID[embID]
			if !ok {
				continue
			}
			newTargetRel, err := handleTarget(srcRel)
			if err != nil {
				return "", err
			}
			newID := newRIdFor(embID)
			ridMap[embID] = newID
			newRels.Relationship = append(newRels.Relationship, Relationship{
				ID:         newID,
				Type:       srcRel.Type,
				Target:     newTargetRel,
				TargetMode: srcRel.TargetMode,
			})
			newWS = rewriteInlineEmbed(newWS, embID, newID)
		}
	}

	// 2c. 将工作表中所有随关系重映射而改变的 r:id 引用（浮动图片、超链接等）改写为新 rId，
	// 避免复制后 rId 悬空导致图片/链接丢失。
	for oldID, newID := range ridMap {
		if oldID != newID {
			newWS = rewriteRIdRef(newWS, oldID, newID)
		}
	}

	// 2d. 若工作表含内联图片块，确保所需命名空间前缀（x14/mc/a/xdr）已在工作表根元素声明，
	// 否则 Excel/WPS 打开会报 "unbound prefix"。
	if strings.Contains(newWS, "x14:picture") || strings.Contains(newWS, "AlternateContent") {
		newWS = ensureWorksheetNamespaces(newWS)
	}

	// 回写改写后的工作表 XML
	fileMap[newSheetFile] = []byte(newWS)

	// 序列化新关系文件
	relsData, err := xml.MarshalIndent(newRels, "", "  ")
	if err != nil {
		return "", err
	}
	relsData = append([]byte(xml.Header), relsData...)
	// 手动添加默认命名空间
	relsData = bytes.Replace(relsData, []byte("<Relationships>"), []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`), 1)
	fileMap[newRelsFile] = relsData

	return newSheetFile, nil
}

// ---------- 辅助函数 ----------

// rewriteMediaTargetsInRels 处理某部件（如 drawingN.xml）的 rels 文件：将其指向
// xl/media/* 的 Target 复制为独立媒体副本，并把 rels 中的 Target 改写为新路径。
// 用于 MediaIndependent 模式，确保复制体与其引用的媒体完全解耦。
//
// rels 文件的相对基准目录为 xl/drawings（drawing 的 rels 即位于此）。
func rewriteMediaTargetsInRels(relsData []byte, fileMap map[string][]byte) []byte {
	relRels := &Relationships{}
	if err := xml.Unmarshal(relsData, relRels); err != nil {
		// 解析失败则原样返回，不阻断主流程
		return relsData
	}
	changed := false
	for i := range relRels.Relationship {
		rel := &relRels.Relationship[i]
		if rel.TargetMode == "External" {
			continue
		}
		// 媒体判定：Target 形如 ../media/image1.png（相对 xl/drawings），或绝对形式
		if !strings.Contains(rel.Target, "../media/") && !strings.HasPrefix(rel.Target, "/xl/media/") && !strings.HasPrefix(rel.Target, "xl/media/") {
			continue
		}
		mediaAbs := resolveTarget("xl/drawings", rel.Target)
		if !fileExistsInMap(fileMap, mediaAbs) {
			continue
		}
		dir := path.Dir(mediaAbs)
		base := strings.TrimSuffix(path.Base(mediaAbs), path.Ext(mediaAbs))
		ext := path.Ext(mediaAbs)
		newMedia := generateUniqueName(fileMap, dir, base, ext)
		fileMap[newMedia] = fileMap[mediaAbs]
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
	out = bytes.Replace(out, []byte("<Relationships>"), []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`), 1)
	return out
}

func getMaxSheetNumber(fileMap map[string][]byte) int {
	max := 0
	for name := range fileMap {
		if strings.HasPrefix(name, "xl/worksheets/sheet") && strings.HasSuffix(name, ".xml") {
			base := strings.TrimSuffix(path.Base(name), ".xml")
			numStr := strings.TrimPrefix(base, "sheet")
			if n, err := strconv.Atoi(numStr); err == nil && n > max {
				max = n
			}
		}
	}
	return max
}

func getMaxRId(rels *Relationships) int {
	max := 0
	for _, rel := range rels.Relationship {
		if strings.HasPrefix(rel.ID, "rId") {
			n, err := strconv.Atoi(strings.TrimPrefix(rel.ID, "rId"))
			if err == nil && n > max {
				max = n
			}
		}
	}
	return max
}

func getMaxSheetID(wb *Workbook) int {
	max := 0
	for _, s := range wb.Sheets.Sheet {
		if s.SheetID > max {
			max = s.SheetID
		}
	}
	return max
}
