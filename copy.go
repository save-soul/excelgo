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

// Options 控制复制行为。使用 DefaultOptions 获得默认值，再用函数式选项调整。
type Options struct {
	// Media 控制媒体（图片等）的复制策略，默认 MediaShared。
	Media MediaStrategy
	// Suffix 追加到新工作表名后的后缀，默认 "_copy"。
	// 仅在未通过 newName 参数显式指定表名时生效。
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
// 仅在 CopySheet/CopySheetTo 的 newName 参数为空时生效；
// 需要精确指定表名时直接传 newName 参数。
func WithSuffix(s string) Option {
	return func(o *Options) { o.Suffix = s }
}

// indexOfSheetName 返回具名工作表在 <sheets> 中的 0 基序号，不存在返回 -1。
func indexOfSheetName(sheets []Sheet, name string) int {
	for i, s := range sheets {
		if s.Name == name {
			return i
		}
	}
	return -1
}

// copySheetInMap 在内存 fileMap 中复制 sheetRef 指定的工作表，追加到末尾。
// newName 为空时用「源表名 + Suffix」；非空时直接作为新表名（重名则补数字序号）。
// 不读写磁盘，仅修改传入的 fileMap（调用方负责 Save/SaveAs 或另存）。
// 返回新工作表名称。
func copySheetInMap(fileMap map[string][]byte, sheetRef, newName string, opts ...Option) (string, error) {
	o := DefaultOptions()
	for _, fn := range opts {
		fn(&o)
	}
	if o.Suffix == "" {
		o.Suffix = "_copy"
	}

	// 解析 workbook.xml（仅读取信息）
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return "", err
	}

	// 查找源工作表
	var srcSheet *Sheet
	srcIndex := -1
	if idx, err := strconv.Atoi(sheetRef); err == nil {
		if idx >= 1 && idx <= len(wb.Sheets.Sheet) {
			srcSheet = &wb.Sheets.Sheet[idx-1]
			srcIndex = idx - 1
		} else {
			return "", fmt.Errorf("工作表索引 %d 超出范围 (1-%d)", idx, len(wb.Sheets.Sheet))
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
			return "", fmt.Errorf("找不到名称为 %q 的工作表", sheetRef)
		}
	}

	// 获取源工作表文件路径：以 RID→rels 解析为唯一权威依据（WPS 删改后文件名与顺序解耦）。
	// 解析失败（rels 缺失/RID 不匹配/目标文件不存在）直接报错，不按位置命名臆测/兜底。
	srcSheetFile, rerr := resolveWorksheetFile(fileMap, srcSheet.RID)
	if rerr != nil {
		return "", rerr
	}

	// 复制工作表
	newSheetNumber := getMaxSheetNumber(fileMap) + 1
	newSheetFile, err := copyWorksheetWithRelationships(fileMap, srcSheetFile, newSheetNumber, o.Media)
	if err != nil {
		return "", err
	}

	// 生成新 ID
	newSheetID := getMaxSheetID(wb) + 1
	wbRels := &Relationships{}
	if err := getXMLFromMap(fileMap, "xl/_rels/workbook.xml.rels", wbRels); err != nil {
		return "", err
	}
	newRId := fmt.Sprintf("rId%d", getMaxRId(wbRels)+1)

	if newName != "" {
		if err := validateSheetName(newName); err != nil {
			return "", err
		}
	}
	newSheetName := newName
	if newSheetName == "" {
		newSheetName = srcSheet.Name + o.Suffix
	}
	if indexOfSheetName(wb.Sheets.Sheet, newSheetName) >= 0 {
		for i := 1; ; i++ {
			cand := fmt.Sprintf("%s_%d", newSheetName, i)
			if indexOfSheetName(wb.Sheets.Sheet, cand) < 0 {
				newSheetName = cand
				break
			}
		}
	}

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

	return newSheetName, nil
}

// CopySheet 复制本工作簿内由 sheetRef（名称或 1 基索引）指定的工作表，追加到末尾
// （excelize 式「同工作簿内复制」语义）。
//
// newName 为空字符串时用「源表名 + WithSuffix 后缀」（默认 "_copy"）；非空时直接
// 作为新表名，重名则自动补数字序号避让。
//
// 操作基于内存 fileMap，不直接写盘；如需落盘请调用 Save/SaveAs。返回新工作表名称。
//
//	f, _ := excelgo.Open("./钢筋.xlsx")
//	newName, _ := f.CopySheet("钢筋表", "")          // 沿用默认后缀
//	newName, _ = f.CopySheet("钢筋表", "钢筋表-副本")  // 指定新表名
//	f.Save()
func (b *Book) CopySheet(sheetRef, newName string, opts ...Option) (string, error) {
	return copySheetInMap(b.fileMap, sheetRef, newName, opts...)
}

// CopySheetTo 把本工作簿中 sheetRef 指定的工作表复制到 dstFile：
//   - dstFile == ""：仅在内存中复制（同 CopySheet），由调用方 Save；
//   - dstFile == b.Filename()：在内存中复制并写回原文件；
//   - dstFile 为其它路径：把该工作表「加入」dstFile 工作簿后写回 dstFile，
//     dstFile 原有工作表全部保留，本工作簿不被修改。
//
// newName 为空字符串时用「源表名 + WithSuffix 后缀」（默认 "_copy"）；非空时直接
// 作为新表名，重名则自动补数字序号避让。
//
// 跨文件复制是「只搬这一张表」：源工作簿的其它工作表一律不参与，dst 也不会被
// 整体覆盖。实现上直接在两个内存 map 之间搬运（源侧 b.fileMap → 目标侧 dst 的
// fileMap），不经过临时文件、不做整簿合并；为保证贴入后格式/图片/打印区域正确，
// 会做跨簿必要的适配：合并 styles.xml 并重映射 s 索引、共享字符串内联化、
// 搬移 drawing/media/批注/图表等关联部件并重映射 rId、复制打印区域并重映射表名。
func (b *Book) CopySheetTo(dstFile, sheetRef, newName string, opts ...Option) error {
	if dstFile == "" || dstFile == b.filename {
		if _, err := b.CopySheet(sheetRef, newName, opts...); err != nil {
			return err
		}
		if dstFile == "" {
			return nil
		}
		return b.Save()
	}

	o := DefaultOptions()
	for _, fn := range opts {
		fn(&o)
	}
	if o.Suffix == "" {
		o.Suffix = "_copy"
	}

	// 读取目标工作簿到内存。严格区分「文件不存在」与「文件存在但打不开」：
	//   - 不存在：创建最小空簿作为容器（正常的一次性创建场景）；
	//   - 存在却打不开（损坏 / 加密 / 非 xlsx / 路径写错）：直接报错，绝不覆盖，
	//     否则会静默把用户的数据文件替换成空工作簿。
	dstBook, err := Open(dstFile)
	if err != nil {
		if _, statErr := os.Stat(dstFile); statErr == nil {
			// 文件确实存在，只是无法解析 —— 属于错误，不应覆盖
			return fmt.Errorf("目标文件 %s 已存在但无法读取（可能已损坏、加密或不是 xlsx），已放弃复制以免覆盖原数据: %w", dstFile, err)
		} else if !os.IsNotExist(statErr) {
			return fmt.Errorf("访问目标文件 %s 失败: %w", dstFile, statErr)
		}
		blank, cerr := Create()
		if cerr != nil {
			return fmt.Errorf("创建目标工作簿 %s 失败: %w", dstFile, cerr)
		}
		if serr := blank.SaveAs(dstFile); serr != nil {
			return fmt.Errorf("创建目标工作簿 %s 失败: %w", dstFile, serr)
		}
		dstBook, err = Open(dstFile)
		if err != nil {
			return err
		}
	}

	if _, err := copySheetAcrossMaps(b.fileMap, dstBook.fileMap, sheetRef, newName, o); err != nil {
		return err
	}
	return dstBook.Save()
}

// copySheetAcrossMaps 把 srcMap 中 sheetRef 指定的那一张工作表搬入 dstMap，
// 追加为末尾的新工作表；dstMap 原有内容（含其它工作表）全部保留。
// 与 mergeOneSheet 的区别：只搬「这一张表」，不会把源工作簿整体并入。
// newName 为空时用「源表名 + o.Suffix」。
func copySheetAcrossMaps(srcMap, dstMap map[string][]byte, sheetRef, newName string, o Options) (string, error) {
	// 1. 定位源工作表
	srcWB := &Workbook{}
	if err := getXMLFromMap(srcMap, "xl/workbook.xml", srcWB); err != nil {
		return "", err
	}
	srcSheetFile, srcName, srcIndex, err := locateSheet(srcMap, srcWB, sheetRef)
	if err != nil {
		return "", err
	}

	// 2. 目标侧分配新的工作表文件名（避让已用编号）
	newSheetNum := nextFreeNumber(dstMap, "xl/worksheets/sheet", ".xml")
	newSheetFile := fmt.Sprintf("xl/worksheets/sheet%d.xml", newSheetNum)

	// 3. 样式：收集源表用到的 s 索引，合并进目标 styles.xml 并重映射
	srcWSBytes := srcMap[srcSheetFile]
	usedStyles := collectUsedStyleIndexes(string(srcWSBytes))
	merger, err := NewStylesMerger(dstMap["xl/styles.xml"], srcMap["xl/styles.xml"])
	if err != nil {
		return "", err
	}
	styleRemap, err := merger.Merge(srcMap["xl/styles.xml"], usedStyles)
	if err != nil {
		return "", err
	}
	newStyles, err := merger.Build()
	if err != nil {
		return "", err
	}
	dstMap["xl/styles.xml"] = newStyles

	// 4. 改写 s 索引 + 共享字符串内联化（使新表自包含，不依赖源簿 sharedStrings 索引）
	newWS := remapStyleIndexes(string(srcWSBytes), styleRemap)
	// 4b. 条件格式的差分样式（dxf）也要搬并重映射 dxfId ——
	//它与 cellXf 是**两张独立的表**，只搬 cellXf 会让 dxfId 悬空
	// （openpyxl 抛 IndexError，Excel 判文件损坏）。
	newWS = mergeDxfsAndRemapSheet(dstMap, srcMap, newWS, merger.StyleXfRemap())
	if sst, ok := srcMap["xl/sharedStrings.xml"]; ok {
		newWS = convertSharedStringsToInline(newWS, string(sst))
	}
	// 5. 搬移关联部件（drawing / media / 批注 / 图表 / 超链接等）并重映射 rId。
	// 该函数会重写工作表里的 rId 引用并**返回**新内容 —— 必须在所有改写
	// （s 索引、dxfId、共享字符串）完成后一次性写回，避免中途覆盖。
	rewritten, err := copySheetPartsToTarget(srcMap, dstMap,
		srcSheetFile, newSheetFile, newWS)
	if err != nil {
		return "", err
	}
	dstMap[newSheetFile] = []byte(rewritten)

	// 6. 表名：与同工作簿内 CopySheet 保持一致 —— newName 非空则直接用，
	// 否则「源表名 + Suffix」；若与既有表重名，再补数字序号避让。
	dstWB := &Workbook{}
	if err := getXMLFromMap(dstMap, "xl/workbook.xml", dstWB); err != nil {
		return "", err
	}
	dstRels := &Relationships{}
	if err := getXMLFromMap(dstMap, "xl/_rels/workbook.xml.rels", dstRels); err != nil {
		return "", err
	}
	preferred := newName
	if preferred == "" {
		preferred = srcName + o.Suffix
	}
	// 避让后仍可能超长（后缀追加所致），此处统一校验
	if err := validateSheetName(preferred); err != nil {
		return "", err
	}
	finalName := preferred
	if indexOfSheetName(dstWB.Sheets.Sheet, finalName) >= 0 {
		for i := 1; ; i++ {
			cand := fmt.Sprintf("%s_%d", preferred, i)
			if indexOfSheetName(dstWB.Sheets.Sheet, cand) < 0 {
				finalName = cand
				break
			}
		}
	}

	newSheetID := getMaxSheetID(dstWB) + 1
	newRId := fmt.Sprintf("rId%d", getMaxRId(dstRels)+1)
	dstMap["xl/workbook.xml"] = insertSheetXML(dstMap["xl/workbook.xml"], finalName, newSheetID, newRId)
	dstMap["xl/_rels/workbook.xml.rels"] = insertRelationshipInRels(
		dstMap["xl/_rels/workbook.xml.rels"], newRId,
		"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
		strings.TrimPrefix(newSheetFile, "xl/"))

	// 7. 打印设置：打印区域 + 重复打印行，重映射 localSheetId 与表名后追加
	newLocalSheetId := len(dstWB.Sheets.Sheet) // append 前已含既有表，末尾即新表序号
	if pas := extractPrintFields(srcMap, srcIndex, srcName, newLocalSheetId, finalName); len(pas) > 0 {
		dstMap["xl/workbook.xml"] = appendDefinedNames(dstMap["xl/workbook.xml"], pas)
	}

	// 8. [Content_Types].xml：为新表及新部件补 Override（幂等）
	ctData := insertOverrideInContentTypes(dstMap["[Content_Types].xml"], "/"+newSheetFile,
		"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
	for name := range dstMap {
		switch {
		case strings.HasPrefix(name, "xl/drawings/") && strings.HasSuffix(name, ".xml"):
			ctData = insertOverrideInContentTypes(ctData, "/"+name, "application/vnd.openxmlformats-officedocument.drawing+xml")
		case strings.HasPrefix(name, "xl/charts/") && strings.HasSuffix(name, ".xml"):
			ctData = insertOverrideInContentTypes(ctData, "/"+name, "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
		case strings.HasPrefix(name, "xl/comments/") && strings.HasSuffix(name, ".xml"):
			ctData = insertOverrideInContentTypes(ctData, "/"+name, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")
		}
	}
	dstMap["[Content_Types].xml"] = ctData

	return finalName, nil
}

// CopySheet 包级便捷函数：将 src 工作簿中由 sheetRef（名称或 1 基索引）指定的工作表，
// 复制到 dst 路径的工作簿中（一次性打开→复制→保存）。只搬这一张表：
// src 的其它工作表不参与，dst 原有工作表全部保留。
//
// newName 为空字符串时用「源表名 + WithSuffix 后缀」（默认 "_copy"）；非空时直接
// 作为新表名，重名则自动补数字序号避让。
//
//	excelgo.CopySheet(src, dst, "Sheet1", "汇总表")   // 指定新表名
//	excelgo.CopySheet(src, dst, "Sheet1", "")         // 沿用默认后缀
//
// 等价于 excelgo.Open(src) 后调用 (*Book).CopySheetTo(dst, sheetRef, newName)。
// 适合「单文件一次性操作」场景；若需要在同一工作簿上连续多次操作，推荐改用对象式 API：
//
//	f, _ := excelgo.Open(src)
//	f.CopySheetTo(dst, sheetRef, newName)   // 跨文件
//	// 或 f.CopySheet(sheetRef, newName) 同一工作簿内复制，再 f.Save()
func CopySheet(src, dst, sheetRef, newName string, opts ...Option) error {
	b, err := Open(src)
	if err != nil {
		return err
	}
	return b.CopySheetTo(dst, sheetRef, newName, opts...)
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
		// 先Close 再判err：ReadAll 失败时 rc 仍需释放；
		// 且 zip 条目的读取器不 Close 会泄漏内部 flate 状态。
		// Close 自身的错误在只读场景下无须处理。
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

// insertSheetInWorkbookXML 向 <sheets> 末尾追加一个 <sheet> 标签。
// sheetName 必须做 XML 属性转义：表名允许含& < > " 等字符（Excel 本身也允许），
// 裸写会破坏 workbook.xml 结构，导致整个 <sheets> 解析失败、所有工作表“消失”。
func insertSheetInWorkbookXML(workbookXML []byte, sheetName string, sheetID int, rid string) []byte {
	content := string(workbookXML)
	// 本函数要写 r:id 前缀，故必须保证前缀有声明。
	// openpyxl 生成的工作簿把 xmlns:r 声明在**每个 <sheet> 元素自己身上**（根上没有），
	// 于是「根未声明 + 原有 sheet 各带声明」的文件在改动前是合法的。但只要流程把那些
	// 带声明的 sheet 删光、只剩本函数新插入的 <sheet>，r: 前缀就失去声明，
	// 文件直接损坏（openpyxl: unbound prefix；Excel: 提示修复甚至打不开）。
	// 与 ensureWorksheetNamespaces 同理：写入前先确保声明，而不是事后补救。
	content = ensureWorkbookRelNamespace(content)
	// workbook.xml 中工作表关系属性使用带 r: 命名空间前缀的 r:id，必须保留前缀，
	// 否则 Excel 无法解析工作表与关系的对应（文件判为损坏）。
	sheetTag := fmt.Sprintf(`<sheet name="%s" sheetId="%d" r:id="%s"/>`,
		safeText(sheetName), sheetID, safeText(rid))
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

// ensureWorkbookRelNamespace 在 workbook.xml 根元素上补齐 xmlns:r 声明（幂等）。
//
// 只在**根元素**缺失时补，不动已有声明 —— openpyxl 把声明写在各个 <sheet> 上是合法
// 的，重复声明会让文件变得啰嗦，且可能干扰按字节比对的下游工具。
// 定位根元素时必须跳过<?xml ...?> 声明，否则会往XML 声明里插属性，产出坏文件。
func ensureWorkbookRelNamespace(wbXML string) string {
	// 跳过可能的 <?xml ...?> 声明与 BOM
	start := strings.Index(wbXML, "<workbook")
	if start == -1 {
		return wbXML
	}
	rootEnd := strings.Index(wbXML[start:], ">")
	if rootEnd == -1 {
		return wbXML
	}
	rootEnd += start
	root := wbXML[start : rootEnd+1]
	// 判据必须只看**根元素**。早先用 strings.Contains(wbXML, ...) 全文搜索，
	// 会被<sheet> 上的声明命中而误判"已声明"直接返回 —— 而恰恰在 openpyxl 那种
	// 「声明写在各 <sheet> 上、根上没有」的文件里失效，正是本函数要修的场景。
	if strings.Contains(root, `xmlns:r=`) {
		return wbXML
	}
	newRoot := root[:len(root)-1] + ` xmlns:r="` + relationshipsNSURI + `">`
	return wbXML[:start] + newRoot + wbXML[rootEnd+1:]
}

func insertRelationshipInRels(relsXML []byte, id, relType, target string) []byte {
	content := string(relsXML)
	// 三个属性值均需 XML 转义（target 可能含 & 等字符的部件名）
	relTag := fmt.Sprintf(`<Relationship Id="%s" Type="%s" Target="%s"/>`,
		safeText(id), safeText(relType), safeText(target))
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

// printSheetFields 列出随工作表复制而需要一并搬运的打印设置类 definedName。
var printSheetFields = []string{"_xlnm.Print_Area", "_xlnm.Print_Titles"}

// 复制源工作表的打印设置类 definedName（_xlnm.Print_Area 打印区域、
// _xlnm.Print_Titles 重复打印行/列）到新工作表。
// 这些设置在 workbook.xml 中以 definedName 形式存在，localSheetId 为 0 基工作表序号，
// 引用范围形如 "Sheet1!$A$1:$J$37" / "Sheet1!$1:$1"。复制时需把 localSheetId 改为
// 新表序号、把表名改为新表名。源表没有相应设置则跳过该项；若一项都没有则原样返回。
func duplicatePrintArea(wbXML []byte, srcIndex, newIndex int, srcName, newName string) []byte {
	content := string(wbXML)
	copied := false
	for _, field := range printSheetFields {
		re := regexp.MustCompile(`(?s)<definedName\s+name="` + regexp.QuoteMeta(field) +
			`"\s+localSheetId="` + strconv.Itoa(srcIndex) + `"[^>]*>([^<]*)</definedName>`)
		m := re.FindStringSubmatch(content)
		if m == nil {
			continue
		}
		newRange := replaceSheetNameInRange(m[1], srcName, newName)
		// newRange 已是 XML 转义态（replaceSheetNameInRange 内部处理），此处不可再转义，
		// 否则表名中的 & 会变成 &amp;amp; 双重转义。
		newTag := fmt.Sprintf(`<definedName name="%s" localSheetId="%d">%s</definedName>`,
			safeText(field), newIndex, newRange)
		// 紧跟在该 definedName 之后插入
		idx := strings.Index(content, m[0]) + len(m[0])
		content = content[:idx] + newTag + content[idx:]
		copied = true
	}
	if !copied {
		return wbXML
	}
	return []byte(content)
}

// replaceSheetNameInRange 把引用范围中的工作表名由 srcName 换成 newName。
//
// 关键点一：Excel 的引用范围是**逗号分隔的多段**，每段各自带表名。
// 同时设置「重复打印行 + 重复打印列」时 Excel 自己就会写出多段：
//
//	_xlnm.Print_Titles = '月报模板'!$1:$6,'月报模板'!$A:$B
//	                     ↑ 行标题                ↑ 列标题 —— 两段都是独立引用
//
// 早期实现用 strings.HasPrefix 只处理**第一段**，多段的后半截原样保留，于是复制/
// 改名后留下悬空引用（'月报模板' 那张表已不存在或已是旧名）。故改为逐段处理。
//
// 关键点二：表名在范围里是**带单引号**的（Excel 对含空格/中文的表名必加引号，
// openpyxl 同样如此）。旧实现拿未加引号的 srcName 做前缀匹配，带引号的真实数据
// 直接匹配不上 —— 结果连单段都没换。故带引号与不带引号两种写法都要认。
//
// 匹配严格按**段边界**：只处理以 srcName（或 'srcName'）紧跟 '!' 开头的段，
// 表名为另一名字前缀时（src=月报模板、引用的是 月报模板X）不会被误伤。
//
// srcName/newName 传入的是「原始（未转义）」表名；rng 来自 workbook.xml 字节，
// 其中的表名是 XML 转义后的形式。表名含 & < > " 时两者不等价，因此匹配与替换
// 都在转义后的空间进行，避免双重转义或替换失败。
func replaceSheetNameInRange(rng, srcName, newName string) string {
	escSrc, escDst := safeText(srcName), safeText(newName)
	segments := strings.Split(rng, ",")
	for i, seg := range segments {
		trimmed := strings.TrimSpace(seg)
		// 带引号写法：'表名'!
		if q := "'" + escSrc + "'!"; strings.HasPrefix(trimmed, q) {
			segments[i] = "'" + escDst + "'!" + trimmed[len(q):]
			continue
		}
		// 无引号写法：表名!
		if p := escSrc + "!"; strings.HasPrefix(trimmed, p) {
			segments[i] = escDst + "!" + trimmed[len(p):]
		}
	}
	return strings.Join(segments, ",")
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

// relationshipsNSURI 是 OOXML 关系（r: 前缀）绑定的命名空间 URI。
// 集中定义，避免在多处硬编码字符串 —— 拼错一个字符就是unbound prefix 级损坏。
const relationshipsNSURI = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

// spreadsheetMainNS 是 SpreadsheetML 主命名空间（workbook/worksheet 根元素的默认 xmlns）。
const spreadsheetMainNS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

// 确保工作表根元素声明内联图片块所需的命名空间前缀。
// WPS 单元格内嵌图片（mc:AlternateContent / x14:picture / a:extLst）依赖这些前缀，
// 缺失会导致 "unbound prefix" 损坏。
func ensureWorksheetNamespaces(wsXML string) string {
	required := []struct {
		prefix string
		uri    string
	}{
		{"r", relationshipsNSURI},
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

		// 表格部件带**工作簿级唯一约束**（id 与 displayName 都必须唯一）。
		// 上面是字节级搬运，源表与副本会拿到相同的 id/名字，导致 openpyxl 报
		// "Table with name X already exists"、Excel 判文件损坏。故搬完立即重映射。
		if strings.HasPrefix(newTargetPath, "xl/tables/") && strings.HasSuffix(newTargetPath, ".xml") {
			fileMap[newTargetPath] = reidentifyTable(fileMap, newTargetPath)
		}

		// 复制目标部件的关系文件（如果有），并按策略处理其引用的媒体
		//
		// 必须**递归**：关系链有多级（sheet -> drawing -> chart -> 嵌入工作簿）。
		// 只处理一层的话，drawing 的 rels 被原样搬完就不管了，chartN.xml 从未被复制，
		// 而 rels 仍指向它 —— 悬空关系，Excel 判定文件损坏。
		targetRelsFile := path.Join(path.Dir(targetPath), "_rels", path.Base(targetPath)+".rels")
		if fileExistsInMap(fileMap, targetRelsFile) {
			newTargetRelsFile := path.Join(path.Dir(newTargetPath), "_rels", path.Base(newTargetPath)+".rels")
			relsData := fileMap[targetRelsFile]
			if media == MediaIndependent {
				// 重写目标部件 rels 中指向媒体的 Target，使其指向独立副本
				relsData = rewriteMediaTargetsInRels(relsData, fileMap)
			}
			// 递归搬运该部件自己引用的下游部件（drawing -> chart 等）。
			// 无论共享还是独立策略都要复制：chart 体积小，且"共享同一份 chart"
			// 语义上就错了（改副本的图表会连带改源表）。
			newRels, rerr := copyDownstreamParts(fileMap, fileMap[targetRelsFile],
				path.Dir(targetPath), path.Dir(newTargetPath))
			if rerr != nil {
				return "", rerr
			}
			fileMap[newTargetRelsFile] = newRels
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

// copyDownstreamParts 递归复制 relsData 所引用的下游部件，并把 rels 里的
// Target 改写为指向新副本。
//
// 为什么必须递归：OOXML 的关系链是多级的 ——
//
//	sheet -> drawing -> chart -> （可选：嵌入工作簿 / colors / style）
//
// 只处理第一层的话，drawing 的 rels 会被原样复制，而它指向的 chartN.xml
// 从未被复制，于是留下悬空关系，Excel/WPS 判定文件损坏。
//
// srcDir/newDir 分别是该部件在源侧与目标侧的目录（用于解析与重算相对 Target）。
// 部件名在目标侧统一重新生成唯一名，避免与既有部件冲突。
// visited 记录已处理过的"源部件绝对路径"，防止环形引用导致无限递归。
func copyDownstreamParts(fm map[string][]byte, relsData []byte, srcDir, newDir string) ([]byte, error) {
	return copyDownstreamPartsVisited(fm, relsData, srcDir, newDir, map[string]bool{})
}

func copyDownstreamPartsVisited(fm map[string][]byte, relsData []byte,
	srcDir, newDir string, visited map[string]bool) ([]byte, error) {
	rels := &Relationships{}
	if err := xml.Unmarshal(relsData, rels); err != nil {
		// 解析失败就原样返回：宁可保留原样，也不要产出半成品关系
		return relsData, nil
	}
	changed := false
	for i := range rels.Relationship {
		rel := &rels.Relationship[i]
		if rel.TargetMode == "External" || rel.Target == "" {
			continue
		}
		srcPath := resolveTarget(srcDir, rel.Target)
		if !fileExistsInMap(fm, srcPath) || visited[srcPath] {
			continue
		}
		// 媒体不在这里复制：drawing 自身 rels 里的 media 已由 MediaShared /
		// MediaIndependent 策略处理过（共享则不复制，独立则复制并重写 Target）。
		// 再复制一遍会产生无引用的孤儿部件。
		if strings.HasPrefix(srcPath, "xl/media/") {
			continue
		}
		visited[srcPath] = true

		// 在目标目录生成唯一名并复制内容
		// 下游部件保留**自身的规范目录**（chart 必须在 xl/charts/ 下）。
		// 不能放到"引用方"的目录里 —— OOXML 的部件路径是约定好的，
		// 放错位置会让 Excel 找不到部件。
		dstDir := path.Dir(srcPath)
		base := strings.TrimSuffix(path.Base(srcPath), path.Ext(srcPath))
		ext := path.Ext(srcPath)
		dstPath := generateUniqueName(fm, dstDir, base, ext)
		fm[dstPath] = fm[srcPath]

		// 表格部件有工作簿级唯一性约束（id / displayName）
		if strings.HasPrefix(dstPath, "xl/tables/") && strings.HasSuffix(dstPath, ".xml") {
			fm[dstPath] = reidentifyTable(fm, dstPath)
		}

		// 递归该部件自己的 rels
		subRels := path.Join(dstDir, "_rels", path.Base(srcPath)+".rels")
		if fileExistsInMap(fm, subRels) {
			newSubRels, err := copyDownstreamPartsVisited(fm, fm[subRels],
				dstDir, dstDir, visited)
			if err != nil {
				return nil, err
			}
			fm[path.Join(dstDir, "_rels", path.Base(dstPath)+".rels")] = newSubRels
		}

		// 重算指向新副本的相对 Target（相对于"引用方"的新目录）
		newRel, err := zipPathRel(newDir, dstPath)
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
