package excelgo

// picture.go 提供图片插入能力：
//   - AddPicture：浮动图片（drawing），可设置锚定单元格、像素偏移、缩放比例；
//   - AddCellPicture：WPS 单元格内嵌图片（mc:AlternateContent / x14:picture），
//     图片随单元格移动与缩放，贴合「嵌入单元格」语义。
//
// 两种插入都只新增必要部件（media / drawing / rels 条目），并保留 worksheet 已有的
// 其他 drawing、图片、形状、布局，不破坏原内容。

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"image"
	_ "image/jpeg"
	_ "image/png"
	"os"
	"path"
	"regexp"
	"strconv"
	"strings"
)

// EMUPerPixel 96 DPI 下每像素对应的 EMU 数（1 px = 9525 EMU）。
const EMUPerPixel = 9525

// PictureOptions 浮动图片的可选参数（参考 excelize 的 GraphicOptions 思路）。
type PictureOptions struct {
	// ScaleX / ScaleY 缩放比例（相对图片原始尺寸），默认 1.0。
	ScaleX float64
	ScaleY float64
	// Position 锚定位置；为 nil 时默认锚定在 A1、偏移 0。
	Position *PicturePosition
}

// PicturePosition 浮动图片的锚定位置。
type PicturePosition struct {
	// Cell 锚定单元格（如 "A1"），图片左上角对齐该单元格左上角。
	Cell string
	// ColOffset / RowOffset 相对锚定单元格左上角的像素偏移。
	ColOffset int
	RowOffset int
}

// AddPicture 在 sheetRef 工作表的指定位置插入一张浮动图片（来自 picPath 文件）。
// 通过 opts 设置锚定单元格、像素偏移与缩放比例。图片被复制到 xl/media/ 下，
// 并新建/复用 drawing 部件承载，保留工作表原有绘图对象。
func AddPicture(filename, sheetRef, picPath string, opts *PictureOptions) error {
	data, err := os.ReadFile(picPath)
	if err != nil {
		return fmt.Errorf("读取图片失败: %w", err)
	}
	return addPictureBytes(filename, sheetRef, data, opts)
}

// AddPictureFromBytes 同 AddPicture，但直接接收图片字节（便于从内存/网络插入）。
func AddPictureFromBytes(filename, sheetRef string, data []byte, opts *PictureOptions) error {
	return addPictureBytes(filename, sheetRef, data, opts)
}

func addPictureBytes(filename, sheetRef string, data []byte, opts *PictureOptions) error {
	if opts == nil {
		opts = &PictureOptions{ScaleX: 1, ScaleY: 1}
	}
	if opts.ScaleX <= 0 {
		opts.ScaleX = 1
	}
	if opts.ScaleY <= 0 {
		opts.ScaleY = 1
	}
	col, row := 0, 0
	if opts.Position != nil && opts.Position.Cell != "" {
		c, r, err := parseCellRef(opts.Position.Cell)
		if err != nil {
			return err
		}
		col, row = c, r
	}
	colOffPx, rowOffPx := 0, 0
	if opts.Position != nil {
		colOffPx = opts.Position.ColOffset
		rowOffPx = opts.Position.RowOffset
	}

	// 图片原始像素尺寸
	bounds, _, err := image.DecodeConfig(bytes.NewReader(data))
	if err != nil {
		return fmt.Errorf("无法解析图片尺寸（需 PNG/JPEG）: %w", err)
	}
	extCx := int(float64(bounds.Width) * opts.ScaleX * EMUPerPixel)
	extCy := int(float64(bounds.Height) * opts.ScaleY * EMUPerPixel)

	fileMap, err := readZipToMap(filename)
	if err != nil {
		return err
	}
	wb := &Workbook{}
	if err := getXMLFromMap(fileMap, "xl/workbook.xml", wb); err != nil {
		return err
	}
	file, _, _, _, err := locateSheetInMap(fileMap, wb, sheetRef)
	if err != nil {
		return err
	}

	// 1) 媒体文件
	mediaExt := imageExtFromData(data)
	mediaPath := generateUniqueName(fileMap, "xl/media", "image", mediaExt)
	fileMap[mediaPath] = data

	// 2) drawing 部件：复用该表已有的 drawing，或新建一个
	sheetRelsFile := "xl/worksheets/_rels/" + path.Base(file) + ".rels"
	sheetRels := loadRels(fileMap, sheetRelsFile)
	drawingRID := ""
	drawingFile := ""
	drawingRelIDForMedia := ""
	if sheetRels != nil {
		for _, rel := range sheetRels.Relationship {
			if rel.Type == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" {
				drawingRID = rel.ID
				drawingFile = resolveTarget("xl/worksheets", rel.Target)
				break
			}
		}
	}
	if drawingFile == "" {
		// 新建 drawing
		drawingNum := nextFreeNumber(fileMap, "xl/drawings/drawing", ".xml")
		drawingFile = fmt.Sprintf("xl/drawings/drawing%d.xml", drawingNum)
		// 分配 worksheet 关系 rId
		wbRels := &Relationships{}
		if err := getXMLFromMap(fileMap, "xl/_rels/workbook.xml.rels", wbRels); err == nil {
			drawingRID = fmt.Sprintf("rId%d", getMaxRId(wbRels)+1)
		} else {
			drawingRID = "rId1"
		}
		if sheetRels == nil {
			sheetRels = &Relationships{XMLName: xmlNameRelationships()}
		}
		sheetRels.Relationship = append(sheetRels.Relationship, Relationship{
			ID:     drawingRID,
			Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
			Target: "../drawings/" + path.Base(drawingFile),
		})
		fileMap[drawingFile] = []byte(blankDrawing())
		// worksheet 加 <drawing r:id>
		ws := string(fileMap[file])
		ws = ensureWorksheetNamespaces(ws)
		ws = insertDrawingRef(ws, drawingRID)
		fileMap[file] = []byte(ws)
		// worksheet rels 回写
		fileMap[sheetRelsFile] = serializeRelationships(sheetRels)
		// Content_Types 追加 drawing Override
		fileMap["[Content_Types].xml"] = insertOverrideInContentTypes(
			fileMap["[Content_Types].xml"], "/"+drawingFile,
			"application/vnd.openxmlformats-officedocument.drawing+xml")
	} else {
		// 已有 drawing：确保 worksheet 含 <drawing r:id> 引用
		ws := string(fileMap[file])
		if !strings.Contains(ws, `<drawing`) {
			ws = ensureWorksheetNamespaces(ws)
			ws = insertDrawingRef(ws, drawingRID)
			fileMap[file] = []byte(ws)
		}
	}

	// 3) drawing rels：加 image 关系
	drawingRelsFile := "xl/drawings/_rels/" + path.Base(drawingFile) + ".rels"
	drawingRels := loadRels(fileMap, drawingRelsFile)
	if drawingRels == nil {
		drawingRels = &Relationships{XMLName: xmlNameRelationships()}
	}
	drawingRelIDForMedia = fmt.Sprintf("rId%d", getMaxRId(drawingRels)+1)
	mediaRel, _ := zipPathRel("xl/drawings", mediaPath)
	drawingRels.Relationship = append(drawingRels.Relationship, Relationship{
		ID:     drawingRelIDForMedia,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
		Target: mediaRel,
	})
	fileMap[drawingRelsFile] = serializeRelationships(drawingRels)

	// 4) 向 drawing 追加一个 oneCellAnchor
	dw := string(fileMap[drawingFile])
	dw = appendOneCellAnchor(dw, col, row, colOffPx, rowOffPx, extCx, extCy, drawingRelIDForMedia,
		len(drawingRels.Relationship))
	fileMap[drawingFile] = []byte(dw)

	// 5) 该 worksheet 可能还没有 drawing rels 文件（上面新建分支已建；若是已有 drawing 但
	//    worksheet rels 已存在则无需再建）。最终统一写回 worksheet rels（已有分支也要确保
	//    新建 drawing 的情况被序列化——上面已处理）。
	if sheetRels != nil && drawingFile != "" && !fileExistsInMap(fileMap, sheetRelsFile) {
		fileMap[sheetRelsFile] = serializeRelationships(sheetRels)
	}

	return writeMapToZip(filename, fileMap)
}

// AddCellPicture 在 sheetRef 工作表的 cell（如 "A1"）插入 WPS「图片嵌入单元格」式图片。
// 图片随单元格移动与缩放（x14:picture），同时保留标准 drawingML 回退（mc:Fallback），
// 保证 Excel/WPS 均可识别。不破坏工作表其他内容。
func AddCellPicture(filename, sheetRef, cell, picPath string) error {
	data, err := os.ReadFile(picPath)
	if err != nil {
		return fmt.Errorf("读取图片失败: %w", err)
	}
	return addCellPictureBytes(filename, sheetRef, cell, data)
}

// AddCellPictureFromBytes 同 AddCellPicture，但直接接收图片字节。
func AddCellPictureFromBytes(filename, sheetRef, cell string, data []byte) error {
	return addCellPictureBytes(filename, sheetRef, cell, data)
}

func addCellPictureBytes(filename, sheetRef, cell string, data []byte) error {
	col, row, err := parseCellRef(cell)
	if err != nil {
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
	file, _, _, _, err := locateSheetInMap(fileMap, wb, sheetRef)
	if err != nil {
		return err
	}

	// 1) 媒体
	mediaExt := imageExtFromData(data)
	mediaPath := generateUniqueName(fileMap, "xl/media", "image", mediaExt)
	fileMap[mediaPath] = data

	// 2) worksheet rels：加 image 关系
	sheetRelsFile := "xl/worksheets/_rels/" + path.Base(file) + ".rels"
	sheetRels := loadRels(fileMap, sheetRelsFile)
	if sheetRels == nil {
		sheetRels = &Relationships{XMLName: xmlNameRelationships()}
	}
	imgRID := fmt.Sprintf("rId%d", getMaxRId(sheetRels)+1)
	mediaRel, _ := zipPathRel("xl/worksheets", mediaPath)
	sheetRels.Relationship = append(sheetRels.Relationship, Relationship{
		ID:     imgRID,
		Type:   "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
		Target: mediaRel,
	})
	fileMap[sheetRelsFile] = serializeRelationships(sheetRels)

	// 3) worksheet 插入 mc:AlternateContent 块（在 sheetData 之后、drawing 之前），
	//    并确保命名空间（mc/x14/a/xdr）。
	ws := string(fileMap[file])
	alt := buildCellPictureAltContent(col, row, imgRID)
	ws = ensureWorksheetNamespaces(ws)
	ws = ensureX14Namespace(ws)
	ws = insertCellPictureBlock(ws, alt)
	fileMap[file] = []byte(ws)

	return writeMapToZip(filename, fileMap)
}

// ---------- drawing / picture XML 构造 ----------

func blankDrawing() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<wsDr xmlns="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"` +
		` xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"` +
		` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"></wsDr>`
}

// appendOneCellAnchor 向 drawing XML 追加一个 oneCellAnchor（单单元格锚点）。
func appendOneCellAnchor(dw string, col, row, colOffPx, rowOffPx, extCx, extCy int, blipRID string, picID int) string {
	fromColOff := colOffPx * EMUPerPixel
	fromRowOff := rowOffPx * EMUPerPixel
	anchor := fmt.Sprintf(
		`<oneCellAnchor><from><col>%d</col><colOff>%d</colOff><row>%d</row><rowOff>%d</rowOff></from>`+
			`<ext cx="%d" cy="%d" />`+
			`<pic><nvPicPr><cNvPr id="%d" name="Picture %d" /><cNvPicPr /></nvPicPr>`+
			`<blipFill><a:blip r:embed="%s" /><a:stretch><a:fillRect /></a:stretch></blipFill>`+
			`<spPr><a:prstGeom prst="rect" /><a:xfrm><a:off x="0" y="0" /><a:ext cx="%d" cy="%d" /></a:xfrm></spPr>`+
			`</pic><clientData /></oneCellAnchor>`,
		col, fromColOff, row, fromRowOff, extCx, extCy,
		picID, picID, blipRID, extCx, extCy,
	)
	closeIdx := strings.LastIndex(dw, "</wsDr>")
	if closeIdx == -1 {
		return dw + anchor
	}
	return dw[:closeIdx] + anchor + dw[closeIdx:]
}

// buildCellPictureAltContent 构造 WPS 单元格内嵌图片的 mc:AlternateContent 块。
func buildCellPictureAltContent(col, row int, rID string) string {
	return `<mc:AlternateContent>` +
		`<mc:Choice Requires="x14">` +
		`<x14:picture>` +
		`<x14:pic>` +
		`<x14:blip r:embed="` + rID + `" />` +
		`<x14:bodyPr />` +
		`<x14:clientData fPrintsWithSheet="1" />` +
		`</x14:pic>` +
		`</x14:picture>` +
		`</mc:Choice>` +
		`<mc:Fallback>` +
		`<xdr:pic>` +
		`<xdr:nvPicPr><xdr:cNvPr id="1" name="Picture" descr="Cell Picture" /><xdr:cNvPicPr /></xdr:nvPicPr>` +
		`<xdr:blipFill><a:blip r:embed="` + rID + `" /><a:stretch><a:fillRect /></a:stretch></xdr:blipFill>` +
		`<xdr:spPr><a:prstGeom prst="rect" /><a:xfrm><a:off x="0" y="0" /><a:ext cx="0" cy="0" /></a:xfrm></xdr:spPr>` +
		`</xdr:pic>` +
		`</mc:Fallback>` +
		`</mc:AlternateContent>`
}

// insertDrawingRef 在 worksheet 中加入 <drawing r:id="rid"/>（放在 sheetData 之后）。
func insertDrawingRef(ws, rid string) string {
	tag := `<drawing r:id="` + rid + `"/>`
	// 放在 sheetData 之后；若无 sheetData 则放在 worksheet 根末尾前
	if idx := strings.LastIndex(ws, "</sheetData>"); idx != -1 {
		return ws[:idx+len("</sheetData>")] + tag + ws[idx+len("</sheetData>"):]
	}
	if idx := strings.LastIndex(ws, "</worksheet>"); idx != -1 {
		return ws[:idx] + tag + ws[idx:]
	}
	return ws + tag
}

// insertCellPictureBlock 把 mc:AlternateContent 块插入 worksheet：sheetData 之后、drawing 之前。
func insertCellPictureBlock(ws, block string) string {
	// 优先放在 </sheetData> 之后
	if idx := strings.LastIndex(ws, "</sheetData>"); idx != -1 {
		insertAt := idx + len("</sheetData>")
		// 若其后紧跟 <drawing>，则插入到 drawing 之前
		rest := ws[insertAt:]
		if d := strings.Index(rest, "<drawing"); d != -1 {
			return ws[:insertAt] + block + rest[:d] + rest[d:]
		}
		return ws[:insertAt] + block + rest
	}
	if idx := strings.LastIndex(ws, "</worksheet>"); idx != -1 {
		return ws[:idx] + block + ws[idx:]
	}
	return ws + block
}

// ensureX14Namespace 确保 worksheet 根声明 x14 命名空间前缀。
func ensureX14Namespace(ws string) string {
	if strings.Contains(ws, `xmlns:x14="`) {
		return ws
	}
	start := strings.Index(ws, "<worksheet")
	if start == -1 {
		return ws
	}
	end := strings.Index(ws[start:], ">")
	if end == -1 {
		return ws
	}
	end += start
	root := ws[:end+1]
	root = strings.Replace(root, "<worksheet", `<worksheet xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"`, 1)
	return root + ws[end+1:]
}

// ---------- 工具 ----------

// parseCellRef 解析 "A1" 形式的单元格引用，返回 0 基列号与行号。
func parseCellRef(ref string) (col, row int, err error) {
	re := regexp.MustCompile(`^([A-Za-z]+)(\d+)$`)
	m := re.FindStringSubmatch(strings.TrimSpace(ref))
	if m == nil {
		return 0, 0, fmt.Errorf("无效的单元格引用: %q", ref)
	}
	colStr, rowStr := m[1], m[2]
	col = 0
	for _, ch := range strings.ToUpper(colStr) {
		col = col*26 + int(ch-'A'+1)
	}
	col-- // 转 0 基
	rowNum, e := strconv.Atoi(rowStr)
	if e != nil {
		return 0, 0, e
	}
	row = rowNum - 1 // 转 0 基
	return col, row, nil
}

// imageExtFromData 根据图片字节头判断扩展名（.png / .jpg）。
func imageExtFromData(data []byte) string {
	if len(data) >= 3 && data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF {
		return ".jpg"
	}
	if len(data) >= 8 && string(data[1:4]) == "PNG" {
		return ".png"
	}
	// 默认 png
	return ".png"
}

// xmlNameRelationships 返回 Relationships 的 xml.Name（用于构造空关系集合）。
func xmlNameRelationships() xml.Name { return xml.Name{Local: "Relationships"} }

// serializeRelationships 序列化 Relationships 为合法 rels 文件字节（带 XML 头与默认命名空间）。
func serializeRelationships(rels *Relationships) []byte {
	data, err := xml.MarshalIndent(rels, "", "  ")
	if err != nil {
		return nil
	}
	out := append([]byte(xml.Header), data...)
	out = bytes.Replace(out, []byte("<Relationships>"),
		[]byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`), 1)
	return out
}
