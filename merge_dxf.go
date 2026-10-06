package excelgo

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

// 跨工作簿搬运条件格式的**差分样式**（dxf）。
//
// 背景：条件格式规则通过 `dxfId` 引用 styles.xml 里 <dxfs> 的第 N 条。
// `dxfId` 是**工作簿级索引** —— 与 cellXf 一样，跨簿不能直接复用。
//
// 早期实现只搬 sheet XML（含 dxfId），不搬 dxfs，也不重映射索引：
// 合并一份带条件格式的工作簿后，目标簿 styles.xml 里没有对应 dxf，
// 而 sheet 里的 dxfId 仍指着某个下标 —— **悬空引用**。
// openpyxl 直接抛
//
//	IndexError: list index out of range
//   at differential.py: __getitem__  (self.styles[idx])
//
// Excel 则报"文件已损坏"。这个缺陷只在"合并带条件格式的跨簿文件"
// 时暴露，单簿测试永远测不到。

// dxfEntry 是一条差分样式的原始 XML。
type dxfEntry struct {
	xml  string
	body string // <dxf>...</dxf> 完整串
}

// parseDxfs 解析 styles.xml 里的 <dxfs>，返回按索引排列的条目。
func parseDxfs(stylesXML []byte) []dxfEntry {
	s := string(stylesXML)
	m := regexp.MustCompile(`(?s)<dxfs\b[^>]*>(.*?)</dxfs>`).FindStringSubmatch(s)
	if m == nil {
		return nil
	}
	var out []dxfEntry
	for _, d := range regexp.MustCompile(`(?s)<dxf>.*?</dxf>`).FindAllString(m[1], -1) {
		out = append(out, dxfEntry{xml: d, body: d})
	}
	return out
}

// dxfBodyHash 取 dxf 内容的归一化形式，用于跨簿去重。
func dxfBodyHash(d string) string {
	return normalizeWhitespace(strings.TrimSpace(d))
}

// mergeDxfsInto 把源 styles 的 dxfs 并入目标 styles，返回
// "源索引 -> 目标索引" 的映射。
//
// 同内容的目标已有条目则复用（省空间且避免同一格式出现多条），
// 否则追加新的。
func mergeDxfsInto(dstStyles []byte, srcStyles []byte) ([]byte, map[int]int) {
	src := parseDxfs(srcStyles)
	if len(src) == 0 {
		return dstStyles, nil
	}
	dst := parseDxfs(dstStyles)

	// 目标已有的内容哈希 -> 索引
	index := make(map[string]int, len(dst))
	for i, d := range dst {
		index[dxfBodyHash(d.body)] = i
	}

	mapping := make(map[int]int, len(src))
	added := make([]string, 0, len(src))
	for i, d := range src {
		h := dxfBodyHash(d.body)
		if j, ok := index[h]; ok {
			mapping[i] = j
			continue
		}
		mapping[i] = len(dst) + len(added)
		index[h] = mapping[i]
		added = append(added, d.body)
	}
	if len(added) == 0 {
		return dstStyles, mapping
	}
	return appendDxfs(dstStyles, dst, added), mapping
}

// appendDxfs 把新条目追加进目标 styles.xml 的 <dxfs>，并更新 count。
func appendDxfs(dstStyles []byte, existing []dxfEntry, added []string) []byte {
	all := make([]string, 0, len(existing)+len(added))
	for _, d := range existing {
		all = append(all, d.body)
	}
	all = append(all, added...)
	// 目标没有 <dxfs> 时插在 cellXfs 之后（OOXML 元素顺序：dxfs 紧随 cellXfs）。
	// 注意必须传**全部** all 而不只是 added —— 否则目标原有条目会被丢弃。
	return []byte(replaceOrInsertBlock(string(dstStyles), "dxfs", all,
		fmt.Sprintf(`<dxfs count="%d">`, len(all)), "</dxfs>",
		[]string{"cellStyles", "tableStyles", "/styleSheet"}))
}

// remapDxfIds 改写 sheet XML 里的 dxfId，使其指向合并后的新索引。
func remapDxfIds(sheetXML string, mapping map[int]int) string {
	if len(mapping) == 0 {
		return sheetXML
	}
	const prefix = `dxfId="`
	re := regexp.MustCompile(`dxfId="(\d+)"`)
	return re.ReplaceAllStringFunc(sheetXML, func(m string) string {
		n, err := strconv.Atoi(m[len(prefix) : len(m)-1])
		if err != nil {
			return m
		}
		if to, ok := mapping[n]; ok && to != n {
			return prefix + strconv.Itoa(to) + `"`
		}
		return m
	})
}

// mergeDxfsAndRemapSheet 是跨簿搬运条件格式的**完整一步**：
// 把源 styles 的 dxfs 并入目标 styles，并把 sheet XML 里的 dxfId 重映射。
//
// 抽成共用函数的原因（真实教训）：最初只在 Merge 的mergeOneSheet 里补了
// 这一步，CopySheetTo 的 copySheetAcrossMaps 走的是另一条路径，
// 于是同样悬空的 dxfId 在那条路上继续产生 —— openpyxl 抛
// IndexError: list index out of range。
//
// 凡是把工作表从 A 簿搬到 B 簿的地方，都必须走这一步。
//
// 返回更新后的 sheet XML；调用方负责在所有改写完成后一次性写回。
func mergeDxfsAndRemapSheet(dstMap, srcMap map[string][]byte, sheetXML string,
	styleXfRemap map[int]int) string {
	srcStyles, ok := srcMap["xl/styles.xml"]
	if !ok {
		return sheetXML
	}
	dstStyles, exists := dstMap["xl/styles.xml"]
	if !exists {
		// 目标还没有 styles.xml：无从"并入"，但源有 dxfs 时必须建一份，
		// 否则 dxfId 指向空表。
		dstStyles = []byte(defaultStylesXML())
		dstMap["[Content_Types].xml"] = insertOverrideInContentTypes(
			dstMap["[Content_Types].xml"], "/xl/styles.xml",
			"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
	}
	merged, mapping := mergeDxfsInto(dstStyles, srcStyles)
	// 命名样式（cellStyles）也在这同一张表上，必须**接着**上面已合并的
	// styles.xml 继续做，不能各自从头写 dstMap["xl/styles.xml"] ——
	// 那样后写的会覆盖先写的（曾导致 dxfs 被 cellStyles 覆盖，
	// openpyxl 又抛 IndexError）。
	merged = appendNamedStyles(merged, srcStyles, styleXfRemap)
	dstMap["xl/styles.xml"] = merged
	return remapDxfIds(sheetXML, mapping)
}

// appendNamedStyles 把源的命名样式并入**已合并的** styles 字节并返回新字节。
//
// 关键点：命名样式的 <cellStyle xfId="N"> 指向 **cellStyleXfs**（不是 cellXfs），
// 而 cellStyleXfs 里的 xf 又指向 fonts/fills/borders —— 是一整条间接链。
// 只搬 <cellStyles> 清单而不搬 <cellStyleXfs>，xfId 就会悬空：
// openpyxl 读命名样式时按 xfId 取 cellStyleXfs[N]，越界即抛
//
//	IndexError: list index out of range
//
// 故这里同时做三件事：
//  1. 把源用到的 cellStyleXfs 条目追加到目标，并把条目内部的
//     fontId/fillId/borderId/numFmtId 按目标簿的偏移平移（同样是工作簿级索引）；
//  2. 把 <cellStyle xfId> 重映射到新的下标；
//  3. 追加 <cellStyles> 清单（按名字去重）。
//
// 本函数是纯函数（收 styles 字节、返回新字节），不写 dstMap —— 与 dxf 合并
// 并存时，两个函数各自写 map 会互相覆盖（曾导致 dxfs 被覆盖后又抛 IndexError）。
func appendNamedStyles(dstStyles, srcStyles []byte, styleXfRemap map[int]int) []byte {
	srcNamed := extractCellStyles(string(srcStyles))
	if len(srcNamed) == 0 {
		return dstStyles
	}
	xml := string(dstStyles)
	// cellStyleXfs 本身已由 StylesMerger 合并 —— 只有它知道 font/fill/border/numFmt
	// 在目标簿里的正确下标。这里若再按"目标表长度"手工算偏移，会与 cellXfs 的
	// 合并**重复累加**，产出越界的 fontId（openpyxl: IndexError）。
	remap := styleXfRemap

	// --- 2) cellStyles：按名字去重后追加，xfId 重映射 ---
	dstNamed := extractCellStyles(xml)
	byName := map[string]bool{}
	for _, e := range dstNamed {
		byName[cellStyleName(e)] = true
	}
	merged := append([]string{}, dstNamed...)
	for _, e := range srcNamed {
		name := cellStyleName(e)
		if byName[name] {
			continue
		}
		byName[name] = true
		merged = append(merged, remapCellStyleXfID(e, remap))
	}
	block := `<cellStyles count="` + itoa(len(merged)) + `">` +
		strings.Join(merged, "") + `</cellStyles>`
	// cellStyles 必须排在 dxfs 之前（OOXML 元素顺序要求）
	if re := regexp.MustCompile(`(?s)<cellStyles\b[^>]*>.*?</cellStyles>`); re.MatchString(xml) {
		return []byte(re.ReplaceAllString(xml, block))
	}
	if i := strings.Index(xml, "<dxfs"); i != -1 {
		return []byte(xml[:i] + block + xml[i:])
	}
	if i := strings.Index(xml, "</styleSheet>"); i != -1 {
		return []byte(xml[:i] + block + xml[i:])
	}
	return []byte(xml)
}

func remapCellStyleXfID(entry string, remap map[int]int) string {
	return regexp.MustCompile(`xfId="(\d+)"`).ReplaceAllStringFunc(entry, func(m string) string {
		n, err := strconv.Atoi(m[len(`xfId="`):])
		if err != nil {
			return m
		}
		if to, ok := remap[n]; ok {
			return `xfId="` + itoa(to) + `"`
		}
		return m
	})
}

// replaceOrInsertBlock 替换或插入一个区块，保持 count 与子元素一致。
func replaceOrInsertBlock(xml string, block xmlTagName, items []string, open, close string,
	anchors []string) string {
	body := open + strings.Join(items, "") + close
	// block 是 schema 里固定的标签名（xmlTagName 认证类型），不含用户数据
	if re := regexp.MustCompile(`(?s)<` + string(block) +
		`\b[^>]*>.*?</` + string(block) + `>`); re.MatchString(xml) {
		return re.ReplaceAllString(xml, body)
	}
	// anchors 传的是**不含**尖括号的标签名（如 cellXfs），这里拼出尖括号。
	// 若调用方误传了带尖括号的值，会变成双尖括号而永远匹配不上——
	// 结果是整段替换静默失败，症状是下游出现悬空索引。
	for _, a := range anchors {
		tag := strings.TrimPrefix(a, "<")
		if i := strings.Index(xml, "<"+tag); i != -1 {
			return xml[:i] + body + xml[i:]
		}
	}
	return xml
}

// extractCellStyles 取出 <cellStyles> 里的各个 <cellStyle .../> 原文（顺序即清单）。
func extractCellStyles(stylesXML string) []string {
	m := regexp.MustCompile(`(?s)<cellStyles\b[^>]*>(.*?)</cellStyles>`).
		FindStringSubmatch(stylesXML)
	if m == nil {
		return nil
	}
	return regexp.MustCompile(`<cellStyle\b[^>]*/>`).FindAllString(m[1], -1)
}

// cellStyleName 取出 <cellStyle name="..."/> 里的名字。
func cellStyleName(entry string) string {
	i := strings.Index(entry, `name="`)
	if i == -1 {
		return ""
	}
	rest := entry[i+len(`name="`):]
	j := strings.Index(rest, `"`)
	if j == -1 {
		return ""
	}
	return rest[:j]
}
