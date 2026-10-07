package excelgo

// formula_cache.go 实现公式缓存结果的清除（对应 excelize 的 UpdateLinkedValue）。

import (
	"fmt"
	"regexp"
	"strings"
)

// UpdateLinkedValue 清除工作簿中所有公式单元格的缓存结果（<v> 标签）。
//
// 背景：Excel 保存文件时会把公式的计算结果一并写入 <v>。若计算因子已变化
// （引用了外部工作簿、被本库改写了输入值、或经 CopySheet/Merge 搬到了新位置），
// 那份缓存就是**过期的**。Excel 2007/2010 打开这类文件时不会自动重算，
// 用户看到的仍是旧结果，只能手动 Ctrl+Alt+F9。
//
// 做法：把每个含 <f> 的单元格里的 <v> 去掉，只留 <f>：
//
//	清除前：<c r="B19"><f>SUM(Sheet2!D2,Sheet2!D11)</f><v>100</v></c>
//	清除后：<c r="B19"><f>SUM(Sheet2!D2,Sheet2!D11)</f></c>
//
// 代价（与 excelize 一致、也是本方法的设计前提）：文档内容发生了变化，
// Excel 重新打开时会重算公式，而它算出来的结果与文件里存的不一样，
// 于是关闭时会提示"是否保存工作簿"。
//
// 无副作用于非公式单元格：不含 <f> 的 <c> 一律不动。
//
// 注意：这是**内存操作**，需自行调用 Save / SaveAs 落盘（与本库其它 Book 方法一致）。
func (b *Book) UpdateLinkedValue() error {
	if b.fileMap == nil {
		return fmt.Errorf("工作簿未打开")
	}
	// 共享字符串缓存的指纹基于部件字节；本方法不改 SST，但保险起见作废，
	// 避免极端情况下缓存持有的是被间接改写的部件快照。
	b.sstCache = nil

	for name, data := range b.fileMap {
		if !isWorksheetPart(name) {
			continue
		}
		out, cleared := clearFormulaCache(string(data))
		if cleared > 0 {
			b.fileMap[name] = []byte(out)
		}
	}
	return nil
}

// isWorksheetPart 判断压缩包内路径是否为工作表部件：xl/worksheets/sheetN.xml。
func isWorksheetPart(name string) bool {
	if !strings.HasPrefix(name, "xl/worksheets/") || !strings.HasSuffix(name, ".xml") {
		return false
	}
	// 排除 _rels/ 子目录与其它嵌套部件
	return !strings.Contains(name, "/_rels/")
}

// 匹配「<c ...>…<f>…</f><v>…</v>…</c>」中的 <v>，仅当同一 <c> 内出现过 <f>。
//
// 逐单元格处理而非整表正则替换 <f>.*?<v>：后者会跨单元格贪婪匹配，
// 把紧随公式单元格之后的普通数值单元格的 <v> 一并吃掉 —— 那是数据不是缓存。
var (
	// 匹配单个 <c …>…</c> 元素（内部可能含 <f> 与 <v>）。
	formulaCellRe = regexp.MustCompile(`(?s)<c\b[^>]*>.*?</c>`)
	// 该 <c> 内的公式标签。
	cellFormulaRe = regexp.MustCompile(`(?s)<f\b[^>]*>.*?</f>|<f\b[^>]*/>`)
	// 该 <c> 内的缓存值标签（带前后空白，避免拼出畸形 XML）。
	cellValueRe = regexp.MustCompile(`(?s)<v\b[^>]*>.*?</v>|<v\b[^>]*/>`)
)

// clearFormulaCache 移除一份工作表 XML 里所有公式单元格的 <v>，返回处理后的 XML
// 与被清除的缓存个数。
func clearFormulaCache(ws string) (string, int) {
	cleared := 0
	out := formulaCellRe.ReplaceAllStringFunc(ws, func(c string) string {
		if !cellFormulaRe.MatchString(c) {
			return c // 非公式单元格：原样保留
		}
		stripped := cellValueRe.ReplaceAllString(c, "")
		if stripped == c {
			return c // 本就没有缓存值
		}
		cleared++
		return stripped
	})
	return out, cleared
}
