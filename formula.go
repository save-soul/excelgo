package excelgo

// formula.go 提供公式引用的自动平移能力。
//
// 背景：此前行列插入/删除时只平移单元格位置引用（<c r="...">），而公式 <f> 内部的
// A1 引用字符串保持原样（与旧版 excelize 行为一致，由调用方负责修正）。本文件补齐
// 这一能力——在插入/删除行列时，自动解析公式文本、按相同 pivot/delta 平移其中的单元格
// 引用，保证公式结果随布局变化而正确更新。
//
// 设计要点：
//   - 采用「手写 tokenizer + 引用重写」而非正则（Go 的 RE2 不支持 lookbehind，且正则
//     容易把函数名误判为单元格引用）。tokenizer 逐字符扫描，能正确区分：
//       · 字符串字面量 "..."（含 "" 转义）
//       · 引号包裹的工作表名 'My Sheet'! 与未引号表名 Sheet1!
//       · 函数名（后跟 '('，如 SUM/LOG10）→ 不平移
//       · 单元格引用（含 $ 绝对引用、跨表、范围 A1:B2）
//   - 绝对引用（$A$1 / $A1 / A$1）的对应维度不平移。
//   - 引用落入被删除的行列区间时，整体改写为 #REF!（与 Excel 行为一致）。
//   - 不支持 R1C1 样式引用（检测到则原样保留，避免破坏）。

import (
	"regexp"
	"strconv"
	"strings"
)

// isFormulaIdentChar 判断字符是否可组成标识符（表名/函数名/引用的一部分）。
func isFormulaIdentChar(c byte) bool {
	return c == '_' || c == '$' ||
		(c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
		(c >= '0' && c <= '9') || c == '.'
}

// isFormulaSpace 判断是否为公式中的空白。
func isFormulaSpace(c byte) bool {
	return c == ' ' || c == '\t' || c == '\n' || c == '\r'
}

// cellRefRe 匹配单个单元格引用（可带 $ 绝对标记），整串匹配。
var cellRefRe = regexp.MustCompile(`^\$?[A-Za-z]+\$?\d+$`)

// isCellRef 判断 ident 是否为合法单元格引用（非函数名/非定义名）。
func isCellRef(ident string) bool {
	return cellRefRe.MatchString(ident)
}

// shiftFormulaText 平移公式文本中的单元格引用。
//   - pivotRow / pivotCol：1 基插入/删除基准（行列操作各自的 pivot）。
//   - deltaRow / deltaCol：引用平移量（插入为正、删除为负）。
//   - delRows / delCols：被删除的 [起始,结束] 区间（1 基，闭区间）；为 nil 表示无删除。
//
// 返回改写后的公式文本；R1C1 样式直接原样返回。
func shiftFormulaText(f string, pivotRow, pivotCol, deltaRow, deltaCol int, delRows, delCols *[2]int) string {
	if isR1C1Formula(f) {
		return f
	}
	var b strings.Builder
	i := 0
	n := len(f)
	for i < n {
		c := f[i]
		switch {
		case c == '"':
			// 字符串字面量，原样保留（含 "" 转义）
			end := i + 1
			for end < n {
				if f[end] == '"' {
					if end+1 < n && f[end+1] == '"' {
						end += 2
						continue
					}
					end++
					break
				}
				end++
			}
			b.WriteString(f[i:end])
			i = end
		case c == '\'':
			// 引号包裹的工作表名 'My Sheet'!，原样保留（含 '' 转义）
			end := i + 1
			for end < n {
				if f[end] == '\'' {
					if end+1 < n && f[end+1] == '\'' {
						end += 2
						continue
					}
					end++
					break
				}
				end++
			}
			b.WriteString(f[i:end])
			i = end
		case isFormulaIdentChar(c):
			// 读取标识符（表名/函数名/引用）
			start := i
			for i < n && isFormulaIdentChar(f[i]) {
				i++
			}
			ident := f[start:i]
			// 检查后续是否紧跟 '!'（工作表前缀）
			j := i
			for j < n && isFormulaSpace(f[j]) {
				j++
			}
			if j < n && f[j] == '!' {
				b.WriteString(ident)
				b.WriteByte('!')
				i = j + 1
				continue
			}
			// 函数名（后跟 '('）不平移
			if i < n && f[i] == '(' {
				b.WriteString(ident)
				continue
			}
			// 单元格引用？
			if isCellRef(ident) {
				newRef := shiftCellRefToken(ident, pivotRow, pivotCol, deltaRow, deltaCol, delRows, delCols)
				// 是否为范围 A1:B2？
				k := i
				for k < n && isFormulaSpace(f[k]) {
					k++
				}
				if k < n && f[k] == ':' {
					// 范围：消耗 ':' 并尝试读取第二个引用
					k2 := k + 1
					for k2 < n && isFormulaSpace(f[k2]) {
						k2++
					}
					seg2Start := k2
					for k2 < n && isFormulaIdentChar(f[k2]) {
						k2++
					}
					ident2 := f[seg2Start:k2]
					if isCellRef(ident2) {
						newRef2 := shiftCellRefToken(ident2, pivotRow, pivotCol, deltaRow, deltaCol, delRows, delCols)
						if newRef == "#REF!" || newRef2 == "#REF!" {
							b.WriteString("#REF!")
						} else {
							b.WriteString(newRef)
							b.WriteByte(':')
							b.WriteString(newRef2)
						}
						i = k2
						continue
					}
					// ':' 后不是引用（极少见），原样输出 ':' 让循环继续处理
					b.WriteString(newRef)
					b.WriteByte(':')
					i = k2
					continue
				}
				b.WriteString(newRef)
				continue
			}
			// 其余（函数名无括号、定义名等）原样保留
			b.WriteString(ident)
		default:
			b.WriteByte(c)
			i++
		}
	}
	return b.String()
}

// isR1C1Formula 粗略判断公式是否使用 R1C1 样式引用（含 R1C1 / R[1]C[1] / RC 等），
// 命中则跳过平移（不在本实现支持范围）。
func isR1C1Formula(f string) bool {
	// 形如 R1C1、R[1]C[2]、RC[-1] 等：R 后接 [ 或数字，紧接 C。
	re := regexp.MustCompile(`(?i)R(\[?\d*\]?)C(\[?\d*\]?)`)
	return re.MatchString(f)
}

// shiftCellRefToken 平移单个引用 token（如 A1 / $A$1 / $A1 / A$1）。
// 返回改写后的 token；若引用落入被删除区间则返回 "#REF!"。
func shiftCellRefToken(ident string, pivotRow, pivotCol, deltaRow, deltaCol int, delRows, delCols *[2]int) string {
	// 拆分绝对标记与列字母、行号
	absCol := false
	absRow := false
	s := ident
	if strings.HasPrefix(s, "$") {
		absCol = true
		s = s[1:]
	}
	// 列字母部分（直到数字或行绝对标记）
	ci := 0
	for ci < len(s) && s[ci] >= 'A' && s[ci] <= 'Z' {
		ci++
	}
	colLetters := s[:ci]
	rest := s[ci:]
	if strings.HasPrefix(rest, "$") {
		absRow = true
		rest = rest[1:]
	}
	rowStr := rest
	col, err1 := colLettersToNum(colLetters)
	row, err2 := strconv.Atoi(rowStr)
	if err1 != nil || err2 != nil {
		return ident // 解析失败则原样返回（保持健壮）
	}

	refValid := true
	if !absRow {
		if delRows != nil && row >= delRows[0] && row <= delRows[1] {
			refValid = false
		} else if row >= pivotRow {
			row += deltaRow
			if row < 1 {
				refValid = false
			}
		}
	}
	if !absCol {
		if delCols != nil && col >= delCols[0] && col <= delCols[1] {
			refValid = false
		} else if col >= pivotCol {
			col += deltaCol
			if col < 1 {
				refValid = false
			}
		}
	}
	if !refValid {
		return "#REF!"
	}

	// 重建（保留绝对标记）
	var b strings.Builder
	if absCol {
		b.WriteByte('$')
	}
	b.WriteString(colNumToLetters(col - 1))
	if absRow {
		b.WriteByte('$')
	}
	b.WriteString(strconv.Itoa(row))
	return b.String()
}

// shiftFormulaInCell 平移单个 <c> 元素内 <f> 公式的引用，以及 <f> 的 ref 属性（共享公式范围）。
// 返回改写后的 <c> 文本。公式内容改写在内存字符串完成，不触碰其他任何属性/子元素。
func shiftFormulaInCell(c string, pivotRow, pivotCol, deltaRow, deltaCol int, delRows, delCols *[2]int) string {
	// 1) 改写 <f> 文本（<f ...>内容</f> 或自闭合 <f .../>）
	fRe := regexp.MustCompile(`(?s)(<f\b[^>]*>)(.*?)(</f>)`)
	c = fRe.ReplaceAllStringFunc(c, func(m string) string {
		sub := fRe.FindStringSubmatch(m)
		open, inner, close := sub[1], sub[2], sub[3]
		newInner := shiftFormulaText(inner, pivotRow, pivotCol, deltaRow, deltaCol, delRows, delCols)
		return open + newInner + close
	})
	// 自闭合 <f .../>
	fSelfRe := regexp.MustCompile(`(?s)<f\b[^>]*/>`)
	c = fSelfRe.ReplaceAllStringFunc(c, func(m string) string {
		// 自闭合通常无公式文本（缓存结果型），但可能含 ref 属性
		return shiftFAttrs(m, pivotRow, pivotCol, deltaRow, deltaCol, delRows, delCols)
	})
	// 2) 改写 <f> 的 ref 属性（共享公式范围，如 ref="A1:B10"）
	c = shiftFAttrs(c, pivotRow, pivotCol, deltaRow, deltaCol, delRows, delCols)
	return c
}

// shiftFAttrs 平移 <f ...> 标签中的 ref 属性值（共享公式范围）。
func shiftFAttrs(tag string, pivotRow, pivotCol, deltaRow, deltaCol int, delRows, delCols *[2]int) string {
	refRe := regexp.MustCompile(`(\bref=")([^"]*)(")`)
	return refRe.ReplaceAllStringFunc(tag, func(m string) string {
		sub := refRe.FindStringSubmatch(m)
		newRange := shiftRangeStr(sub[2], pivotRow, pivotCol, deltaRow, deltaCol, delRows, delCols)
		return sub[1] + newRange + sub[3]
	})
}

// shiftRangeStr 平移形如 "A1:B2" 的范围字符串（用于 <f ref> 等场景）。
func shiftRangeStr(rng string, pivotRow, pivotCol, deltaRow, deltaCol int, delRows, delCols *[2]int) string {
	parts := strings.Split(rng, ":")
	for i := range parts {
		if isCellRef(parts[i]) {
			parts[i] = shiftCellRefToken(parts[i], pivotRow, pivotCol, deltaRow, deltaCol, delRows, delCols)
		}
	}
	return strings.Join(parts, ":")
}
