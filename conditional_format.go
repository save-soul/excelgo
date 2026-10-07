package excelgo

// 条件格式：多规则与多规则类型的完整实现。
//
// 原实现只支持"单条 expression/cellIs 规则"，而 OOXML 的 cfRule 有 18 种类型，
// 且每种有不同的子元素结构（colorScale 要 cfvo+color、dataBar 要 cfvo、iconSet
// 要 iconSet+cfvo、top10 要 rank+percent…）。本文件补齐这些类型。
//
// 设计要点：
//  1. **规则与样式分离** —— 规则描述"何时命中"（type/operator/formula/色阶），
//     样式（dxf）描述"命中后长什么样"。多数规则类型（colorScale/dataBar/iconSet）
//     不使用 dxf，颜色直接内联在规则里。
//  2. **多个规则可挂同一区域** —— OOXML 里一个 <conditionalFormatting> 可含多个
//     <cfRule>，按 priority 依次求值（数字小的优先）。
//  3. **类型相关的字段集中在一个结构里**，避免为 18 种类型设计 18 个结构体。

import (
	"fmt"
	"strconv"
	"strings"
)

// ConditionalFormatOperator 是比较运算符（OOXML ST_Operator）。
type ConditionalFormatOperator string

// 常见比较运算符。取值为 OOXML 规范名。
const (
	OpLessThan    ConditionalFormatOperator = "lessThan"
	OpLessEqual   ConditionalFormatOperator = "lessThanOrEqual"
	OpEqual       ConditionalFormatOperator = "equal"
	OpNotEqual    ConditionalFormatOperator = "notEqual"
	OpGreaterEq   ConditionalFormatOperator = "greaterThanOrEqual"
	OpGreaterThan ConditionalFormatOperator = "greaterThan"
	OpBetween     ConditionalFormatOperator = "between"
	OpNotBetween  ConditionalFormatOperator = "notBetween"
)

// ColorScalePoint 是色阶的一个刻度点。
type ColorScalePoint struct {
	// Type 刻度类型："min" / "max" / "percentile" / "percent" / "num"
	Type string
	// Value 刻度值。Type 为 min/max 时可为空；percentile 需给 0-100。
	Value string
	// Color 颜色 ARGB，如 "FFFF0000"
	Color string
}

// ConditionalFormatRule 是一条完整的条件格式规则。
//
// 按 Type 组合使用：
//   - expression:        需 Formula
//   - cellIs:            需 Operator + Formula（between 用 Formula2）
//   - top10:             Rank + Percent + Bottom
//   - aboveAverage:      可选 AboveAverage（默认 true）与 EqualAverage
//   - containsText 等:    需 Text + Operator
//   - colorScale:        需 ColorScale（2 或 3 个点）
//   - dataBar:           需 Color，可选 MinType/MaxType
//   - iconSet:           需 IconSet
//   - duplicateValues / uniqueValues: 无额外字段
type ConditionalFormatRule struct {
	// Type 规则类型（OOXML CT_CfRule/@type）
	Type string
	// Operator 比较运算符（cellIs / containsText 等用）
	Operator string
	// Formula 条件公式。cellIs 的 between 需两条，用 Formula2。
	Formula  string
	Formula2 string
	// Text 文本类规则用（containsText / beginsWith / endsWith）
	Text string
	// StopIfTrue 命中后停止后续规则求值
	StopIfTrue bool

	// --- top10 ---
	Rank    int
	Percent bool
	Bottom  bool

	// --- aboveAverage ---
	AboveAverage *bool
	EqualAverage *bool

	// --- colorScale / dataBar / iconSet ---
	ColorScale   []ColorScalePoint
	Color        string
	MinType      string
	MaxType      string
	ShowValue    *bool
	IconSet      string
	ReverseIcons bool

	// --- dxf 样式（命中后应用）---
	// Style 为 nil 且规则类型不需要 dxf 时不生成 dxfId。
	Style *Style

	// Priority 优先级（1 起，数字小者优先）。为 0 时自动按添加顺序分配。
	Priority int
}

// validCFOps 是 OOXML ST_Operator 的合法取值。
var validCFOps = map[string]bool{
	"greaterThan": true, "lessThan": true, "greaterThanOrEqual": true,
	"lessThanOrEqual": true, "equal": true, "notEqual": true,
	"between": true, "notBetween": true,
	"containsText": true, "notContainsText": true, "beginsWith": true,
	"endsWith": true,
}

// validCFScaleTypes 是色阶/数据条刻度类型（OOXML ST_CfvoType）。
var validCFScaleTypes = map[string]bool{
	"min": true, "max": true, "num": true, "percent": true, "percentile": true,
	"formula": true,
}

// validIconSets 是 OOXML ST_IconSetType 的合法取值。
var validIconSets = map[string]bool{
	"3Arrows": true, "3ArrowsGray": true, "3Flags": true, "3TrafficLights1": true,
	"3TrafficLights2": true, "3Signs": true, "3Symbols": true, "3Symbols2": true,
	"4Arrows": true, "4ArrowsGray": true, "4RedToBlack": true,
	"4Rating": true, "4TrafficLights": true, "5Arrows": true, "5ArrowsGray": true,
	"5Rating": true, "5Quarters": true,
}

// validateCFRule 校验一条规则的类型特定字段。
func validateCFRule(r ConditionalFormatRule) error {
	if !validCFTypes[r.Type] {
		return fmt.Errorf("无效的条件格式类型 %q，合法的 OOXML 取值如："+
			"expression / cellIs / colorScale / dataBar / iconSet / top10 / "+
			"aboveAverage / duplicateValues / uniqueValues / containsText / "+
			"beginsWith / endsWith", r.Type)
	}
	if r.Operator != "" && !validCFOps[r.Operator] {
		return fmt.Errorf("无效的比较运算符 %q，合法的 OOXML 取值如："+
			"lessThan / lessThanOrEqual / equal / notEqual / greaterThanOrEqual / "+
			"greaterThan / between / notBetween", r.Operator)
	}
	switch r.Type {
	case "expression":
		if r.Formula == "" {
			return fmt.Errorf("expression 规则需要 Formula（条件表达式）")
		}
	case "cellIs":
		if r.Operator == "" {
			return fmt.Errorf("cellIs 规则需要 Operator")
		}
		if r.Formula == "" {
			return fmt.Errorf("cellIs 规则需要 Formula")
		}
		if (r.Operator == "between" || r.Operator == "notBetween") && r.Formula2 == "" {
			return fmt.Errorf("cellIs 的 %s 需要两条 Formula（用 Formula2 给第二条）", r.Operator)
		}
	case "containsText", "beginsWith", "endsWith":
		if r.Text == "" {
			return fmt.Errorf("%s 规则需要 Text", r.Type)
		}
		if r.Formula == "" {
			// 文本类规则需一条公式给出比较目标
			return fmt.Errorf("%s 规则需要 Formula（比较目标文本）", r.Type)
		}
	case "colorScale":
		if n := len(r.ColorScale); n != 2 && n != 3 {
			return fmt.Errorf("colorScale 需要 2 或 3 个刻度点，实际 %d", n)
		}
		for i, p := range r.ColorScale {
			if !validCFScaleTypes[p.Type] {
				return fmt.Errorf("色阶第 %d 个刻度类型 %q 非法，应为 min/max/num/percent/percentile",
					i+1, p.Type)
			}
			if p.Color == "" {
				return fmt.Errorf("色阶第 %d 个刻度缺少 Color", i+1)
			}
		}
	case "dataBar":
		if r.Color == "" {
			return fmt.Errorf("dataBar 需要 Color（数据条颜色）")
		}
		for _, t := range []struct{ name, v string }{{"MinType", r.MinType}, {"MaxType", r.MaxType}} {
			if t.v != "" && !validCFScaleTypes[t.v] {
				return fmt.Errorf("dataBar 的 %s=%q 非法，应为 min/max/num/percent/percentile",
					t.name, t.v)
			}
		}
	case "iconSet":
		if r.IconSet == "" {
			return fmt.Errorf("iconSet 需要 IconSet（如 3TrafficLights1）")
		}
		if !validIconSets[r.IconSet] {
			return fmt.Errorf("无效的图标集 %q，合法的如：3Arrows / 3TrafficLights1 / "+
				"3Signs / 4Rating / 5Quarters", r.IconSet)
		}
	}
	return nil
}

// buildCFRule 把规则序列化为 <cfRule> XML。
// dxfID 为命中后样式在 dxfs 里的索引；-1 表示不需要 dxf。
func buildCFRule(r ConditionalFormatRule, priority, dxfID int) (string, error) {
	if err := validateCFRule(r); err != nil {
		return "", err
	}
	var attrs []string
	attrs = append(attrs, `type="`+safeAttr(r.Type)+`"`)
	if r.Operator != "" {
		attrs = append(attrs, `operator="`+safeAttr(r.Operator)+`"`)
	}
	// 文本类规则必须带 text 属性（Excel 靠它知道匹配什么；
	// 只给 formula 的话消费方读回的 text 为空，规则无法求值）
	// 文本类规则必须带 text 属性（Excel 靠它知道匹配什么；
	// 只给 formula 的话消费方读回的 text 为空，规则无法求值）
	if r.Text != "" {
		attrs = append(attrs, `text="`+safeAttr(r.Text)+`"`)
	}
	if dxfID >= 0 {
		attrs = append(attrs, `dxfId="`+strconv.Itoa(dxfID)+`"`)
	}
	attrs = append(attrs, `priority="`+strconv.Itoa(priority)+`"`)
	if r.StopIfTrue {
		attrs = append(attrs, `stopIfTrue="1"`)
	}
	// top10 专有
	if r.Type == "top10" {
		if r.Rank <= 0 {
			r.Rank = 10
		}
		attrs = append(attrs, `rank="`+strconv.Itoa(r.Rank)+`"`)
		if r.Percent {
			attrs = append(attrs, `percent="1"`)
		}
		if r.Bottom {
			attrs = append(attrs, `bottom="1"`)
		}
	}
	// aboveAverage 专有
	if r.Type == "aboveAverage" {
		if r.AboveAverage != nil && !*r.AboveAverage {
			attrs = append(attrs, `aboveAverage="0"`)
		}
		if r.EqualAverage != nil && *r.EqualAverage {
			attrs = append(attrs, `equalAverage="1"`)
		}
	}

	open := `<cfRule ` + strings.Join(attrs, " ") + `>`

	// 需要子元素的类型走分支
	switch r.Type {
	case "colorScale":
		var b strings.Builder
		b.WriteString(`<colorScale>`)
		for _, p := range r.ColorScale {
			b.WriteString(`<cfvo type="` + safeAttr(p.Type) + `"`)
			if p.Value != "" {
				b.WriteString(` val="` + safeAttr(p.Value) + `"`)
			}
			b.WriteString(`/>`)
		}
		for _, p := range r.ColorScale {
			b.WriteString(`<color rgb="` + safeAttr(p.Color) + `"/>`)
		}
		b.WriteString(`</colorScale>`)
		return open + b.String() + `</cfRule>`, nil

	case "dataBar":
		var b strings.Builder
		b.WriteString(`<dataBar`)
		if r.ShowValue != nil && !*r.ShowValue {
			b.WriteString(` showValue="0"`)
		}
		b.WriteString(`>`)
		minT := r.MinType
		if minT == "" {
			minT = "min"
		}
		b.WriteString(`<cfvo type="` + safeAttr(minT) + `"/>`)
		maxT := r.MaxType
		if maxT == "" {
			maxT = "max"
		}
		b.WriteString(`<cfvo type="` + safeAttr(maxT) + `"/>`)
		b.WriteString(`<color rgb="` + safeAttr(r.Color) + `"/>`)
		b.WriteString(`</dataBar>`)
		return open + b.String() + `</cfRule>`, nil

	case "iconSet":
		var b strings.Builder
		b.WriteString(`<iconSet`)
		if r.IconSet != "" {
			b.WriteString(` iconSet="` + safeAttr(r.IconSet) + `"`)
		}
		if r.ReverseIcons {
			b.WriteString(` reverse="1"`)
		}
		b.WriteString(`>`)
		// iconSet 的 cfvo 数量由图标集决定：3 个图标 -> 2 个 cfvo，依此类推
		n := iconSetCfvoCount(r.IconSet)
		for i := 0; i < n; i++ {
			b.WriteString(`<cfvo type="percent" val="` + strconv.Itoa(i*100/(n-1)) + `"/>`)
		}
		b.WriteString(`</iconSet>`)
		return open + b.String() + `</cfRule>`, nil
	}

	// 其余类型：一条或多条 <formula>
	//
	// **无公式的类型不得写空 <formula>** —— aboveAverage / top10 /
	// duplicateValues / uniqueValues 靠属性与区域求值，附一个空公式会让
	// 消费方读出 formula=[''] 甚至误判规则。空公式必须省略。
	hasFormula := r.Formula != "" || r.Formula2 != ""
	if !hasFormula {
		return strings.TrimSuffix(open, ">") + `/>`, nil
	}
	open += `>`
	open += `<formula>` + safeText(r.Formula) + `</formula>`
	if r.Formula2 != "" {
		open += `<formula>` + safeText(r.Formula2) + `</formula>`
	}
	open += `</cfRule>`
	return open, nil
}

// iconSetCfvoCount 返回图标集对应的 cfvo 数量（图标数 - 1）。
func iconSetCfvoCount(iconSet string) int {
	n := 3
	if len(iconSet) >= 1 {
		if c := iconSet[0]; c >= '3' && c <= '5' {
			n = int(c - '0')
		}
	}
	if n < 2 {
		n = 2
	}
	return n - 1
}

// ---------- 公开 API ----------

// SetConditionalFormatRules 在 ref 区域设置一组条件格式规则。
//
// rules 中每条按 Type 组合字段（见 ConditionalFormatRule 说明）。
// 规则按给定顺序追加到同一 <conditionalFormatting> 内，priority 从 1 递增
// （除非规则显式指定了 Priority）。
//
// 示意：
//
//	// 高于均值标红
//	above := true
//	_ = ws.SetConditionalFormatRules("B2:B100", []excelgo.ConditionalFormatRule{
//	    {Type: "aboveAverage", AboveAverage: &above,
//	        Style: &excelgo.Style{Font: &excelgo.FontStyle{Color: "FFFF0000"}}},
//	})
//
//	// 三色阶
//	_ = ws.SetConditionalFormatRules("C2:C100", []excelgo.ConditionalFormatRule{
//	    {Type: "colorScale", ColorScale: []excelgo.ColorScalePoint{
//	        {Type: "min", Color: "FF63BE7B"},
//	        {Type: "percentile", Value: "50", Color: "FFFFEB84"},
//	        {Type: "max", Color: "FFF8696B"},
//	    }},
//	})
func (s *WorkSheet) SetConditionalFormatRules(ref string, rules []ConditionalFormatRule) error {
	if len(rules) == 0 {
		return fmt.Errorf("条件格式规则列表为空")
	}
	// 区域引用校验：未校验的坐标会直接进 XML 属性
	r := newRangeRef(ref)
	if r == "" {
		return fmt.Errorf("无效的条件格式区域引用: %q", ref)
	}
	// 提前校验全部规则，避免写一半失败留下半成品
	for i, rule := range rules {
		if err := validateCFRule(rule); err != nil {
			return fmt.Errorf("第 %d 条规则无效: %w", i+1, err)
		}
	}
	fm := s.fm()
	var b strings.Builder
	b.WriteString(`<conditionalFormatting sqref="` + string(r) + `">`)
	for i, rule := range rules {
		priority := rule.Priority
		if priority <= 0 {
			priority = existingRuleCount(string(s.ws())) + i + 1
		}
		dxfID := -1
		if rule.Style != nil {
			id, err := addDxfToStyles(fm, *rule.Style)
			if err != nil {
				return err
			}
			dxfID = id
		}
		xml, err := buildCFRule(rule, priority, dxfID)
		if err != nil {
			return err
		}
		b.WriteString(xml)
	}
	b.WriteString(`</conditionalFormatting>`)
	ws := insertAfterSheetData(string(s.ws()), frag(b.String()))
	s.setWS(ws)
	return nil
}

// SetConditionalFormatRules 在 ref 区域设置一组条件格式规则（包级 API）。
func SetConditionalFormatRules(filename, sheetRef, ref string,
	rules []ConditionalFormatRule) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.SetConditionalFormatRules(ref, rules)
	})
}

// existingRuleCount 统计已有 cfRule 数量，用于自动分配 priority。
func existingRuleCount(ws string) int {
	return strings.Count(ws, "<cfRule")
}
