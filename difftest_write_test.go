package excelgo

// 写侧差分场景：同一组操作分别用 excelgo 与 openpyxl 执行，比对语义快照。
//
// 覆盖策略：按"能力域"组织，每个域若干场景。域内尽量覆盖边界值
// （空串、超长、中文、emoji、负数、零、前导零等），
// 因为差异最常出现在边界而非典型值。

import (
	"fmt"
	"testing"
)

func TestDiffWriteSide(t *testing.T) {
	cases := []difftestCase{
		// ---------- 单元格值 ----------
		{
			Name: "字符串-常规",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellStr(f, "Sheet1", "A1", "hello"))
				must(t, SetCellStr(f, "Sheet1", "B1", ""))
			},
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"hello"},
			       {"op":"set_str","sheet":"Sheet1","cell":"B1","value":""}`,
			// 空串：excelgo 写入共享字符串表并能读回 ""；openpyxl 写 inlineStr 空元素，
			// 读回是 None —— 它连自己写的空串都读不回来（已实测 openpyxl 往返同样为 None）。
			// 这是 openpyxl 的往返损失，excelgo 的行为更符合"空串是有效值"的语义。
			Exempts: []string{".sheets.Sheet1.cells.B1"},
		},
		{
			Name: "字符串-中文与emoji",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellStr(f, "Sheet1", "A1", "中文测试"))
				must(t, SetCellStr(f, "Sheet1", "A2", "emoji 🚀🌏"))
				must(t, SetCellStr(f, "Sheet1", "A3", "带\"引号'单引号"))
				must(t, SetCellStr(f, "Sheet1", "A4", "换行\n第二行"))
				must(t, SetCellStr(f, "Sheet1", "A5", "制表\t符"))
			},
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"中文测试"},
			       {"op":"set_str","sheet":"Sheet1","cell":"A2","value":"emoji 🚀🌏"},
			       {"op":"set_str","sheet":"Sheet1","cell":"A3","value":"带\"引号'单引号"},
			       {"op":"set_str","sheet":"Sheet1","cell":"A4","value":"换行\n第二行"},
			       {"op":"set_str","sheet":"Sheet1","cell":"A5","value":"制表\t符"}`,
		},
		{
			Name: "字符串-XML敏感字符",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellStr(f, "Sheet1", "A1", `<b>&amp;</b>`))
				must(t, SetCellStr(f, "Sheet1", "A2", `a<b>c & d`))
				must(t, SetCellStr(f, "Sheet1", "A3", `]]>`))
			},
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"<b>&amp;</b>"},
			       {"op":"set_str","sheet":"Sheet1","cell":"A2","value":"a<b>c & d"},
			       {"op":"set_str","sheet":"Sheet1","cell":"A3","value":"]]>"}`,
		},
		{
			Name: "数值-边界",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellNumeric(f, "Sheet1", "A1", 0))
				must(t, SetCellNumeric(f, "Sheet1", "A2", -1))
				must(t, SetCellNumeric(f, "Sheet1", "A3", 3.14159))
				must(t, SetCellNumeric(f, "Sheet1", "A4", -0.5))
				must(t, SetCellNumeric(f, "Sheet1", "A5", 1e10))
				must(t, SetCellNumeric(f, "Sheet1", "A6", 0.1+0.2))
				must(t, SetCellInt(f, "Sheet1", "A7", -2147483648))
			},
			Ops: `{"op":"set_num","sheet":"Sheet1","cell":"A1","value":0},
			       {"op":"set_num","sheet":"Sheet1","cell":"A2","value":-1},
			       {"op":"set_num","sheet":"Sheet1","cell":"A3","value":3.14159},
			       {"op":"set_num","sheet":"Sheet1","cell":"A4","value":-0.5},
			       {"op":"set_num","sheet":"Sheet1","cell":"A5","value":1e10},
			       {"op":"set_num","sheet":"Sheet1","cell":"A6","value":0.30000000000000004},
			       {"op":"set_num","sheet":"Sheet1","cell":"A7","value":-2147483648}`,
		},
		{
			Name: "布尔值",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellBool(f, "Sheet1", "A1", true))
				must(t, SetCellBool(f, "Sheet1", "A2", false))
			},
			Ops: `{"op":"set_bool","sheet":"Sheet1","cell":"A1","value":true},
			       {"op":"set_bool","sheet":"Sheet1","cell":"A2","value":false}`,
		},
		{
			Name: "公式",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellFormula(f, "Sheet1", "A1", "SUM(B1:C1)", "10"))
				must(t, SetCellFormula(f, "Sheet1", "A2", "IF(A1>5,\"big\",\"small\")", "big"))
			},
			Ops: `{"op":"set_formula","sheet":"Sheet1","cell":"A1","formula":"SUM(B1:C1)"},
			       {"op":"set_formula","sheet":"Sheet1","cell":"A2","formula":"IF(A1>5,\"big\",\"small\")"}`,
			// openpyxl 不写缓存结果值（<v>），本库写 result。快照只看公式文本，
			// 但为稳妥起见声明该字段差异。
			Exempts: []string{".sheets.Sheet1.cells.A1", ".sheets.Sheet1.cells.A2"},
			SkipReason: "openpyxl 写公式时不写缓存结果值，openpyxl 读回会看到 None 而非结果字符串；" +
				"这是 openpyxl 的已知限制（data_only=False 下只给公式），非本库缺陷。",
		},

		// ---------- 区域写入 ----------
		{
			Name: "SetRange-混合类型",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetRange(f, "Sheet1", "A1", [][]interface{}{
					{"名称", "数量", "金额"},
					{"钢管", 12, 1500.5},
					{"扣件", 0, -20.25},
				}))
			},
			Ops: `{"op":"set_range","sheet":"Sheet1","values":[["名称","数量","金额"],
			       ["钢管",12,1500.5],["扣件",0,-20.25]]}`,
		},

		// ---------- 样式 ----------
		{
			Name: "样式-字体",
			ExcelGo: func(t *testing.T, f string) {
				_, err := SetCellStyle(f, "Sheet1", "A1", Style{
					Font: &FontStyle{Bold: true, Italic: true, Size: 14, Color: "FFFF0000"},
				})
				must(t, err)
			},
			// 传 8 位 ARGB（与 Go 侧字面量完全一致）。若这里传 6 位 RGB，
			// op.py 会补 FF 前缀，两侧就变成了不同的颜色值 —— 那是用例不对称，
			// 不是被测库的差异。
			Ops: `{"op":"set_style","sheet":"Sheet1","cell":"A1",
			       "style":{"bold":true,"italic":true,"size":14,"color":"FFFF0000"}}`,
		},
		{
			Name: "样式-填充与对齐",
			ExcelGo: func(t *testing.T, f string) {
				_, err := SetCellStyle(f, "Sheet1", "A1", Style{
					Fill:      &FillStyle{Color: "FF00FF00", PatternType: "solid"},
					Alignment: &AlignmentStyle{Horizontal: "center", Vertical: "center", WrapText: true},
				})
				must(t, err)
			},
			Ops: `{"op":"set_style","sheet":"Sheet1","cell":"A1",
			       "style":{"fill":"00FF00","halign":"center","valign":"center","wrap":true}}`,
		},
		{
			Name: "样式-数字格式",
			ExcelGo: func(t *testing.T, f string) {
				_, err := SetCellStyle(f, "Sheet1", "A1", Style{NumFmt: "0.00"})
				must(t, err)
				_, err = SetCellStyle(f, "Sheet1", "A2", Style{NumFmt: "yyyy-mm-dd"})
				must(t, err)
				_, err = SetCellStyle(f, "Sheet1", "A3", Style{NumFmt: "0.00%"})
				must(t, err)
			},
			Ops: `{"op":"set_style","sheet":"Sheet1","cell":"A1","style":{"numfmt":"0.00"}},
			       {"op":"set_style","sheet":"Sheet1","cell":"A2","style":{"numfmt":"yyyy-mm-dd"}},
			       {"op":"set_style","sheet":"Sheet1","cell":"A3","style":{"numfmt":"0.00%"}}`,
		},
		{
			Name: "样式-边框",
			ExcelGo: func(t *testing.T, f string) {
				_, err := SetCellStyle(f, "Sheet1", "A1", Style{
					Border: &BorderStyle{
						Left: BorderSide{Style: "thin", Color: "000000"},
						Top:  BorderSide{Style: "medium", Color: "FF0000"},
					},
				})
				must(t, err)
			},
			Ops: `{"op":"set_style","sheet":"Sheet1","cell":"A1",
			       "style":{"border":{"left":"thin","top":"medium"}}}`,
		},
		{
			Name: "样式-区域套用",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellStr(f, "Sheet1", "A1", "表头"))
				_, err := SetCellStyleRange(f, "Sheet1", "A1:C1", Style{
					Font: &FontStyle{Bold: true},
				})
				must(t, err)
			},
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"表头"},
			       {"op":"set_style","sheet":"Sheet1","cell":"A1","style":{"bold":true}},
			       {"op":"set_style","sheet":"Sheet1","cell":"B1","style":{"bold":true}},
			       {"op":"set_style","sheet":"Sheet1","cell":"C1","style":{"bold":true}}`,
		},

		// ---------- 行列 ----------
		{
			Name: "列宽行高",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetColWidth(f, "Sheet1", 2, 25.5))
				must(t, SetColWidthRange(f, "Sheet1", 4, 6, 12))
				must(t, SetRowHeight(f, "Sheet1", 3, 28))
			},
			Ops: `{"op":"col_width","sheet":"Sheet1","col":2,"width":25.5},
			       {"op":"col_width","sheet":"Sheet1","min":4,"max":6,"width":12},
			       {"op":"row_height","sheet":"Sheet1","row":3,"height":28}`,
		},
		{
			Name: "插行-首行",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellStr(f, "Sheet1", "A1", "原A1"))
				must(t, SetCellStr(f, "Sheet1", "A2", "原A2"))
				must(t, InsertRows(f, "Sheet1", 1, 1))
			},
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"原A1"},
			       {"op":"set_str","sheet":"Sheet1","cell":"A2","value":"原A2"},
			       {"op":"insert_rows","sheet":"Sheet1","row":1,"count":1}`,
		},
		{
			Name: "插行-中部并带公式",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellStr(f, "Sheet1", "A1", "a"))
				must(t, SetCellFormula(f, "Sheet1", "B1", "A1&\"!\"", "a!"))
				must(t, SetCellStr(f, "Sheet1", "A2", "b"))
				must(t, InsertRows(f, "Sheet1", 2, 2))
			},
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"a"},
			       {"op":"set_formula","sheet":"Sheet1","cell":"B1","formula":"A1&\"!\""},
			       {"op":"set_str","sheet":"Sheet1","cell":"A2","value":"b"},
			       {"op":"insert_rows","sheet":"Sheet1","row":2,"count":2}`,
			SkipReason: "openpyxl 的 insert_rows 不平移公式引用（Excel 会把 A1 变 A3），" +
				"本库平移。属实现语义差异，需专项对比而非等价对比。",
		},
		{
			Name: "插列",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellStr(f, "Sheet1", "A1", "a"))
				must(t, SetCellStr(f, "Sheet1", "B1", "b"))
				must(t, SetCellStr(f, "Sheet1", "C1", "c"))
				must(t, InsertCols(f, "Sheet1", 2, 1))
			},
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"a"},
			       {"op":"set_str","sheet":"Sheet1","cell":"B1","value":"b"},
			       {"op":"set_str","sheet":"Sheet1","cell":"C1","value":"c"},
			       {"op":"insert_cols","sheet":"Sheet1","col":2,"count":1}`,
		},
		{
			Name: "删除行列",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellStr(f, "Sheet1", "A1", "a"))
				must(t, SetCellStr(f, "Sheet1", "B1", "b"))
				must(t, SetCellStr(f, "Sheet1", "C1", "c"))
				must(t, RemoveCols(f, "Sheet1", 2, 1))
			},
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"a"},
			       {"op":"set_str","sheet":"Sheet1","cell":"B1","value":"b"},
			       {"op":"set_str","sheet":"Sheet1","cell":"C1","value":"c"},
			       {"op":"remove_cols","sheet":"Sheet1","col":2,"count":1}`,
		},

		// ---------- 布局 ----------
		{
			Name: "合并单元格",
			ExcelGo: func(t *testing.T, f string) {
				must(t, MergeCells(f, "Sheet1", "A1:C1"))
				must(t, MergeCells(f, "Sheet1", "A3:B4"))
				must(t, SetCellStr(f, "Sheet1", "A1", "跨列"))
			},
			Ops: `{"op":"merge","sheet":"Sheet1","ref":"A1:C1"},
			       {"op":"merge","sheet":"Sheet1","ref":"A3:B4"},
			       {"op":"set_str","sheet":"Sheet1","cell":"A1","value":"跨列"}`,
		},
		{
			Name: "取消合并",
			ExcelGo: func(t *testing.T, f string) {
				must(t, MergeCells(f, "Sheet1", "A1:B2"))
				must(t, UnmergeCells(f, "Sheet1", "A1:B2"))
			},
			Ops: `{"op":"merge","sheet":"Sheet1","ref":"A1:B2"},
			       {"op":"unmerge","sheet":"Sheet1","ref":"A1:B2"}`,
		},
		{
			Name: "冻结窗格",
			ExcelGo: func(t *testing.T, f string) {
				must(t, FreezePanes(f, "Sheet1", "B2"))
			},
			Ops: `{"op":"freeze","sheet":"Sheet1","ref":"B2"}`,
		},
		{
			Name: "冻结-仅列",
			ExcelGo: func(t *testing.T, f string) {
				must(t, FreezePanes(f, "Sheet1", "C1"))
			},
			Ops: `{"op":"freeze","sheet":"Sheet1","ref":"C1"}`,
		},
		{
			Name: "自动筛选",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellStr(f, "Sheet1", "A1", "列A"))
				must(t, AutoFilter(f, "Sheet1", "A1:C10"))
			},
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"列A"},
			       {"op":"autofilter","sheet":"Sheet1","ref":"A1:C10"}`,
		},
		{
			Name: "打印区域与重复打印行",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetSheetProps(f, "Sheet1", SheetProps{
					PrintArea: "A1:D20", PrintTitleRows: "1:1",
				}))
			},
			Ops: `{"op":"print_area","sheet":"Sheet1","ref":"A1:D20"},
			       {"op":"print_titles","sheet":"Sheet1","rows":"1:1"}`,
		},

		// ---------- 其它功能 ----------
		{
			Name: "批注",
			ExcelGo: func(t *testing.T, f string) {
				must(t, AddComment(f, "Sheet1", "A1", "这是批注", "张三"))
				must(t, AddComment(f, "Sheet1", "B2", "含特殊字符 <&>\"'", "李四"))
			},
			Ops: `{"op":"comment","sheet":"Sheet1","cell":"A1","text":"这是批注","author":"张三"},
			       {"op":"comment","sheet":"Sheet1","cell":"B2","text":"含特殊字符 <&>\"'","author":"李四"}`,
		},
		{
			Name: "超链接",
			ExcelGo: func(t *testing.T, f string) {
				must(t, AddHyperlink(f, "Sheet1", "A1", "https://example.com", "示例站点"))
				must(t, AddHyperlink(f, "Sheet1", "A2", "mailto:a@b.com", "邮件"))
			},
			Ops: `{"op":"hyperlink","sheet":"Sheet1","cell":"A1","url":"https://example.com","display":"示例站点"},
			       {"op":"hyperlink","sheet":"Sheet1","cell":"A2","url":"mailto:a@b.com","display":"邮件"}`,
		},
		{
			Name: "表格",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetCellStr(f, "Sheet1", "A1", "列1"))
				must(t, SetCellStr(f, "Sheet1", "B1", "列2"))
				must(t, SetCellStr(f, "Sheet1", "A2", "a"))
				must(t, SetCellStr(f, "Sheet1", "B2", "b"))
				must(t, AddTable(f, "Sheet1", "A1:B2", "表1"))
			},
			Ops: `{"op":"set_str","sheet":"Sheet1","cell":"A1","value":"列1"},
			       {"op":"set_str","sheet":"Sheet1","cell":"B1","value":"列2"},
			       {"op":"set_str","sheet":"Sheet1","cell":"A2","value":"a"},
			       {"op":"set_str","sheet":"Sheet1","cell":"B2","value":"b"},
			       {"op":"table","sheet":"Sheet1","ref":"A1:B2","name":"表1"}`,
		},
		{
			Name: "数据验证-列表",
			ExcelGo: func(t *testing.T, f string) {
				must(t, AddDataValidationList(f, "Sheet1", "A1:A5",
					[]string{"是", "否"}, true))
			},
			Ops: `{"op":"datavalidation","sheet":"Sheet1","ref":"A1:A5",
			       "type":"list","values":["是","否"],"allow_blank":true}`,
		},
		{
			Name: "定义名称",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetDefinedName(f, "销售区域", "Sheet1!$A$1:$C$10"))
				must(t, SetDefinedName(f, "全局名", "Sheet1!$D$1"))
			},
			Ops: `{"op":"defined_name","name":"销售区域","refers_to":"Sheet1!$A$1:$C$10"},
			       {"op":"defined_name","name":"全局名","refers_to":"Sheet1!$D$1"}`,
		},

		// ---------- 工作表管理 ----------
		{
			Name: "多表与顺序",
			ExcelGo: func(t *testing.T, f string) {
				_, err := NewSheet(f, "第二表")
				must(t, err)
				_, err = NewSheet(f, "第三表")
				must(t, err)
				must(t, SetCellStr(f, "第二表", "A1", "x"))
			},
			Ops: `{"op":"create_sheet","name":"第二表"},
			       {"op":"create_sheet","name":"第三表"},
			       {"op":"set_str","sheet":"第二表","cell":"A1","value":"x"}`,
		},
		{
			Name: "重命名工作表",
			ExcelGo: func(t *testing.T, f string) {
				must(t, RenameSheet(f, "Sheet1", "数据表"))
				must(t, SetCellStr(f, "数据表", "A1", "x"))
			},
			Ops: `{"op":"rename_sheet","old":"Sheet1","new":"数据表"},
			       {"op":"set_str","sheet":"数据表","cell":"A1","value":"x"}`,
		},
		{
			Name: "删除工作表",
			ExcelGo: func(t *testing.T, f string) {
				_, err := NewSheet(f, "临时表")
				must(t, err)
				must(t, DeleteSheet(f, "临时表"))
			},
			Ops: `{"op":"create_sheet","name":"临时表"},
			       {"op":"delete_sheet","name":"临时表"}`,
		},
		{
			Name: "工作表可见性",
			ExcelGo: func(t *testing.T, f string) {
				_, err := NewSheet(f, "隐藏表")
				must(t, err)
				must(t, SetSheetVisible(f, "隐藏表", "hidden"))
			},
			Ops: `{"op":"create_sheet","name":"隐藏表"},
			       {"op":"sheet_visible","name":"隐藏表","state":"hidden"}`,
		},
		{
			Name: "标签颜色",
			ExcelGo: func(t *testing.T, f string) {
				must(t, SetSheetProps(f, "Sheet1", SheetProps{TabColor: "FF0000"}))
			},
			Ops: `{"op":"tab_color","name":"Sheet1","color":"FF0000"}`,
		},
	}

	for _, c := range cases {
		c := c
		t.Run(c.Name, func(t *testing.T) { runDiffCase(t, c) })
	}
}

func must(t *testing.T, err error) {
	t.Helper()
	if err != nil {
		t.Fatalf("操作失败: %v", err)
	}
}

var _ = fmt.Sprint
