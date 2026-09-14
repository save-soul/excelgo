package main

import (
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/save-soul/excelgo"
)

func runMergeCells(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo mergecells <文件.xlsx> <表> <区域>")
		os.Exit(1)
	}
	if err := excelgo.MergeCells(pos[0], pos[1], pos[2]); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已合并 %s!%s\n", pos[1], pos[2])
}

func runUnmerge(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo unmerge <文件.xlsx> <表> <区域>")
		os.Exit(1)
	}
	if err := excelgo.UnmergeCells(pos[0], pos[1], pos[2]); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已取消合并 %s!%s\n", pos[1], pos[2])
}

func runColWidth(args []string) {
	flagArgs, pos := splitFlagsAnyPos(args)
	fs := flag.NewFlagSet("colwidth", flag.ExitOnError)
	maxCol := fs.Int("max", 0, "末列（默认与起列相同，即单列）")
	fs.Parse(flagArgs)
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo colwidth <文件.xlsx> <表> <列(1基)> <宽度> [--max 末列]")
		os.Exit(1)
	}
	col, err := strconv.Atoi(pos[2])
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: 列须为整数: %v\n", err)
		os.Exit(1)
	}
	width, err := strconv.ParseFloat(pos[3], 64)
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: 宽度须为数值: %v\n", err)
		os.Exit(1)
	}
	mx := *maxCol
	if mx < col {
		mx = col
	}
	if err := excelgo.SetColWidthRange(pos[0], pos[1], col, mx, width); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已设置 %s 列宽 %v\n", pos[1], width)
}

func runRowHeight(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo rowheight <文件.xlsx> <表> <行(1基)> <高度>")
		os.Exit(1)
	}
	row, err := strconv.Atoi(pos[2])
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: 行须为整数: %v\n", err)
		os.Exit(1)
	}
	h, err := strconv.ParseFloat(pos[3], 64)
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: 高度须为数值: %v\n", err)
		os.Exit(1)
	}
	if err := excelgo.SetRowHeight(pos[0], pos[1], row, h); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已设置 %s 第 %d 行高 %v\n", pos[1], row, h)
}

func runFreeze(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo freeze <文件.xlsx> <表> <冻结点(如 A2/B1/B2)>")
		os.Exit(1)
	}
	if err := excelgo.FreezePanes(pos[0], pos[1], pos[2]); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已在 %s!%s 处冻结窗格\n", pos[1], pos[2])
}

func runHyperlink(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo hyperlink <文件.xlsx> <表> <单元格> <URL> [显示文本]")
		os.Exit(1)
	}
	text := ""
	if len(pos) >= 5 {
		text = pos[4]
	}
	if err := excelgo.AddHyperlink(pos[0], pos[1], pos[2], pos[3], text); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已为 %s!%s 添加超链接\n", pos[1], pos[2])
}

func runAutoFilter(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo autofilter <文件.xlsx> <表> <区域>")
		os.Exit(1)
	}
	if err := excelgo.AutoFilter(pos[0], pos[1], pos[2]); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已为 %s!%s 添加自动筛选\n", pos[1], pos[2])
}

func runDropdown(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo dropdown <文件.xlsx> <表> <区域> <选项CSV>")
		os.Exit(1)
	}
	values := strings.Split(pos[3], ",")
	if err := excelgo.AddDataValidationList(pos[0], pos[1], pos[2], values, true); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已为 %s!%s 添加下拉列表（%d 项）\n", pos[1], pos[2], len(values))
}

func runValidation(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 5 {
		fmt.Fprintln(os.Stderr, "用法: excelgo validation <文件.xlsx> <表> <区域> <类型> <公式1> [公式2]")
		os.Exit(1)
	}
	f2 := ""
	if len(pos) >= 6 {
		f2 = pos[5]
	}
	if err := excelgo.AddDataValidation(pos[0], pos[1], pos[2], pos[3], "", pos[4], f2, true); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已为 %s!%s 添加数据验证（%s）\n", pos[1], pos[2], pos[3])
}

func runCFormat(args []string) {
	fs := flag.NewFlagSet("cformat", flag.ExitOnError)
	fill := fs.String("fill", "", "命中时填充色 ARGB")
	bold := fs.Bool("bold", false, "命中时加粗")
	flagArgs, pos := splitFlagsAnyPos(args)
	fs.Parse(flagArgs)
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo cformat <文件.xlsx> <表> <区域> <公式> [--fill ARGB] [--bold]")
		os.Exit(1)
	}
	st := excelgo.Style{}
	if *fill != "" {
		st.Fill = &excelgo.FillStyle{Color: *fill}
	}
	if *bold {
		st.Font = &excelgo.FontStyle{Bold: true}
	}
	if err := excelgo.SetConditionalFormat(pos[0], pos[1], pos[2], "expression", pos[3], 1, st); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已为 %s!%s 添加条件格式（公式=%s）\n", pos[1], pos[2], pos[3])
}

func runTable(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo table <文件.xlsx> <表> <区域> <表名>")
		os.Exit(1)
	}
	if err := excelgo.AddTable(pos[0], pos[1], pos[2], pos[3]); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已创建表格 %s（%s!%s）\n", pos[3], pos[1], pos[2])
}

func runComment(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo comment <文件.xlsx> <表> <单元格> <文本> [作者]")
		os.Exit(1)
	}
	author := ""
	if len(pos) >= 5 {
		author = pos[4]
	}
	if err := excelgo.AddComment(pos[0], pos[1], pos[2], pos[3], author); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已为 %s!%s 添加批注\n", pos[1], pos[2])
}

func runSheetProps(args []string) {
	fs := flag.NewFlagSet("sheetprops", flag.ExitOnError)
	tabColor := fs.String("tab-color", "", "标签色 ARGB（如 FFFF0000）")
	gridLines := fs.Int("grid-lines", -1, "网格线 1=显示 0=隐藏（-1 不变）")
	zoom := fs.Int("zoom", 0, "缩放百分比（0 不变）")
	flagArgs, pos := splitFlagsAnyPos(args)
	fs.Parse(flagArgs)
	if len(pos) < 2 {
		fmt.Fprintln(os.Stderr, "用法: excelgo sheetprops <文件.xlsx> <表> [--tab-color ARGB] [--grid-lines 0|1] [--zoom N]")
		os.Exit(1)
	}
	opts := excelgo.SheetProps{TabColor: *tabColor, Zoom: *zoom}
	if *gridLines == 0 {
		v := false
		opts.ShowGridLines = &v
	} else if *gridLines == 1 {
		v := true
		opts.ShowGridLines = &v
	}
	if err := excelgo.SetSheetProps(pos[0], pos[1], opts); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已更新 %s 工作表属性\n", pos[1])
}

func runProtect(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 2 {
		fmt.Fprintln(os.Stderr, "用法: excelgo protect <文件.xlsx> <表> [密码哈希]")
		os.Exit(1)
	}
	pwd := ""
	if len(pos) >= 3 {
		pwd = pos[2]
	}
	if err := excelgo.ProtectSheet(pos[0], pos[1], pwd); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已保护工作表 %s\n", pos[1])
}

func runGroupRows(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 5 {
		fmt.Fprintln(os.Stderr, "用法: excelgo grouprows <文件.xlsx> <表> <起行> <止行> <级别>")
		os.Exit(1)
	}
	r1, _ := strconv.Atoi(pos[2])
	r2, _ := strconv.Atoi(pos[3])
	lv, _ := strconv.Atoi(pos[4])
	if err := excelgo.GroupRows(pos[0], pos[1], r1, r2, lv); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已对 %s 行 %d-%d 设置分组级别 %d\n", pos[1], r1, r2, lv)
}

func runGroupCols(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 5 {
		fmt.Fprintln(os.Stderr, "用法: excelgo groupcols <文件.xlsx> <表> <起列> <止列> <级别>")
		os.Exit(1)
	}
	c1, _ := strconv.Atoi(pos[2])
	c2, _ := strconv.Atoi(pos[3])
	lv, _ := strconv.Atoi(pos[4])
	if err := excelgo.GroupCols(pos[0], pos[1], c1, c2, lv); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已对 %s 列 %d-%d 设置分组级别 %d\n", pos[1], c1, c2, lv)
}

func runReplace(args []string) {
	_, pos := splitFlagsAnyPos(args)
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo replace <文件.xlsx> <表> <旧文本> <新文本>")
		os.Exit(1)
	}
	n, err := excelgo.ReplaceText(pos[0], pos[1], pos[2], pos[3])
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：替换 %d 处\n", n)
}

func runDocProps(args []string) {
	fs := flag.NewFlagSet("docprops", flag.ExitOnError)
	title := fs.String("title", "", "标题")
	author := fs.String("author", "", "作者")
	subject := fs.String("subject", "", "主题")
	flagArgs, pos := splitFlagsAnyPos(args)
	fs.Parse(flagArgs)
	if len(pos) < 1 {
		fmt.Fprintln(os.Stderr, "用法: excelgo docprops <文件.xlsx> [--title T] [--author A] [--subject S]")
		os.Exit(1)
	}
	props := map[string]string{}
	if *title != "" {
		props["title"] = *title
	}
	if *author != "" {
		props["creator"] = *author
	}
	if *subject != "" {
		props["subject"] = *subject
	}
	if err := excelgo.SetDocProps(pos[0], props); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已更新文档属性\n")
}
