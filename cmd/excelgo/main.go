// 命令行入口：Excel 工作表复制 / 跨工作簿合并 / 工作表管理 / 单元格 / 图片 / 行列操作。
//
// 用法概览：
//
//	excelgo copy   <输入.xlsx> <输出.xlsx> <表名或索引> [--independent] [--suffix 后缀]
//	excelgo merge  <目标.xlsx> <源.xlsx>:<表> [源.xlsx>:<表> ...] [--suffix 冲突后缀]
//	excelgo list   <文件.xlsx>
//	excelgo newsheet <文件.xlsx> <新表名>
//	excelgo delsheet <文件.xlsx> <表名或索引>
//	excelgo movesheet <文件.xlsx> <表名或索引> <目标位置(1基)>
//	excelgo renamesheet <文件.xlsx> <旧名> <新名>
//	excelgo getcell <文件.xlsx> <表> <单元格>         例 A1
//	excelgo setcell <文件.xlsx> <表> <单元格> <值> [--type str|num|bool|formula]
//	excelgo addpic <文件.xlsx> <表> <图片路径> [--cell A1] [--col-off 0] [--row-off 0] [--scale 1]
//	excelgo addcellpic <文件.xlsx> <表> <单元格> <图片路径>
//	excelgo rows  insert|remove <文件.xlsx> <表> <行号(1基)> <数量>
//	excelgo cols  insert|remove <文件.xlsx> <表> <列号(1基)> <数量>
//
// 所有写操作均尽量不破坏原有格式、数据、布局、图片。
package main

import (
	"flag"
	"fmt"
	"os"
	"strconv"
	"strings"

	"github.com/save-soul/excelgo"
)

func main() {
	if len(os.Args) < 2 {
		printUsage()
		os.Exit(1)
	}

	sub := os.Args[1]
	rest := os.Args[2:]

	switch sub {
	case "copy":
		runCopy(rest)
	case "merge":
		runMerge(rest)
	case "list":
		runList(rest)
	case "newsheet":
		runNewSheet(rest)
	case "delsheet":
		runDeleteSheet(rest)
	case "movesheet":
		runMoveSheet(rest)
	case "renamesheet":
		runRenameSheet(rest)
	case "getcell":
		runGetCell(rest)
	case "setcell":
		runSetCell(rest)
	case "addpic":
		runAddPic(rest)
	case "addcellpic":
		runAddCellPic(rest)
	case "rows":
		runRows(rest)
	case "cols":
		runCols(rest)
	case "rangeget":
		runRangeGet(rest)
	case "rangeset":
		runRangeSet(rest)
	case "setstyle":
		runSetStyle(rest)
	case "mergecells":
		runMergeCells(rest)
	case "unmerge":
		runUnmerge(rest)
	case "colwidth":
		runColWidth(rest)
	case "rowheight":
		runRowHeight(rest)
	case "freeze":
		runFreeze(rest)
	case "hyperlink":
		runHyperlink(rest)
	case "autofilter":
		runAutoFilter(rest)
	case "dropdown":
		runDropdown(rest)
	case "validation":
		runValidation(rest)
	case "cformat":
		runCFormat(rest)
	case "table":
		runTable(rest)
	case "comment":
		runComment(rest)
	case "sheetprops":
		runSheetProps(rest)
	case "protect":
		runProtect(rest)
	case "grouprows":
		runGroupRows(rest)
	case "groupcols":
		runGroupCols(rest)
	case "replace":
		runReplace(rest)
	case "docprops":
		runDocProps(rest)
	case "-h", "--help", "help":
		printUsage()
	default:
		fmt.Fprintf(os.Stderr, "未知子命令: %q\n\n", sub)
		printUsage()
		os.Exit(1)
	}
}

func printUsage() {
	fmt.Fprintf(os.Stderr, `excelgo - 纯 Go 的 Excel 工具（工作表复制/合并/管理/单元格/图片/行列）

用法：
  excelgo copy       <输入.xlsx> <输出.xlsx> <表名或索引> [flags]
  excelgo merge      <目标.xlsx> <源.xlsx>:<表> [源.xlsx>:<表> ...] [flags]
  excelgo list       <文件.xlsx>
  excelgo newsheet   <文件.xlsx> <新表名>
  excelgo delsheet   <文件.xlsx> <表名或索引>
  excelgo movesheet  <文件.xlsx> <表名或索引> <目标位置(1基)>
  excelgo renamesheet <文件.xlsx> <旧名> <新名>
  excelgo getcell    <文件.xlsx> <表> <单元格>
  excelgo setcell    <文件.xlsx> <表> <单元格> <值> [--type str|num|bool|formula]
  excelgo addpic     <文件.xlsx> <表> <图片路径> [--cell A1] [--col-off N] [--row-off N] [--scale F]
  excelgo addcellpic <文件.xlsx> <表> <单元格> <图片路径>
  excelgo rows       insert|remove <文件.xlsx> <表> <行号(1基)> <数量>
  excelgo cols       insert|remove <文件.xlsx> <表> <列号(1基)> <数量>
  excelgo rangeget   <文件.xlsx> <表> <区域>              例 A1:C10
  excelgo rangeset   <文件.xlsx> <表> <区域> <数据>        例 "a,b;1,2"（行以;分，单元格以,分；值以=开头视为公式）
  excelgo setstyle    <文件.xlsx> <表> <单元格或区域> [样式flags]
  excelgo mergecells  <文件.xlsx> <表> <区域>                合并单元格（如 A1:C3）
  excelgo unmerge     <文件.xlsx> <表> <区域>                取消合并
  excelgo colwidth    <文件.xlsx> <表> <列(1基)> <宽度> [--max 末列]
  excelgo rowheight   <文件.xlsx> <表> <行(1基)> <高度>
  excelgo freeze      <文件.xlsx> <表> <冻结点>              如 A2(冻首行) B1(冻首列) B2(双冻)
  excelgo hyperlink   <文件.xlsx> <表> <单元格> <URL> [显示文本]
  excelgo autofilter  <文件.xlsx> <表> <区域>
  excelgo dropdown    <文件.xlsx> <表> <区域> <选项CSV>       下拉列表（如 甲,乙,丙）
  excelgo validation  <文件.xlsx> <表> <区域> <类型> <公式1> [公式2]
  excelgo cformat     <文件.xlsx> <表> <区域> <公式> [--fill ARGB] [--bold]
  excelgo table       <文件.xlsx> <表> <区域> <表名>          结构化表格
  excelgo comment     <文件.xlsx> <表> <单元格> <文本> [作者]
  excelgo sheetprops  <文件.xlsx> <表> [--tab-color ARGB] [--grid-lines 0|1] [--zoom N]
  excelgo protect     <文件.xlsx> <表> [密码哈希]
  excelgo grouprows   <文件.xlsx> <表> <起行> <止行> <级别>
  excelgo groupcols   <文件.xlsx> <表> <起列> <止列> <级别>
  excelgo replace     <文件.xlsx> <表> <旧文本> <新文本>
  excelgo docprops    <文件.xlsx> [--title T] [--author A] [--subject S]

copy flags:    --independent 独立媒体副本  --suffix 后缀(默认 _copy)
merge flags:   --suffix 冲突后缀(默认 _merge)
setcell types: str(共享字符串,默认) | num(数值) | bool(布尔) | formula(公式,需 --result 可选)
addpic flags:  --cell 锚定单元格(默认 A1) --col-off 列偏移像素(默认0) --row-off 行偏移像素(默认0) --scale 缩放比(默认1)
setstyle flags: --bold --italic --underline single --strike --size 11 --name 字体 --color ARGB
               --fill ARGB --numfmt 0.00 --align-h left|center|right --align-v top|center|bottom --wrap
               （单元格参数可写单格 A1 或区域 A1:C3，区域批量套用同一样式）

示例：
  excelgo copy  "in.xlsx" "out.xlsx" Sheet1
  excelgo merge "target.xlsx" "src1.xlsx:数据A" "src2.xlsx:数据B"
  excelgo list  "book.xlsx"
  excelgo newsheet "book.xlsx" "汇总"
  excelgo setcell "book.xlsx" Sheet1 A1 "你好" --type str
  excelgo rows insert "book.xlsx" Sheet1 3 2
  excelgo addpic "book.xlsx" Sheet1 logo.png --cell B2 --scale 0.5
  excelgo rangeget "book.xlsx" Sheet1 A1:C3
  excelgo rangeset "book.xlsx" Sheet1 A1 "1,2,3;4,5,6"
  excelgo setstyle "book.xlsx" Sheet1 A1:C3 --bold --fill FFFFFF00 --align-h center
`)
}

func runCopy(args []string) {
	fs := flag.NewFlagSet("copy", flag.ExitOnError)
	independent := fs.Bool("independent", false, "复制体拥有独立媒体副本（可单独编辑/删除），默认共享媒体")
	suffix := fs.String("suffix", "_copy", "新工作表名后缀")
	fs.Parse(args)

	pos := fs.Args()
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo copy <输入.xlsx> <输出.xlsx> <工作表名或索引> [--independent] [--suffix 后缀]")
		os.Exit(1)
	}
	inputFile := pos[0]
	outputFile := pos[1]
	sheetRef := pos[2]

	var opts []excelgo.Option
	opts = append(opts, excelgo.WithSuffix(*suffix))
	if *independent {
		opts = append(opts, excelgo.WithMedia(excelgo.MediaIndependent))
	} else {
		opts = append(opts, excelgo.WithMedia(excelgo.MediaShared))
	}

	if err := excelgo.CopySheet(inputFile, outputFile, sheetRef, opts...); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已将工作表 %q 复制至 %q（策略：%s）\n", sheetRef, outputFile, strategyName(*independent))
}

func runMerge(args []string) {
	fs := flag.NewFlagSet("merge", flag.ExitOnError)
	suffix := fs.String("suffix", "_merge", "命名冲突时追加的后缀")
	fs.Parse(args)

	pos := fs.Args()
	if len(pos) < 2 {
		fmt.Fprintln(os.Stderr, "用法: excelgo merge <目标.xlsx> <源.xlsx>:<表名或索引> [源.xlsx:表 ...] [--suffix 冲突后缀]")
		os.Exit(1)
	}
	dst := pos[0]
	var sources []excelgo.SourceRef
	for _, spec := range pos[1:] {
		idx := strings.LastIndex(spec, ":")
		if idx < 0 {
			fmt.Fprintf(os.Stderr, "错误: 源参数 %q 缺少 \":表名\" 后缀，正确格式如 \"src.xlsx:Sheet1\"\n", spec)
			os.Exit(1)
		}
		wb := spec[:idx]
		sheet := spec[idx+1:]
		if wb == "" || sheet == "" {
			fmt.Fprintf(os.Stderr, "错误: 源参数 %q 解析失败，正确格式如 \"src.xlsx:Sheet1\"\n", spec)
			os.Exit(1)
		}
		sources = append(sources, excelgo.SourceRef{Workbook: wb, Sheet: sheet})
	}

	if err := excelgo.MergeWorkbook(dst, sources, excelgo.WithMergeSuffix(*suffix)); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已将 %d 个工作表合并进 %q（冲突后缀：%s）\n", len(sources), dst, *suffix)
}

func runList(args []string) {
	fs := flag.NewFlagSet("list", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 1 {
		fmt.Fprintln(os.Stderr, "用法: excelgo list <文件.xlsx>")
		os.Exit(1)
	}
	list, err := excelgo.GetSheetList(pos[0])
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("工作表（共 %d 个）：\n", len(list))
	for i, name := range list {
		fmt.Printf("  %d. %s\n", i+1, name)
	}
}

func runNewSheet(args []string) {
	fs := flag.NewFlagSet("newsheet", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 2 {
		fmt.Fprintln(os.Stderr, "用法: excelgo newsheet <文件.xlsx> <新表名>")
		os.Exit(1)
	}
	idx, err := excelgo.NewSheet(pos[0], pos[1])
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已新建工作表 %q（序号 %d）\n", pos[1], idx)
}

func runDeleteSheet(args []string) {
	fs := flag.NewFlagSet("delsheet", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 2 {
		fmt.Fprintln(os.Stderr, "用法: excelgo delsheet <文件.xlsx> <表名或索引>")
		os.Exit(1)
	}
	if err := excelgo.DeleteSheet(pos[0], pos[1]); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已删除工作表 %q\n", pos[1])
}

func runMoveSheet(args []string) {
	fs := flag.NewFlagSet("movesheet", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo movesheet <文件.xlsx> <表名或索引> <目标位置(1基)>")
		os.Exit(1)
	}
	to, err := strconv.Atoi(pos[2])
	if err != nil {
		fmt.Fprintln(os.Stderr, "目标位置必须为整数（1 基）")
		os.Exit(1)
	}
	if err := excelgo.MoveSheet(pos[0], pos[1], to); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已将工作表 %q 移动到位置 %d\n", pos[1], to)
}

func runRenameSheet(args []string) {
	fs := flag.NewFlagSet("renamesheet", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo renamesheet <文件.xlsx> <旧名> <新名>")
		os.Exit(1)
	}
	if err := excelgo.RenameSheet(pos[0], pos[1], pos[2]); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已重命名工作表 %q -> %q\n", pos[1], pos[2])
}

func runGetCell(args []string) {
	fs := flag.NewFlagSet("getcell", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo getcell <文件.xlsx> <表> <单元格>")
		os.Exit(1)
	}
	val, err := excelgo.GetCell(pos[0], pos[1], pos[2])
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Println(val)
}

func runSetCell(args []string) {
	// 提取 --type / --result（支持放在任意位置，避免 flag 包的顺序限制）
	typ, result, posArgs := extractFlags(args, "str", "")
	if len(posArgs) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo setcell <文件.xlsx> <表> <单元格> <值> [--type str|num|bool|formula] [--result 预计算结果]")
		os.Exit(1)
	}
	file, sheet, cell, value := posArgs[0], posArgs[1], posArgs[2], posArgs[3]
	var err error
	switch typ {
	case "str":
		err = excelgo.SetCellStr(file, sheet, cell, value)
	case "num":
		var f float64
		f, err = strconv.ParseFloat(value, 64)
		if err == nil {
			err = excelgo.SetCellNumeric(file, sheet, cell, f)
		}
	case "bool":
		var b bool
		b, err = strconv.ParseBool(value)
		if err == nil {
			err = excelgo.SetCellBool(file, sheet, cell, b)
		}
	case "formula":
		err = excelgo.SetCellFormula(file, sheet, cell, value, result)
	default:
		fmt.Fprintf(os.Stderr, "未知 --type: %q（支持 str|num|bool|formula）\n", typ)
		os.Exit(1)
	}
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已写入 %s!%s = %q (type=%s)\n", sheet, cell, value, typ)
}

// extractFlags 从参数中提取 --type 与 --result（支持任意位置），返回 (type, result, 剩余位置参数)。
func extractFlags(args []string, defType, defResult string) (string, string, []string) {
	typ := defType
	result := defResult
	var pos []string
	i := 0
	for i < len(args) {
		a := args[i]
		switch a {
		case "--type":
			if i+1 < len(args) {
				typ = args[i+1]
				i += 2
				continue
			}
		case "--result":
			if i+1 < len(args) {
				result = args[i+1]
				i += 2
				continue
			}
		default:
			if strings.HasPrefix(a, "--type=") {
				typ = strings.TrimPrefix(a, "--type=")
			} else if strings.HasPrefix(a, "--result=") {
				result = strings.TrimPrefix(a, "--result=")
			} else {
				pos = append(pos, a)
			}
			i++
		}
	}
	return typ, result, pos
}

func runAddPic(args []string) {
	fs := flag.NewFlagSet("addpic", flag.ExitOnError)
	cell := fs.String("cell", "A1", "锚定单元格，如 B2")
	colOff := fs.Int("col-off", 0, "列偏移像素")
	rowOff := fs.Int("row-off", 0, "行偏移像素")
	scale := fs.Float64("scale", 1.0, "缩放比例（默认 1）")
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo addpic <文件.xlsx> <表> <图片路径> [--cell A1] [--col-off N] [--row-off N] [--scale F]")
		os.Exit(1)
	}
	opts := &excelgo.PictureOptions{
		ScaleX: *scale, ScaleY: *scale,
		Position: &excelgo.PicturePosition{Cell: *cell, ColOffset: *colOff, RowOffset: *rowOff},
	}
	if err := excelgo.AddPicture(pos[0], pos[1], pos[2], opts); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已在 %s!%s 插入浮动图片 %q（缩放 %v）\n", pos[1], *cell, pos[2], *scale)
}

func runAddCellPic(args []string) {
	fs := flag.NewFlagSet("addcellpic", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo addcellpic <文件.xlsx> <表> <单元格> <图片路径>")
		os.Exit(1)
	}
	if err := excelgo.AddCellPicture(pos[0], pos[1], pos[2], pos[3]); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已在 %s!%s 插入 WPS 单元格内嵌图片 %q\n", pos[1], pos[2], pos[3])
}

func runRows(args []string) {
	fs := flag.NewFlagSet("rows", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo rows insert|remove <文件.xlsx> <表> <行号(1基)> <数量>")
		os.Exit(1)
	}
	op, file, sheet, rowStr, nStr := pos[0], pos[1], pos[2], pos[3], pos[4]
	row, err := strconv.Atoi(rowStr)
	if err != nil {
		fmt.Fprintln(os.Stderr, "行号必须为整数")
		os.Exit(1)
	}
	n, err := strconv.Atoi(nStr)
	if err != nil {
		fmt.Fprintln(os.Stderr, "数量必须为整数")
		os.Exit(1)
	}
	switch op {
	case "insert":
		err = excelgo.InsertRows(file, sheet, row, n)
	case "remove":
		err = excelgo.RemoveRows(file, sheet, row, n)
	default:
		fmt.Fprintln(os.Stderr, "rows 操作仅支持 insert 或 remove")
		os.Exit(1)
	}
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已对 %s!%s 执行 rows %s（行 %d，数量 %d）\n", file, sheet, op, row, n)
}

func runCols(args []string) {
	fs := flag.NewFlagSet("cols", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo cols insert|remove <文件.xlsx> <表> <列号(1基)> <数量>")
		os.Exit(1)
	}
	op, file, sheet, colStr, nStr := pos[0], pos[1], pos[2], pos[3], pos[4]
	col, err := strconv.Atoi(colStr)
	if err != nil {
		fmt.Fprintln(os.Stderr, "列号必须为整数")
		os.Exit(1)
	}
	n, err := strconv.Atoi(nStr)
	if err != nil {
		fmt.Fprintln(os.Stderr, "数量必须为整数")
		os.Exit(1)
	}
	switch op {
	case "insert":
		err = excelgo.InsertCols(file, sheet, col, n)
	case "remove":
		err = excelgo.RemoveCols(file, sheet, col, n)
	default:
		fmt.Fprintln(os.Stderr, "cols 操作仅支持 insert 或 remove")
		os.Exit(1)
	}
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已对 %s!%s 执行 cols %s（列 %d，数量 %d）\n", file, sheet, op, col, n)
}

func strategyName(independent bool) string {
	if independent {
		return "独立媒体副本"
	}
	return "共享媒体"
}

func runRangeGet(args []string) {
	fs := flag.NewFlagSet("rangeget", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo rangeget <文件.xlsx> <表> <区域>  例 A1:C10")
		os.Exit(1)
	}
	grid, err := excelgo.GetRange(pos[0], pos[1], pos[2])
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	for _, row := range grid {
		fmt.Println(strings.Join(row, "\t"))
	}
}

func runRangeSet(args []string) {
	fs := flag.NewFlagSet("rangeset", flag.ExitOnError)
	fs.Parse(args)
	pos := fs.Args()
	if len(pos) < 4 {
		fmt.Fprintln(os.Stderr, "用法: excelgo rangeset <文件.xlsx> <表> <区域> <数据>  例 \"a,b;1,2\"（行以;分，单元格以,分；值以=开头视为公式）")
		os.Exit(1)
	}
	file, sheet, rng, data := pos[0], pos[1], pos[2], pos[3]
	rows := strings.Split(data, ";")
	var values [][]interface{}
	for _, r := range rows {
		var line []interface{}
		for _, c := range strings.Split(r, ",") {
			c = strings.TrimSpace(c)
			if c == "" {
				line = append(line, nil)
				continue
			}
			if strings.HasPrefix(c, "=") {
				line = append(line, excelgo.CellFormula{Formula: c[1:]})
				continue
			}
			if f, e := strconv.ParseFloat(c, 64); e == nil {
				line = append(line, f)
			} else {
				line = append(line, c)
			}
		}
		values = append(values, line)
	}
	if err := excelgo.SetRange(file, sheet, rng, values); err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已将 %d 行数据写入 %s!%s\n", len(values), sheet, rng)
}

// splitFlagsAnyPos 将命令行参数拆分为 flag 段（支持任意位置书写）与位置参数。
// 标准库 flag 在遇到首个非 flag 位置参数即停止解析，因此若样式 flag 写在
// 位置参数 <文件> <表> <单元格> 之后会全部失效。这里先做预扫描把 flag 抽出来。
// 已知布尔 flag（无值）：--bold/--italic/--strike/--wrap
// 已知取值 flag（取下一参数或 key=value 形式）：--underline/--size/--name/--color/--fill/--numfmt/--align-h/--align-v
// splitFlagsAnyPos 把任意位置的 --flag / --flag=value / --flag value 抽离出来，
// 让 flag.Parse 不再因首个位置参数而截断后续 flag。约定：所有 flag 以 "--" 开头，
// 位置参数（文件名、单元格、文本值）不以 "--" 开头。若 --flag 不带 "=" 且后接一个
// 同样不以 "--" 开头的 token，则将该 token 视为其值；否则视为无值（bool）flag。
func splitFlagsAnyPos(args []string) (flagArgs, posArgs []string) {
	i := 0
	for i < len(args) {
		a := args[i]
		if key, ok := strings.CutPrefix(a, "--"); ok {
			if strings.Contains(key, "=") {
				flagArgs = append(flagArgs, a)
				i++
				continue
			}
			// --name 形式：看下一 token 是否像「值」（不以 -- 开头且存在）
			if i+1 < len(args) && !strings.HasPrefix(args[i+1], "--") {
				flagArgs = append(flagArgs, a, args[i+1])
				i += 2
				continue
			}
			flagArgs = append(flagArgs, a)
			i++
			continue
		}
		posArgs = append(posArgs, a)
		i++
	}
	return
}

func runSetStyle(args []string) {
	fs := flag.NewFlagSet("setstyle", flag.ExitOnError)
	bold := fs.Bool("bold", false, "加粗")
	italic := fs.Bool("italic", false, "斜体")
	underline := fs.String("underline", "", "下划线 single|double|singleAccounting|doubleAccounting")
	strike := fs.Bool("strike", false, "删除线")
	size := fs.Float64("size", 0, "字号（0=默认）")
	name := fs.String("name", "", "字体名")
	color := fs.String("color", "", "字体颜色 ARGB，如 FFFF0000")
	fill := fs.String("fill", "", "填充色 ARGB，如 FFFFFF00")
	numfmt := fs.String("numfmt", "", "数字格式代码，如 0.00")
	alignH := fs.String("align-h", "", "水平对齐 left|center|right|general|justify")
	alignV := fs.String("align-v", "", "垂直对齐 top|center|bottom|justify")
	wrap := fs.Bool("wrap", false, "自动换行")
	// 预扫描：把分散在任意位置的样式 flag 抽出来，避免 flag.Parse 在首个位置参数处停止。
	flagArgs, pos := splitFlagsAnyPos(args)
	fs.Parse(flagArgs)
	if len(pos) < 3 {
		fmt.Fprintln(os.Stderr, "用法: excelgo setstyle <文件.xlsx> <表> <单元格或区域> [样式flags]")
		os.Exit(1)
	}
	file, sheet, cell := pos[0], pos[1], pos[2]

	style := excelgo.Style{}
	if *bold || *italic || *underline != "" || *strike || *size != 0 || *name != "" || *color != "" {
		style.Font = &excelgo.FontStyle{
			Bold: *bold, Italic: *italic, Underline: *underline, Strike: *strike,
			Size: *size, Name: *name, Color: *color,
		}
	}
	if *fill != "" {
		style.Fill = &excelgo.FillStyle{Color: *fill}
	}
	if *alignH != "" || *alignV != "" || *wrap {
		style.Alignment = &excelgo.AlignmentStyle{Horizontal: *alignH, Vertical: *alignV, WrapText: *wrap}
	}
	if *numfmt != "" {
		style.NumFmt = *numfmt
	}

	idx, err := excelgo.SetCellStyleRange(file, sheet, cell, style)
	if err != nil {
		fmt.Fprintf(os.Stderr, "错误: %v\n", err)
		os.Exit(1)
	}
	fmt.Printf("成功：已将样式（s=%d）应用到 %s!%s\n", idx, sheet, cell)
}
