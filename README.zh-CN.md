# excelgo

> 🌐 English version / 英文文档: [README.md](README.md)

纯 Go 的 Excel（`.xlsx`）工具库，目标是在不依赖 Microsoft Excel / WPS 的前提下，
覆盖工作表复制、跨工作簿合并、读写、样式、行列、图表等常见场景。

当前已实现：
- 工作表复制：将一个工作表（含浮动图片、WPS 单元格内嵌图片、图表、批注、打印区域等）原样复制到**同一工作簿**的新工作表；
- 跨工作簿合并：将多个源工作簿中的工作表搬移为独立工作表，保留页面布局、图片、形状、分页符、打印属性等，并对相同样式做去重合并。

输出是符合 OOXML 规范的 `.xlsx`，可被 Excel / WPS 正常打开。

## 特性

**工作表复制与合并**
- 复制浮动图片（`xl/drawings/drawingN.xml`）
- 复制 WPS「图片嵌入单元格」内联图片（`mc:AlternateContent` / `x14:picture`）
- 复制图表、批注等通过关系引用的部件
- 复制打印区域（`_xlnm.Print_Area`），自动重映射到新表序号与名称
- 可配置媒体（图片）复制策略：共享 vs 独立

**就地编辑（不破坏样式 / 布局 / 数据 / 图片）**
- 工作表生命周期：新建 / 删除 / 移动 / 改名
- 单元格与区域读写（自动类型推断、公式、公式引用随行列操作自动平移）
- 行列插入 / 删除（平移合并区域、列宽、分页符、打印区域）
- 合并单元格、列宽 / 行高 / 隐藏、冻结窗格
- 超链接、自动筛选、数据验证（下拉 / 通用）、条件格式、结构化表格
- 批注、工作表属性（标签色 / 网格线 / 缩放）、工作表保护、单元格样式（字体 / 填充 / 边框 / 对齐 / 数字格式）
- 行列分组大纲、工作表内查找替换、文档核心属性

**使用方式**
- 纯 Go 实现，无 CGO、无外部 Office 依赖
- 同时提供全局函数 API 与 `file → sheet → cell` 链式对象 API（见「面向对象 API」）

## 安装

```bash
go get github.com/save-soul/excelgo
```

## 作为库使用

```go
package main

import (
	"log"

	"github.com/save-soul/excelgo"
)

func main() {
	// 默认：共享媒体（体积小）
	err := excelgo.CopySheet("input.xlsx", "output.xlsx", "Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	// 独立媒体：复制体拥有自己的图片副本，可单独编辑/删除而不影响原表
	err = excelgo.CopySheet("input.xlsx", "output2.xlsx", "Sheet1",
		excelgo.WithMedia(excelgo.MediaIndependent))
	if err != nil {
		log.Fatal(err)
	}

	// 也可用工作表索引（1 基）与自定义后缀
	err = excelgo.CopySheet("input.xlsx", "output3.xlsx", "1",
		excelgo.WithSuffix("_副本"))
	if err != nil {
		log.Fatal(err)
	}
}
```

### 媒体策略说明

| 策略 | 常量 | 行为 | 适用场景 |
| --- | --- | --- | --- |
| 共享媒体 | `excelgo.MediaShared` | 复制体与原表指向同一个 `xl/media/imageN.png` | 体积最小；WPS「嵌入单元格」轻量语义 |
| 独立媒体 | `excelgo.MediaIndependent` | 复制体拥有独立图片副本，引用与关系一并重写 | 复制体需独立编辑/删除，与原表解耦 |

默认策略为 `MediaShared`。

## 命令行工具

仓库附带 CLI（`cmd/excelgo`），支持两个子命令：`copy`（工作表复制）与 `merge`（跨工作簿合并）。

```bash
# 构建
go build -o excelgo ./cmd/excelgo

# —— copy：工作表复制 ——
# 共享媒体（默认）
excelgo copy input.xlsx output.xlsx Sheet1

# 独立媒体副本
excelgo copy --independent input.xlsx output.xlsx Sheet1

# 自定义后缀
excelgo copy --suffix _副本 input.xlsx output.xlsx Sheet1

# —— merge：跨工作簿合并 ——
# 将 src1 的「数据A」表、src2 的「数据B」表搬入 target（target 需已存在，作为容器）
excelgo merge target.xlsx "src1.xlsx:数据A" "src2.xlsx:数据B"

# 命名冲突时追加的后缀（默认 _merge）
excelgo merge target.xlsx "src1.xlsx:Sheet1" --suffix _合
```

> merge 的每个源参数格式为 `工作簿路径:工作表名或索引`（用 `:` 分隔），例如 `"src1.xlsx:数据A"` 或 `"src1.xlsx:1"`。

## API

```go
func CopySheet(src, dst, sheetRef string, opts ...Option) error
```

- `src` 源 `.xlsx` 路径
- `dst` 输出 `.xlsx` 路径（会被覆盖写入）
- `sheetRef` 工作表名称（如 `"Sheet1"`）或 1 基数字索引（如 `"1"`）
- `opts` 可选配置：
  - `WithMedia(excelgo.MediaStrategy)` 设置媒体策略
  - `WithSuffix(string)` 设置新工作表名后缀（默认 `"_copy"`）

### 合并工作表（跨工作簿）

把多个源工作簿中的指定工作表，作为**独立工作表**搬移到同一个目标工作簿中
（合并后各表仍彼此独立）。适合把分散在多个文件里的工作表整合进一个文件。

```go
func MergeWorkbook(dst string, sources []SourceRef, opts ...MergeOption) error
```

- `dst` 目标工作簿路径（**必须已存在**，作为合并容器，会被覆盖写入）
- `sources` 源工作表列表，每项含 `Workbook`（源文件路径）与 `Sheet`（工作表名或 1 基索引）
- `opts` 可选配置：
  - `WithRename(name string)` 统一指定搬入后的工作表名（同名冲突时自动加后缀避让）
  - 命名冲突后缀默认 `"_merge"`

特性与保证：

- **样式合并去重**：跨工作簿搬表时，按 `cellXfs/fontId/fillId/borderId/numFmtId`
  语义比对，复用目标已有的等价样式条目，仅追加确实缺失的样式，避免样式数量溢出。
  单元格的 `s` 索引随之重映射，格式不串。
- **部件随表迁移**：浮动图片、形状（drawing）、媒体、批注、图表等关联部件一并复制
  并重编号、重写关系引用。
- **页面属性保留**：打印区域（`_xlnm.Print_Area`，自动重映射 `localSheetId` 与表名）、
  分页符（`rowBreaks/colBreaks`）、页面布局等随工作表保留。

```go
err := excelgo.MergeWorkbook("汇总.xlsx", []excelgo.SourceRef{
    {Workbook: "1月.xlsx", Sheet: "Sheet1"},
    {Workbook: "2月.xlsx", Sheet: "数据"},
})
```

### 工作表生命周期（就地编辑）

以下操作均直接修改原文件，**仅改动必要部件**（workbook.xml / rels / content types /
对应 worksheet / media / drawing），不破坏其他工作表的内容、样式、布局与图片。

```go
// 新建工作表，返回其 1 基序号（同名自动加后缀避让）
idx, _ := excelgo.NewSheet("wb.xlsx", "新表")

// 删除工作表（连同其 worksheet / drawing / 独占 media 一并清理，共享图片保留）
_ = excelgo.DeleteSheet("wb.xlsx", "Sheet1")        // 名称或 1 基索引
_ = excelgo.DeleteSheet("wb.xlsx", "2")

// 移动工作表到指定位置（toIndex 为 1 基，1 表示最前）
_ = excelgo.MoveSheet("wb.xlsx", "Sheet2", 1)

// 改名（同名自动加后缀避让）
_ = excelgo.RenameSheet("wb.xlsx", "旧名", "新名")
```

### 单元格读写（多种格式）

参考 excelize 的细分写入风格，自动推断或按类型精确写入，保留原单元格 `s` 样式索引。

```go
// 读取（共享字符串自动解析为真实文本；公式返回结果 <v>，无结果返回 "=公式"）
val, _ := excelgo.GetCell("wb.xlsx", "Sheet1", "A1")

// 自动推断类型（string→共享字符串，int/float→数值，bool→布尔）
_ = excelgo.SetCellValue("wb.xlsx", "Sheet1", "A1", "文本")
_ = excelgo.SetCellValue("wb.xlsx", "Sheet1", "A2", 123)
_ = excelgo.SetCellValue("wb.xlsx", "Sheet1", "A3", true)

// 按类型细分写入
_ = excelgo.SetCellStr("wb.xlsx", "Sheet1", "B1", "共享字符串")
_ = excelgo.SetCellInt("wb.xlsx", "Sheet1", "B2", 42)
_ = excelgo.SetCellNumeric("wb.xlsx", "Sheet1", "B3", 3.14)
_ = excelgo.SetCellBool("wb.xlsx", "Sheet1", "B4", true)
_ = excelgo.SetCellInline("wb.xlsx", "Sheet1", "B5", "内联字符串（不依赖 sharedStrings）")
_ = excelgo.SetCellFormula("wb.xlsx", "Sheet1", "B6", "A1*2", "结果预计算值") // result 可空
```

> 写入时新单元格严格置于对应 `<row r="行号">` 内，符合 OOXML 规范；原有单元格的 `s` 样式
> 索引会被保留，未写样式的单元格默认使用样式 0（常规）。

### 插入图片

两种模式，均不破坏工作表已有绘图对象。

```go
// 浮动图片：可设锚定单元格、像素偏移与缩放比例
_ = excelgo.AddPicture("wb.xlsx", "Sheet1", "photo.png", &excelgo.PictureOptions{
    ScaleX: 0.5, ScaleY: 0.5,                       // 缩放比例，默认 1
    Position: &excelgo.PicturePosition{Cell: "H1",  // 锚定单元格
        ColOffset: 10, RowOffset: 5},               // 相对锚点的像素偏移
})
// 也可直接传图片字节：excelgo.AddPictureFromBytes(..., data, opts)

// WPS「图片嵌入单元格」式图片：随单元格移动与缩放（mc:AlternateContent / x14:picture）
_ = excelgo.AddCellPicture("wb.xlsx", "Sheet2", "B2", "photo.png")
// 字节版：excelgo.AddCellPictureFromBytes(..., data)
```

> 浮动图片写入 `xl/drawings/` 与 `xl/media/`，并复用/新建 drawing 部件；WPS 内嵌图片写入
> worksheet 的 `mc:AlternateContent` 块（含 `x14:picture` 与标准 drawingML 回退），同时补全
> 所需命名空间（mc/x14/a/xdr/r），保证 Excel / WPS 均可识别。

### 工作表查询与整表读取

```go
// 工作表列表（按 workbook.xml 顺序）
names, _ := excelgo.GetSheetList("wb.xlsx")
// 1 基序号（不存在返回 0 与错误）
idx, _ := excelgo.GetSheetIndex("wb.xlsx", "Sheet1")
// 是否存在
exists, _ := excelgo.SheetExists("wb.xlsx", "Sheet1")
// 整表读取为二维数组（共享字符串解析、公式取结果，空单元格为 ""）
rows, _ := excelgo.GetSheetData("wb.xlsx", "Sheet1") // 等价于 GetRows
```

### 行列插入与删除

插入/删除行或列。仅平移受影响的元素（`<row r>`、`<c r>`、合并区域 `mergeCell ref`、
列宽 `col min/max`、分页符 `brk`、打印区域 `definedName`），保留单元格 `s` 样式索引与图片，
不破坏其他数据/布局。行/列号均为 1 基。

```go
// 在第 3 行前插入 2 行
_ = excelgo.InsertRows("wb.xlsx", "Sheet1", 3, 2)
// 删除第 3 行起的 1 行
_ = excelgo.RemoveRows("wb.xlsx", "Sheet1", 3, 1)
// 在第 2 列（B 列）前插入 1 列
_ = excelgo.InsertCols("wb.xlsx", "Sheet1", 2, 1)
// 删除第 2 列起的 1 列
_ = excelgo.RemoveCols("wb.xlsx", "Sheet1", 2, 1)
```

### 区域读写（按矩形块）

以矩形区域（如 `"A1:C10"`）为单位的整块读取与写入，复用在内存 map 上的一次读盘 / 逐格改写 /
一次写回，保留原 `s` 样式索引与布局。

```go
// 读取区域为二维字符串（共享字符串解析、公式取结果，空单元格为 ""）
grid, _ := excelgo.GetRange("wb.xlsx", "Sheet1", "A1:C10")

// 写入区域：单格起点（如 "A1"）按数据尺寸自动扩展；含 ":" 的区域按边界校验
_ = excelgo.SetRange("wb.xlsx", "Sheet1", "A1", [][]interface{}{
    {"姓名", "年龄", "分数"},
    {"张三", 18, 95.5},
    {"李四", 20, nil},                       // nil 跳过该单元格
    {excelgo.CellFormula{Formula: "B2+C2", Result: "113.5"}}, // 公式单元格
})
```

- `GetRange(filename, sheetRef, rangeRef)` → `([][]string, error)`，区域或单格均可。
- `SetRange(filename, sheetRef, rangeRef, values)` 支持 `string / int / float64 / bool / nil /
  excelgo.CellFormula`；单元格值以 `=` 开头的字符串会自动按公式写入。
- 公式单元格用 `excelgo.CellFormula{Formula: "A1+B1", Result: "结果"}`，`Result` 可空（仅写 `<f>`
  不写缓存 `<v>`，由 Excel 打开时重算）。

### 样式设置 API

程序化设置单元格字体 / 填充 / 边框 / 对齐 / 数字格式，写入 `xl/styles.xml`；`font / fill /
border / cellXfs` 均按内容去重，避免样式表膨胀，单元格 `s` 索引随之复用。

```go
style := excelgo.Style{
    Font:      &excelgo.FontStyle{Bold: true, Color: "FFFF0000", Size: 14, Name: "微软雅黑"},
    Fill:      &excelgo.FillStyle{Color: "FFFFFF00", PatternType: "solid"},
    Border:    &excelgo.BorderStyle{                       // 四边/对角线均可单独设置
        Left:  excelgo.BorderSide{Style: "thin", Color: "FF000000"},
        Right: excelgo.BorderSide{Style: "thin", Color: "FF000000"},
    },
    Alignment: &excelgo.AlignmentStyle{Horizontal: "center", Vertical: "center", WrapText: true},
    NumFmt:    "0.00",                                     // 数字格式代码；内置 id 直接复用，自定义从 164 起
}

// 应用到区域（推荐）；同样式再次应用会复用同一 cellXf 索引（去重）
idx, _ := excelgo.SetCellStyleRange("wb.xlsx", "Sheet1", "A1:C3", style)

// 应用到单格（转调 SetCellStyleRange）
idx, _ = excelgo.SetCellStyle("wb.xlsx", "Sheet1", "A1", style)
```

选项结构：

- `FontStyle{Bold, Italic, Underline("single|double|..."), Strike bool; Size float64; Name, Color string}`
- `FillStyle{Color string; PatternType string}`（`Color` 为 ARGB，如 `FFFFFF00`；省略用 `solid`）
- `BorderStyle{Left, Right, Top, Bottom, Diagonal BorderSide}`、`BorderSide{Style, Color string}`
- `AlignmentStyle{Horizontal, Vertical string; WrapText bool}`
- `Style{Font, Fill, Border, Alignment *...; NumFmt string}`

> 仅设置用到的维度（如只传 `Font`）。未设置的维度沿用样式 0（常规），不会清空已有其他维度。



> 注意：插入/删除行或列时，公式单元格内 `<f>` 的 A1 引用会**自动平移**（手写 tokenizer
> 解析公式，区分函数名、字符串字面量、`'表名'!` 前缀、绝对引用 `$` 与跨表/范围引用），
> 不依赖第三方库，也不会把函数名误判为单元格。规则如下：
> - 相对引用随行/列平移（如插入行后 `A1+A2` → `A2+A3`）；绝对维度（`$A$1`/`$A1`/`A$1`）对应维不变。
> - 删除操作落入被删区间的引用改写为 `#REF!`（如删除第 2~3 列后 `A2+B2+C2` → `A2+#REF!+#REF!`）。
> - 不支持 R1C1 引用（检测到原样保留）；跨表与带空格表名（`'My Sheet'!A1`）均正确识别。
>
> 浮动图片锚点不随行列操作平移，以保留图纸。

### 面向对象 API（file → sheet → cell）

除了上文的全局函数，还提供链式对象模型，避免反复写「文件名」，并支持在同一工作簿上连续多次修改后**一次性保存**（改动都在内存 map 中累积，调用 `Save`/`SaveAs` 才落盘）。对象层与全局函数是同一套底层原语，行为完全一致，都不破坏样式 / 布局 / 数据 / 图片。

#### 生命周期与定位

```go
// 打开已有工作簿（先把整个包读进内存 map）
b, err := excelgo.Open("wb.xlsx")
if err != nil { log.Fatal(err) }
defer b.Save()                       // 写回原路径；或 b.SaveAs("out.xlsx") 另存为新文件

// 列出所有工作表名（按 workbook.xml 顺序）
names := b.SheetNames()

// 取已有工作表；不存在会返回 error
s, err := b.Sheet("Sheet1")

// 新建工作表（同名自动加后缀避让，返回 *WorkSheet）
s2, err := b.NewSheet("Feat")

// 其他工作簿级操作（等价于全局函数）
_ = b.RemoveSheet("旧表")             // 删除工作表（连同 worksheet/drawing/独占 media 清理，共享 media 保留）
_ = b.RenameSheet("旧名", "新名")     // 同名冲突自动加后缀
_ = b.MoveSheet("Sheet2", 1)          // 移动到第 1 个位置（1 基）
```

#### 单元格 Cell

`s.Cell(ref)` 返回一个 `*Cell`（轻量句柄，不直接持有数据），再调用方法读写：

```go
c := s.Cell("A1")
c.SetStr("分包队")                          // 共享字符串
c.Set("文本或数字或布尔")                   // 自动推断类型：string→共享串 / int·float64→数值 / bool→布尔
c.SetFormula("A2*2", "150")                // 公式 + 预计算缓存结果（result 可空，留空则 Excel 打开时重算）
c.SetStyle(excelgo.Style{Font: &excelgo.FontStyle{Bold: true}})  // 设置样式（详见「样式设置」）

// 读取
_ = c.Get()          // 共享字符串解析为真实文本；公式返回结果 <v>，无结果返回 "=公式"
_ = c.GetStr()       // 等价 Get()
_ = c.GetNum()       // 解析为 float64（非数值返回 0，失败不报错）
_ = c.GetFormula()   // 返回公式串（不含 "="），空则返回 ""
_ = c.GetStyle()     // 返回单元格 s 样式索引 int
_ = c.MergeTo("C1")  // 以当前单元格为左上角，合并到 C1
```

> 说明：所有 `Set*` 都会先确保目标单元格所在的 `<row>` 存在并置于正确的 `<row r="行号">` 内，符合 OOXML 规范；原有单元格的 `s` 样式索引会被保留，未写样式的单元格默认为样式 0（常规）。

#### 区域 Range

`s.Range(ref)` 返回一个 `*Range`，用于整块读写与样式 / 合并：

```go
r := s.Range("A1:C10")
grid, _ := r.Get()                 // 二维 [][]string（共享字符串解析、公式取结果，空单元格为 ""）
r.Set([][]interface{}{            // 起点为区域左上角，按数据尺寸自动扩展
    {"姓名", "年龄", "分数"},
    {"张三", 18, 95.5},
    {"李四", 20, nil},            // nil 跳过该单元格（保留原内容/样式）
    {excelgo.CellFormula{Formula: "B2+C2", Result: "113.5"}}, // 公式单元格
})
r.SetStyle(excelgo.Style{Fill: &excelgo.FillStyle{Color: "FFFFFF00", PatternType: "solid"}})
r.Merge()                          // 把整个 ref 区域合并
```

- 传入含 `:` 的区域（如 `"A1:C10"`）会按边界校验；也可以只传单格（如 `"A1"`）作为写入起点。
- 单元格值以 `=` 开头的字符串会被自动按公式写入（等价于 `CellFormula`），但显式用 `excelgo.CellFormula` 可附带 `Result` 预计算值。

#### 工作表级操作（WorkSheet 方法）

对象层把「行列 / 合并 / 冻结 / 列宽 / 分组 / 筛选 / 批注 / 表格 / 条件格式 / 属性 / 保护 / 替换」都挂在 `*WorkSheet` 上，与全局函数一一对应：

```go
s.InsertRows(3, 1)                // 在第 3 行前插入 1 行
s.RemoveRows(3, 1)                // 删除第 3 行起的 1 行
s.InsertCols(2, 1)                // 在第 2 列（B 列）前插入 1 列
s.RemoveCols(2, 1)                // 删除第 2 列起的 1 列
s.MergeCells("A1:C1")             // 合并区域
s.UnmergeCells("A1:C1")           // 取消合并
s.SetColWidth(1, 20)              // 第 1 列宽 20
s.SetColWidthRange(1, 3, 20)      // 第 1~3 列宽 20
w := s.GetColWidth(1)             // 读取第 1 列宽（默认 0）
s.SetRowHeight(1, 24)             // 第 1 行高 24
s.SetRowVisible(2, false)         // 隐藏第 2 行
s.SetColVisible(2, false)         // 隐藏第 2 列
s.FreezePanes("A2")               // 冻结首行（A2=冻首行，B1=冻首列，B2=双冻）
s.AddHyperlink("D1", "https://example.com", "链接")   // 超链接（displayText 可选）
s.AutoFilter("A1:C10")            // 自动筛选
s.AddDataValidationList("A4", []string{"上海章麒","上海环届","上海基强"}, true) // 下拉
s.AddDataValidation("B4", "whole", "between", "1", "100", true)                // 通用数据验证
s.SetConditionalFormat("C2:C10", "expression", "C2>70", 1,
    excelgo.Style{Fill: &excelgo.FillStyle{Color: "FFFF0000", PatternType: "solid"}})
s.AddTable("A1:B10", "MyTable")   // 结构化表格
s.AddComment("A1", "备注文本", "张三")   // 批注（author 可空，默认 "Author"）
_ = s.GetComment("A1")            // 返回批注文本
s.SetProps(excelgo.SheetProps{TabColor: "FFFF0000", ShowGridLines: boolPtr(false), Zoom: 120})
s.Protect("")                     // 工作表保护（空密码则仅限制锁定单元格）
s.GroupRows(2, 5, 1)              // 第 2~5 行分组，大纲级别 1
s.GroupCols(1, 3, 1)              // 第 1~3 列分组，大纲级别 1
n, _ := s.Replace("旧", "新")     // 工作表内查找替换，返回替换次数
```

> **列分组（GroupCols）行为说明**：分组级别写入 `<col>` 的 `outlineLevel` 属性。若目标区间已被更早的 `SetColWidth`/`SetColWidthRange` 覆盖（例如先 `SetColWidthRange(1,3,20)` 产生 `<col min=1 max=3 width=20>`），本库会把重叠的 `<col>` 按「左侧原属性段 + 重叠段（带 outlineLevel）+ 右侧原属性段」拆分，既保留原有列宽，又让分组生效——不会因覆盖区间不匹配而丢失分组（这是早期版本的缺陷，已修复）。

### 合并单元格 / 列宽行高 / 冻结 / 超链接 / 筛选

```go
// 合并（区域）与取消合并
_ = excelgo.MergeCells("wb.xlsx", "Sheet1", "A1:C1")
_ = excelgo.UnmergeCells("wb.xlsx", "Sheet1", "A1:C1")

// 列宽 / 行高（列号、行号均 1 基）；隐藏行/列用 Set*Visible
_ = excelgo.SetColWidth("wb.xlsx", "Sheet1", 1, 18)        // 第 1 列宽 18
_ = excelgo.SetColWidthRange("wb.xlsx", "Sheet1", 1, 3, 18) // 第 1~3 列宽 18
_ = excelgo.GetColWidth("wb.xlsx", "Sheet1", 1)            // 读取第 1 列宽（默认 0）
_ = excelgo.SetRowHeight("wb.xlsx", "Sheet1", 1, 24)       // 第 1 行高 24
_ = excelgo.SetRowVisible("wb.xlsx", "Sheet1", 2, false)   // 隐藏第 2 行
_ = excelgo.SetColVisible("wb.xlsx", "Sheet1", 2, false)   // 隐藏第 2 列

// 冻结窗格：A2=冻首行，B1=冻首列，B2=双冻
_ = excelgo.FreezePanes("wb.xlsx", "Sheet1", "A2")

// 超链接（外部 URL；displayText 可选，写入单元格）
_ = excelgo.AddHyperlink("wb.xlsx", "Sheet1", "D1", "https://example.com", "链接")

// 自动筛选（区域，需包含表头行）
_ = excelgo.AutoFilter("wb.xlsx", "Sheet1", "A1:C10")
```

参数与注意：

- **MergeCells / UnmergeCells**：`ref` 为标准 A1 区域（如 `"A1:C1"`）。合并时若区域已存在完全一致的合并定义则跳过；取消合并会移除匹配的合并区域。两者均只改 `mergeCells` 部件，不影响单元格内容与样式。
- **SetColWidth / SetColWidthRange**：列号为 1 基；宽度单位为 Excel 字符宽。`SetColWidthRange(min, max, width)` 会把与 `[min,max]` 重叠的已有 `<col>` 按属性拆分重写，避免覆盖冲突（详见对象层「列分组」说明）。
- **SetRowHeight**：行号为 1 基；`SetRowVisible(col, visible)` / `SetColVisible(col, visible)` 的 `visible=false` 即隐藏（写 `hidden="1"`）。
- **FreezePanes**：`ref` 为左上角冻结点（`A2`/`B1`/`B2`），写入 `<sheetViews><pane>`，自动清除已有的 pane/sheetView 状态以避免冲突。
- **AddHyperlink**：先写入单元格显示文本（若提供 `displayText`），再写 `<hyperlinks>` 与对应 rels；URL 作为 `r:id` 指向新 relation，不破坏其它单元格。
- **AutoFilter**：在 `ref` 区域写入 `<autoFilter ref="...">`，不清除已有数据时可直接配合「筛选」下拉。

### 数据验证（下拉）/ 条件格式 / 表格 / 批注 / 工作表属性

```go
// 下拉列表（逗号分隔选项）
_ = excelgo.AddDataValidationList("wb.xlsx", "Sheet1", "A4", []string{"上海章麒","上海环届","上海基强"}, true)
// 通用数据验证：type 如 list/whole/decimal/date，op 如 between，formula1/formula2 为条件
_ = excelgo.AddDataValidation("wb.xlsx", "Sheet1", "B4", "whole", "between", "1", "100", true)

// 条件格式（cfType 如 expression/cellIs；formula 为条件；st 为命中样式）
_ = excelgo.SetConditionalFormat("wb.xlsx", "Sheet1", "C2:C10", "expression", "C2>70", 1,
    excelgo.Style{Fill: &excelgo.FillStyle{Color: "FFFF0000", PatternType: "solid"}})

// 结构化表格（ListObject）
_ = excelgo.AddTable("wb.xlsx", "Sheet1", "A1:B10", "MyTable")

// 批注（author 可空，默认 "Author"）
_ = excelgo.AddComment("wb.xlsx", "Sheet1", "A1", "备注文本", "张三")
_ = excelgo.GetComment("wb.xlsx", "Sheet1", "A1")   // 返回批注文本

// 工作表属性：标签色（ARGB，写在 workbook.xml 的 <sheet> 上）、网格线、缩放
_ = excelgo.SetSheetProps("wb.xlsx", "Sheet1", excelgo.SheetProps{
    TabColor: "FFFF0000", ShowGridLines: boolPtr(false), Zoom: 120})
```

参数与注意：

- **AddDataValidationList**：`values` 为字符串切片，内部拼接为逗号分隔串写入 `<dataValidation type="list">`；`allowBlank=true` 允许空值。
- **AddDataValidation**：`type` 支持 `list`/`whole`/`decimal`/`date` 等，`op` 支持 `between`/`notBetween`/`greaterThan` 等；`formula2` 在双值条件（如 between）时必须提供。`allowBlank` 控制是否允许空单元格。
- **SetConditionalFormat**：`cfType` 为 `expression`（公式）或 `cellIs`（单元格值比较），`priority` 为优先级序号（数值越小越优先），命中样式经 `xl/styles.xml` 的 `dxfs` 差分格式写入，自动补全 Content_Types 与样式引用。
- **AddTable**：生成 `tableN.xml` 与 `xl/worksheets/_rels` 关系条目，表名需唯一；区域首行通常作为表头。
- **AddComment / Get.Comment**：批注写入 `xl/comments/commentsN.xml` 并补全 `xl/worksheets/_rels` 关系；`GetComment` 返回该单元格批注文本（无则返回空串）。
- **SetSheetProps**：`TabColor` 为 ARGB（如 `FFFF0000`）；`ShowGridLines` 控制网格线显示；`Zoom` 为显示缩放百分比。`nil` 字段表示不修改该项。

### 工作表保护 / 行列分组 / 查找替换 / 文档属性

```go
// 工作表保护（password 为空则不设密码，仅限制锁定单元格）
_ = excelgo.ProtectSheet("wb.xlsx", "Sheet1", "")

// 行列分组大纲（level 为大纲级别，1 基）
_ = excelgo.GroupRows("wb.xlsx", "Sheet1", 2, 5, 1)
_ = excelgo.GroupCols("wb.xlsx", "Sheet1", 1, 3, 1)

// 工作表内查找替换（覆盖共享字符串、内联字符串与数值结果文本）
n, _ := excelgo.ReplaceText("wb.xlsx", "Sheet1", "旧", "新")

// 文档核心属性（title/subject/creator/keywords/description/lastModifiedBy）
_ = excelgo.SetDocProps("wb.xlsx", map[string]string{"title": "G15日志", "creator": "newbie"})
```

参数与注意：

- **ProtectSheet**：`password` 为空时仅启用「锁定单元格」保护（未显式解锁的单元格不可编辑）；传入非空字符串会按 Excel 的 hash 算法写入 `<sheetProtection password="...">`。`protect` 命令支持同样的语义（CLI 传入的是 hash 后的密码串）。
- **GroupRows / GroupCols**：`r1,r2` 或 `c1,c2` 为起止行/列（1 基），`level` 为大纲级别（1 基，通常 1~7）。行分组写入 `<row outlineLevel="N">`；列分组写入 `<col outlineLevel="N">`（与已有的 `<col>` 重叠时会按属性拆分重写，详见对象层说明）。多次嵌套分组时逐级累加级别即可。
- **ReplaceText**：在共享字符串表、内联字符串 `<is>` 与公式结果 `<v>` 中匹配并替换，返回实际替换次数；注意若 `new` 包含 `old` 子串，链式替换可能产生重复替换，调用方应保证语义正确。
- **SetDocProps**：写入 `docProps/core.xml` 的 `coreProperties`（标题、主题、创建者、关键词、描述、最后修改者）。只写传入的字段，不影响其它已有属性。

> 以上所有写入均遵循「读盘 → 精确改写内存 map → 写回」，不破坏原有样式 / 布局 / 数据 / 图片；
> 新建的从属部件（comments / table / dxfs 差分格式等）都会补全 Content_Types 与关系条目。

## 命令行工具

`cmd/excelgo` 暴露上述全部能力（工作表生命周期、单元格读写、图片、样式、行列、合并、筛选、数据验证、条件格式、表格、批注、分组、替换、文档属性与复制/合并）：

```bash
# 构建
go build -o excelgo ./cmd/excelgo

excelgo list   <文件.xlsx>
excelgo newsheet <文件.xlsx> <新表名>
excelgo delsheet <文件.xlsx> <表名或索引>
excelgo movesheet <文件.xlsx> <表名或索引> <目标位置(1基)>
excelgo renamesheet <文件.xlsx> <旧名> <新名>
excelgo getcell <文件.xlsx> <表> <单元格>
excelgo setcell <文件.xlsx> <表> <单元格> <值> [--type str|num|bool|formula] [--result 预计算]
excelgo addpic <文件.xlsx> <表> <图片> [--cell A1] [--col-off N] [--row-off N] [--scale 1]
excelgo addcellpic <文件.xlsx> <表> <单元格> <图片>
excelgo rows  insert|remove <文件.xlsx> <表> <行号(1基)> <数量>
excelgo cols  insert|remove <文件.xlsx> <表> <列号(1基)> <数量>
excelgo rangeget <文件.xlsx> <表> <区域>           例 A1:C10
excelgo rangeset <文件.xlsx> <表> <区域> <数据>    例 "a,b;1,2"（行以;分、单元格以,分；值以=开头视为公式）
excelgo setstyle <文件.xlsx> <表> <单元格或区域> [样式flags]   例 A1:C3 --bold --fill FFFFFF00 --align-h center --wrap
excelgo mergecells <文件.xlsx> <表> <区域>         合并单元格（如 A1:C3）
excelgo unmerge <文件.xlsx> <表> <区域>            取消合并
excelgo colwidth <文件.xlsx> <表> <列(1基)> <宽度> [--max 末列]
excelgo rowheight <文件.xlsx> <表> <行(1基)> <高度>
excelgo freeze <文件.xlsx> <表> <冻结点>           如 A2(冻首行) B1(冻首列) B2(双冻)
excelgo hyperlink <文件.xlsx> <表> <单元格> <URL> [显示文本]
excelgo autofilter <文件.xlsx> <表> <区域>
excelgo dropdown <文件.xlsx> <表> <区域> <选项CSV>  下拉列表（如 甲,乙,丙）
excelgo validation <文件.xlsx> <表> <区域> <类型> <公式1> [公式2]
excelgo cformat <文件.xlsx> <表> <区域> <公式> [--fill ARGB] [--bold]
excelgo table <文件.xlsx> <表> <区域> <表名>       结构化表格
excelgo comment <文件.xlsx> <表> <单元格> <文本> [作者]
excelgo sheetprops <文件.xlsx> <表> [--tab-color ARGB] [--grid-lines 0|1] [--zoom N]
excelgo protect <文件.xlsx> <表> [密码哈希]
excelgo grouprows <文件.xlsx> <表> <起行> <止行> <级别>
excelgo groupcols <文件.xlsx> <表> <起列> <止列> <级别>
excelgo replace <文件.xlsx> <表> <旧文本> <新文本>
excelgo docprops <文件.xlsx> [--title T] [--author A] [--subject S] [--keywords K]
excelgo copy  <输入.xlsx> <输出.xlsx> <表> [--independent] [--suffix 后缀]
excelgo merge <目标.xlsx> <源.xlsx>:<表> [源.xlsx>:<表> ...] [--suffix 冲突后缀]
```

> 所有带 flag 的子命令（`setstyle`/`sheetprops`/`docprops`/`cformat` 等）均支持 flag 写在
> 任意位置（位置参数之后也生效），解决了 Go 标准库 `flag` 在首个位置参数处停止解析后续 flag 的问题。

#### 子命令与库函数对照

| CLI 子命令 | 对应库函数 | 说明 |
| --- | --- | --- |
| `list` | `GetSheetList` | 列出工作表名 |
| `newsheet` | `NewSheet` | 新建工作表 |
| `delsheet` | `DeleteSheet` | 删除工作表（清理从属部件） |
| `movesheet` | `MoveSheet` | 调整工作表顺序 |
| `renamesheet` | `RenameSheet` | 改名（冲突自动加后缀） |
| `getcell` | `GetCell` | 读单元格 |
| `setcell` | `SetCellValue` / `SetCellStr` / `SetCellInt` / `SetCellNumeric` / `SetCellBool` / `SetCellFormula` | 按 `--type` 写入 |
| `addpic` | `AddPicture` | 浮动图片（支持 `--cell/--col-off/--row-off/--scale`） |
| `addcellpic` | `AddCellPicture` | 单元格内嵌图片 |
| `rows insert\|remove` | `InsertRows` / `RemoveRows` | 行插入/删除（公式引用自动平移） |
| `cols insert\|remove` | `InsertCols` / `RemoveCols` | 列插入/删除（公式引用自动平移） |
| `rangeget` | `GetRange` | 区域读取 |
| `rangeset` | `SetRange` | 区域写入（行以 `;` 分、单元格以 `,` 分、`=` 开头为公式） |
| `setstyle` | `SetCellStyleRange` | 样式（字体/填充/边框/对齐/数字格式） |
| `mergecells` | `MergeCells` | 合并 |
| `unmerge` | `UnmergeCells` | 取消合并 |
| `colwidth` | `SetColWidth` / `GetColWidth` | 列宽（`--max` 指定末列实现范围） |
| `rowheight` | `SetRowHeight` | 行高 |
| `freeze` | `FreezePanes` | 冻结窗格 |
| `hyperlink` | `AddHyperlink` | 超链接 |
| `autofilter` | `AutoFilter` | 自动筛选 |
| `dropdown` | `AddDataValidationList` | 下拉列表 |
| `validation` | `AddDataValidation` | 通用数据验证（`类型` 如 list/whole/decimal/date） |
| `cformat` | `SetConditionalFormat` | 条件格式（命中样式填充/加粗） |
| `table` | `AddTable` | 结构化表格 |
| `comment` | `AddComment` / `GetComment` | 批注读写 |
| `sheetprops` | `SetSheetProps` | 标签色/网格线/缩放 |
| `protect` | `ProtectSheet` | 工作表保护 |
| `grouprows` / `groupcols` | `GroupRows` / `GroupCols` | 行列分组大纲 |
| `replace` | `ReplaceText` | 查找替换（返回替换次数） |
| `docprops` | `SetDocProps` | 文档核心属性 |
| `copy` | `CopySheet` | 工作表复制 |
| `merge` | `MergeWorkbook` | 跨工作簿合并 |

> 对象 API（`Open` → `Book.Sheet` → `WorkSheet.Cell/Range`）与上述全局函数行为完全一致，可先在内存中连续修改再一次性 `Save`/`SaveAs`。

样式 flag（`setstyle`）支持任意位置书写（位置参数之后也生效）：

```bash
# 按区域整块写入（行以 ; 分隔，单元格以 , 分隔；公式以 = 开头）
excelgo rangeset wb.xlsx Sheet1 A1 "姓名,年龄,分数;张三,18,95.5;李四,20,=B2+C2"

# 按区域整块读取（制表符分隔输出）
excelgo rangeget wb.xlsx Sheet1 A1:C10

# 设置样式：字体/填充/对齐/数字格式，flag 可写在位置参数之后
excelgo setstyle wb.xlsx Sheet1 A1:C3 --bold --italic --underline single \
    --size 14 --name 微软雅黑 --color FFFF0000 \
    --fill FFFFFF00 --numfmt 0.00 \
    --align-h center --align-v center --wrap
# 等价写法（flag 前置）：excelgo setstyle --bold --fill FFFFFF00 wb.xlsx Sheet1 A1:C3
```

## 测试

```bash
go test ./...
```

单元测试（`excelgo/*_test.go`、`cmd/excelgo/*_test.go`）覆盖对象 API、公式引用平移、复制/合并、样式去重、合并单元格/列宽/冻结/超链接/筛选、数据验证/条件格式/表格/批注/工作表属性、保护/分组/替换/文档属性等逻辑。

### 测试夹具（零外部依赖）

所有测试夹具已内嵌于 `testfixtures/` 子包（`//go:embed` 打包 `full.xlsx`、`grid.xlsx`、`target.xlsx`、`src1.xlsx`、`src2.xlsx` 及 `ph_red.png`、`ph_blue.png`），并由 `excelgo` 包的 `TestMain` 在测试运行前自动物化到 `../testdata_tmp/`。

- **无需** 安装 Python / openpyxl，也 **无需** 预先准备任何样本文件或运行生成脚本；
- `go test ./...` 开箱即跑绿；`testdata_tmp/` 为测试瞬态产物，已被 `.gitignore` 忽略，不会进入版本库；
- `testfixtures/` 仅含内嵌资源，编译进测试二进制、不污染库本体。

> 说明：`TestCopySheetSharedVsIndependent` 仍会尝试读取仓库根目录的 `新建 XLSX 工作表.xlsx`，缺失时该用例自动跳过（不影响其余测试）。如需覆盖此用例，可把任意含浮动图片与打印区域的工作簿放到仓库根目录同名文件。

## 许可证

MIT
