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
	// 打开工作簿（返回 *excelgo.Book，别名 *excelgo.File）
	f, err := excelgo.Open("./input.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// 把工作表复制到另一个文件（跨工作簿）；默认共享媒体
	if err := f.CopySheetTo("output.xlsx", "Sheet1", ""); err != nil {
		log.Fatal(err)
	}

	// 独立媒体：复制体拥有自己的图片副本，可单独编辑/删除而不影响原表
	if err := f.CopySheetTo("output2.xlsx", "Sheet1", "",
		excelgo.WithMedia(excelgo.MediaIndependent)); err != nil {
		log.Fatal(err)
	}

	// 也可在同一工作簿内用 1 基索引与自定义后缀复制，最后统一 Save
	if _, err := f.CopySheet("1", "", excelgo.WithSuffix("_副本")); err != nil {
		log.Fatal(err)
	}
	if err := f.Save(); err != nil {
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

### 对象式 API（类 excelize）

`Open` / `Create` 返回 `*Book`（别名 `*File`）句柄，修改先缓存在内存，直到调用 `Save` / `SaveAs` 才一次性写盘——与 excelize 的用法一致。工作表复制（`CopySheet` / `CopySheetTo`）与跨工作簿合并（`Merge` / `MergeWorkbook`）既提供 `*Book` 方法，也提供等价的包级便捷函数（`CopySheet` / `MergeWorkbook`），可按场景选择（包级函数适合「单文件一次性操作」，对象式适合「同一工作簿连续多步操作后一次性保存」）。

```go
package main

import (
	"log"

	"github.com/save-soul/excelgo"
)

func main() {
	// 打开工作簿（返回 *excelgo.Book，别名 *excelgo.File）
	f, err := excelgo.Open("./钢筋.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// 在同一工作簿内复制工作表（类 excelize 语义）；返回新表名
	newName, err := f.CopySheet("Sheet1", "")
	if err != nil {
		log.Fatal(err)
	}
	_ = newName

	// 把工作表复制到另一个文件（跨工作簿）
	if err := f.CopySheetTo("output.xlsx", "Sheet1", "", excelgo.WithMedia(excelgo.MediaIndependent)); err != nil {
		log.Fatal(err)
	}

	// 把其它工作簿的工作表合并进本工作簿
	if err := f.Merge([]excelgo.SourceRef{
		{Workbook: "1月.xlsx", Sheet: "数据A"},
		{Workbook: "2月.xlsx", Sheet: "数据B"},
	}); err != nil {
		log.Fatal(err)
	}

	// 一次性写盘（也可用 f.SaveAs("new.xlsx") 另存为新文件）
	if err := f.Save(); err != nil {
		log.Fatal(err)
	}
}
```

由于修改缓存在内存，同一个 `f` 可以先做多次编辑（复制 / 合并 / 读写字单元格），最后只 `Save` 一次。

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

# 直接指定新表名（--name）
excelgo copy input.xlsx output.xlsx Sheet1 --name 汇总表

# flags 可放在位置参数之前或之后，两者等价
excelgo copy --name 汇总表 input.xlsx output.xlsx Sheet1

# —— merge：跨工作簿合并 ——
# 将 src1 的「数据A」表、src2 的「数据B」表搬入 target（target 需已存在，作为容器）
excelgo merge target.xlsx "src1.xlsx:数据A" "src2.xlsx:数据B"

# 命名冲突时追加的后缀（默认 _merge）
excelgo merge target.xlsx "src1.xlsx:Sheet1" --suffix _合
```

> merge 的每个源参数格式为 `工作簿路径:工作表名或索引`（用 `:` 分隔），例如 `"src1.xlsx:数据A"` 或 `"src1.xlsx:1"`。

## API

```go
// 包级便捷函数（一次性打开→复制→另存，适合单文件操作）：
func CopySheet(src, dst, sheetRef, newName string, opts ...Option) error

// 对象式 API（推荐；可在同一工作簿连续多步操作后一次性 Save）：
func (b *Book) CopySheet(sheetRef, newName string, opts ...Option) (string, error)    // 同一工作簿内复制
func (b *Book) CopySheetTo(dstFile, sheetRef, newName string, opts ...Option) error   // 复制到另一个文件
```

- `sheetRef` 工作表名称（如 `"Sheet1"`）或 1 基数字索引（如 `"1"`）
- `newName` 复制后工作表名。**传空串 `""` 则用「源表名 + `WithSuffix` 后缀」**（默认 `_copy`）；
  非空则直接用该名字，重名时自动补 `_1`/`_2` 序号避让
- `dstFile` `CopySheetTo` 的目标文件路径；为空或与当前文件名相同则仅改内存/写回原文件
- `opts` 可选配置：
  - `WithMedia(excelgo.MediaStrategy)` 设置媒体策略
  - `WithSuffix(string)` 设置新工作表名后缀（默认 `"_copy"`，仅在 `newName` 为空时生效）

```go
// 最常见：直接指定新表名
excelgo.CopySheet("src.xlsx", "dst.xlsx", "Sheet1", "汇总表")

// 不关心名字，沿用默认后缀 Sheet1_copy
excelgo.CopySheet("src.xlsx", "dst.xlsx", "Sheet1", "")
```

> **跨文件复制的语义**：包级 `CopySheet(src, dst, ...)` 与 `(*Book).CopySheetTo(dst, ...)`
> 只搬运**被指定的那一张（或多张）工作表**到目标文件 —— 源工作簿的其它工作表一律不参与，
> 目标工作簿原有工作表全部保留，源工作簿也不会被修改。它不是「整簿合并」，
> 也不覆盖目标文件。实现上直接在两个内存 map 之间搬运该表的部件，不经过临时文件。
>
> 为保证贴入后格式/图片/打印区域正确，跨簿时会做必要的适配：合并 `styles.xml`
> 并重映射 `s` 索引、共享字符串内联化、搬移 drawing/媒体/批注/图表等关联部件
> 并重映射 `rId`、复制打印区域并重映射表名。
>
> 目标文件不存在时会先创建一个最小空工作簿再写入；若目标已存在则在其基础上追加。
> 若目标文件**存在但无法解析**（损坏 / 加密 / 非 xlsx / 路径写错），则直接返回错误
> 并**保持原文件不变** —— 绝不会静默用空工作簿覆盖你的数据。
> 表名默认「源表名 + `WithSuffix` 后缀」（默认 `_copy`）；传第 4 个位置参数
> `newName` 可直接指定（`excelgo.CopySheet(src, dst, "Sheet1", "汇总表")`），
> 或用 `WithSuffix` 自定义后缀；重名时自动补数字序号避让。

### 合并工作表（跨工作簿）

把多个源工作簿中的指定工作表，作为**独立工作表**搬移到同一个目标工作簿中
（合并后各表仍彼此独立）。适合把分散在多个文件里的工作表整合进一个文件。

```go
// 包级便捷函数（一次性打开→合并→保存，dst 为目标容器路径，必须已存在）：
func MergeWorkbook(dst string, sources []SourceRef, opts ...MergeOption) error

// 对象式 API（推荐；可在同一工作簿连续多步操作后一次性 Save）：
func (b *Book) Merge(sources []SourceRef, opts ...MergeOption) error            // 合并进本工作簿
```

- 合并的目标容器即当前 `Open` 得到的工作簿 `f` 自身（无需再传目标路径）；源工作表由 `sources` 指定从哪些文件搬入。
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

> 重复打印行（`_xlnm.Print_Titles`）也会一并复制到新表，表名与 `localSheetId` 同步重映射。

```go
f, err := excelgo.Open("汇总.xlsx")
if err != nil {
    log.Fatal(err)
}
if err := f.Merge([]excelgo.SourceRef{
    {Workbook: "1月.xlsx", Sheet: "Sheet1"},
    {Workbook: "2月.xlsx", Sheet: "数据"},
}); err != nil {
    log.Fatal(err)
}
if err := f.Save(); err != nil {
    log.Fatal(err)
}
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

// 改名（同名自动加序号避让，如 新名_1）
_ = excelgo.RenameSheet("wb.xlsx", "旧名", "新名")
```

> **表名限制（遵循 Excel 规则）**：新建 / 改名 / 复制时，表名最长 **31 个字符**，
> 且不能含 `\ / ? * [ ] :`（`AddSheet` / `NewSheet` / `RenameSheet` / `CopySheet` 均会校验）。
> 表名本身**允许**含 `& < > "` 等 XML 特殊字符 —— 写入 `workbook.xml` 时会自动转义，
> 复制/合并/重命名均能正确处理。

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

> **表尾边界校验**：插入位置允许等于「表尾 + 1」（在表尾之后追加，符合 Excel 行为），
> 超过则返回错误；删除范围超出表尾同样报错。避免误传大数值在文件里生成大片无意义空行/空列。
> 例：`InsertRows(wb, "Sheet1", 9999, 1)` 会返回
> `插入行 9999 超出表尾（当前最大行 N，最多可在 N+1 处插入）`。
> 表尾位置按实际内容判定（`<row r>`、`<c r>`、`<col min/max>`、`<row spans>`），
> 包级函数与 `(*WorkSheet)` 方法行为一致。

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
> 浮动图片锚点（`oneCellAnchor` / `twoCellAnchor`）与 Excel 逻辑一致地随行列操作平移：
> - 插入行/列时，锚定在该位置**及其之后**的图片按行列同步偏移，`colOff`/`rowOff` 像素偏移保持不变；
> - 删除行/列时同样前移；若图片锚定的 `from` 落在被删区间内，则该图片随内容一并移除。
> - 兼容 `xdr:` 前缀与默认命名空间两种 drawing 写法（Excel/WPS 与本库写出的格式均支持）。

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
excelgo copy  <输入.xlsx> <输出.xlsx> <表> [--name 新表名] [--independent] [--suffix 后缀]
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
| `renamesheet` | `RenameSheet` | 改名（冲突自动加序号后缀） |
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
| `copy` | `CopySheet` / `(*Book).CopySheetTo` | 工作表复制（包级便捷 / 对象式） |
| `merge` | `MergeWorkbook` / `(*Book).Merge` | 跨工作簿合并（包级便捷 / 对象式） |

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

### 差分测试（openpyxl 作为外部 oracle）

单元测试有个结构性盲区：它们验证的是"本库输出符合本库的预期"。**若预期本身写错了，就查不出来。**

`difftest/` 把 openpyxl（独立实现、被 Excel 生态广泛验证）当作外部 oracle：同一组文档、同一组方法，两侧分别执行，再把产物归一化为**语义快照**逐字段比对。它能抓出两类既有测试抓不到的问题：

- 本库产出了别的工具读不出、或读成别的值的东西（真 bug）
- 本库读不出 openpyxl 的正常产物（兼容性缺口）

| 文件 | 作用 |
| --- | --- |
| `difftest/snapshot.py` | 把 xlsx 归一化为可比 JSON（语义层，不比 XML 字节） |
| `difftest/op.py` | 按 JSON 指令用 openpyxl 执行同一组操作 |
| `difftest_test.go` | 差分框架：驱动两侧、比对、按字段路径豁免 |
| `difftest_write_test.go` | 写侧 34 个场景（值/公式/样式/行列/布局/工作表管理） |
| `difftest_read_test.go` | 读侧 12 个场景（openpyxl 写 → excelgo 读）+ 往返一致性 |
| `difftest_composite_test.go` | 复合操作：`CopySheet` / `CopySheetTo` / `MergeWorkbook` 的版式保真 |
| `difftest_picture_test.go` | 图片：WPS 内嵌图片、浮动图片、媒体策略、`ExtractPicture` |
| `difftest_style_test.go` | 样式枚举细节：下划线/边框线型/对齐/数字格式 + 非法值拦截 |
| `difftest_chart_test.go` | 图表：**多级关系链**（sheet→drawing→chart）的搬运与闭合 |
| `difftest_protect_test.go` | 工作表/工作簿保护：哈希算法与 Excel 兼容性、复制/合并后是否保留 |
| `difftest_combo_test.go` | **组合场景（搬运既有文件）**：6 类 openpyxl 生成的组合 × 3 种操作 |
| `difftest_selfcombo_test.go` | **组合场景（本库自产）**：图表创建 + 命名样式 + 表格 + 数组公式 + 条件格式 + 数据验证 + 折叠 + 合并单元格全部叠在一张表上，含"副本再复制一次"的二次组合 |
| `difftest_readsem_test.go` | 读侧语义：值/类型、公式、共享字符串、边界坐标、日期语义差异 |

复合操作不能"两侧执行同一操作再比对" —— openpyxl 没有 `MergeWorkbook`，
`copy_worksheet` 也只是浅复制（不搬图片/条件格式/数据验证）。那里 openpyxl
转为**结果验证器**：由 excelgo 执行复合操作，再用 openpyxl 独立打开产物，
检查语义是否一致 + **关联部件是否跟着搬走**（drawing / media / tables /
comments）—— 后者才是"版式保真"的核心，且普通语义快照完全看不到。

```bash
go test -run TestDiff ./...        # 约 2.5 分钟
```

未安装 openpyxl 时相关测试自动跳过（不阻塞 CI）。可用 `EXCELGO_PYTHON` 环境变量指定解释器。

### 它抓到的真实 bug

差分测试的价值在于：这些缺陷**单元测试全绿也发现不了**，因为它们要靠外部实现
才能暴露。

| 缺陷 | 症状 | 根因 |
| --- | --- | --- |
| 列宽范围写入不互操作 | 同一份文件，excelgo 读出 3 列都是 12，openpyxl 只读出 D 列 | 原写 `<col min="4" max="6">`（规范合法但消费方按稀疏字典处理）。openpyxl 自己逐列展开，本库改为逐列 |
| 复制带表格的表 → 整个文件打不开 | openpyxl 抛 `Table with name X already exists` | 部件搬运是字节级的，副本带着相同的 `id` 与 `name`/`displayName`，而两者都是**工作簿级唯一**的 |
| 条件格式类型写错 → 整个文件打不开 | openpyxl 抛 `Unable to read workbook` | `cfType` 原样写入 XML。库注释写的是 `cellIs`，但传 `cell`（直觉命名）就产出坏文件，且**无任何校验** |
| 数据验证类型/运算符同上 | 同上 | `typ`/`op` 同样原样写入、无校验 |
| 填充图案 / 下划线 / 边框线型 / 对齐 | 同上 | 同一类缺陷：枚举值被当自由文本处理 |
| 含图表的工作表被复制/合并 → 图表全丢、文件损坏 | openpyxl 抛 `There is no item named 'xl/charts/chart1.xml'` | 部件搬运**只处理一层 rels**。关系链是多级的（sheet→drawing→chart），drawing 的 rels 被原样搬完就不管了，chart 从未被复制 |
| 合并带条件格式的跨簿文件 → 文件损坏 | openpyxl 抛 `IndexError: list index out of range`（读 `differential_styles[dxfId]`） | `dxfId` 是**工作簿级索引**。早期只搬 `cellXf` 不搬 `<dxfs>`，sheet 里的 `dxfId` 变成悬空引用 |
| 合并带图表/表格的跨簿文件 → 文件损坏 | openpyxl 抛 `Unknown relationship: rId1` | **双写冲突**：关联部件搬运内部重写了 sheet 的 rId 并写回，调用方随后又用自己的副本覆盖，把 rId 重写抹掉了 |
| `CopySheetTo` 跨簿搬运带条件格式的文件 → 文件损坏 | openpyxl 抛 `IndexError: list index out of range` | 与上一条同类但**路径不同**（`copySheetAcrossMaps` 而非 `mergeOneSheet`）。修一处漏一处，故抽成 `mergeDxfsAndRemapSheet` 供两条路径共用 |
| 跨簿搬运命名样式 → 文件损坏 | openpyxl 抛 `TypeError: expected <class 'int'>` | `<cellStyle xfId>` 指向的是 **cellStyleXfs**（不是 cellXfs），且那里的 xf 又引用 fontId/fillId。只搬 `<cellStyles>` 清单而不搬整条链，xfId 必然悬空。现由 `StylesMerger.mergeStyleXfs` 统一处理 |
| Integer 属性写空串 → 文件损坏 | openpyxl 抛 `TypeError: expected <class 'int'>` | `buildXF` 把源 xf 缺失的 `xfId` 写成 `xfId=""`。**Integer 类型属性无值时必须整个省略**，不能写空串 |

**一条通用教训**：写入端必须校验 OOXML 枚举取值域。转义挡不住这类问题 ——
它们不是"值不对"，而是"整个文件消费方读不出来"，且**写入返回 nil、文件也能生成**，
只有 Excel/WPS 打开时才报"文件损坏"，用户极难定位到是哪个字段。

守卫方式：`validateStyleEnums` / `validatePatternType` / `validCFTypes` /
`validDVTypes` 等集中校验，非法值返回明确错误并给出正确取值清单
（如传 `cell` 会提示"应为 cellIs"）。

### 与 openpyxl 的已知语义差异

有几处**有意**的语义差异 —— 单独看每个实现都自洽，但迁移代码时会踩。

| 场景 | openpyxl | 本库 | 说明 |
| --- | --- | --- | --- |
| 读日期单元格（数字 + 日期格式） | `cell.value` 直接给 `datetime` | `GetCellValue` 给原始序列号；`Cell.GetTime()` 转 `time.Time` | 本库不丢原始值，转换显式 |
| 写公式 | 不写缓存结果值 | 写 result | 本库可写出"有缓存结果"的公式 |
| 写空字符串 | 往返后变 `None` | 往返后仍是 `""` | openpyxl 的往返损失，本库更符合"空串是有效值" |
| 条件格式类型名 | `cellIs` | 同（并接受别名提示） | 一致 |
| 列宽范围 | 逐列展开 | 逐列展开 | 一致（规范允许合并，但消费方按稀疏字典处理） |

`TestDateSemanticsDivergence` 锁死了日期这一项：若有人改成自动转 `datetime`，
测试会失败并提示同步更新本表。

### 相对 openpyxl 的能力现状

原先列出的短板已**逐项补齐**：

| 能力 | 状态 | API |
| --- | --- | --- |
| 图表创建 | ✅ 12 种类型（bar/line/pie/scatter/area/doughnut/radar/bubble/surface + 3D） | `AddChart` / `(*WorkSheet).AddChart` |
| 数组公式 | ✅ 写 `t="array"` + `ref`，支持 CSE | `SetCellArrayFormula` |
| 条件格式多规则 | ✅ 10 种类型，含 ColorScale / DataBar / IconSet / top10 / aboveAverage | `SetConditionalFormatRules` |
| 移动区域 | ✅ 值+样式搬移，公式引用自动重映射 | `MoveRange` |
| 命名样式 | ✅ 工作簿级 `cellStyles` 模板，可跨表复用 | `NewNamedStyle` / `SetNamedStyle` / `GetNamedStyles` |
| 大纲折叠 | ✅ hidden + collapsed 双写，摘要行位置可配 | `CollapseRows` / `ExpandRows` / `SetOutlineSummary` |
| 只读流式模式 | ✅ 按需读 sheet XML，不载入无关部件 | `OpenReader` / `StreamRows` |
| write_only 流式写 | ✅ 逐行追加输出，内存不随数据量增长 | `NewStreamWriter` / `StreamSheet.Append` |
| 打印机设置二进制 | ✅ 读取/写入/搬运保真，改页面设置不丢 `r:id` | `GetPrinterSettings` / `SetPrinterSettings` |

仍缺失（优先级低，openpyxl 自身支持也有限）：

| 缺失能力 | 说明 |
| --- | --- |
| 数据透视表 | 完全不支持（openpyxl 也只支持有限读写定义） |
| 图表工作表 Chartsheet | 不支持（图表只能嵌在普通工作表里） |
| 打印机设置二进制的**生成** | 本库能读、能搬、能写回，但无法**合成** DEVMODE（它编码的是具体打印机驱动的能力集）。这与 openpyxl 一致 |
| 图表高级特性 | 数据表/趋势线/双轴/组合图未开放 |

**本库的优势方向**（openpyxl 做不到或很难做）：
- `CopySheet` / `MergeWorkbook` 保留版式、图片、图表、批注、条件格式、打印设置
- WPS 单元格内嵌图片（`x14:picture` + `mc:AlternateContent`）的保真搬运
- 默认**共享**媒体，`WithMedia(MediaIndependent)` 可选独立副本
- 零依赖、单二进制、无 Python 运行时要求

### 性能：write_only 流式写

常规写入每写一格都重扫整份 sheet XML，复杂度 **O(单元格数 × 文档长度)**。
实测 3000 行 × 4 列需 **4 分多钟**；流式写入 20000 行只需 **224 毫秒**
（约 3000 倍差距）。

```go
w, _ := excelgo.NewStreamWriter("out.xlsx")
head, _ := w.AddStyle(excelgo.Style{Font: &excelgo.FontStyle{Bold: true}})
w.SetColWidth(1, 3, 18)

ws, _ := w.NewSheet("数据")
ws.Append([]interface{}{excelgo.StreamCell{Value: "名称", Style: head}, "数量"})
for i := 1; i <= 20000; i++ {
    ws.Append([]interface{}{fmt.Sprintf("项目-%d", i), i})
}
w.Save()
```

与 openpyxl 的 `write_only` 相同的约束：

- **必须先声明表名再写行** —— sheet 名与 `sheetN.xml` 的对应关系在
  `workbook.xml` 里定死，写完行再改名会让数据无处安放
- **行必须按序追加** —— `<row>` 是顺序流，随机回填需要重扫全文档
- **样式必须先 `AddStyle` 登记** —— 单元格写出后无法回填样式索引
- 字符串走 `t="inlineStr"` 内联，不做共享去重（去重需回查整张表，
  正是流式要消除的操作）；代价是文件略大

实现上有一处**格式硬约束**值得记住：`archive/zip` 的 Writer 一次只允许一个
打开的条目，`Create` 新条目会关闭前一个。所以多表不能交错写 —— 数据行
先落到临时文件，落盘时再按序输出为 zip 条目。

### 两条经验

**1. 关系链必须递归搬运。** OOXML 的部件关系是多级的
（`sheet → drawing → chart → 嵌入工作簿`），只处理第一层会留下悬空关系。
库现在有 `copyDownstreamParts`（同簿）与 `copyDownstreamPartsCross`（跨簿）负责递归，
并保证下游部件落在**自身的规范目录**（chart 必须在 `xl/charts/`，不能塞进
`xl/drawings/`）。

**2. 关联部件搬运与工作表写回必须是同一个写入点。**
`copySheetPartsToTarget` 要做两件事：搬部件、**重写 sheet 里的 rId 引用**。
早期它自己写 `dstMap`，调用方随后又用外层变量覆盖一次 —— 后者不含rId 重写，
于是 sheet 仍指向源文件的旧 rId，而 rels 已换成新 ID，出现悬空引用。
现在该函数**返回**改写后的内容，由调用方在所有改写（s 索引 / dxfId / 共享字符串 /
rId）完成后一次性写回。`TestPartTransportSingleWritePoint` 守卫这条不变量。

**3. 跨簿搬运要按"工作簿级索引"逐张表清点。** `styles.xml` 里每张表的下标都是
工作簿级的：`s` → cellXfs、`dxfId` → dxfs、`xfId` → cellStyleXfs，
而这些 xf 内部又引用 fontId / fillId / borderId / numFmtId —— **一整条间接链**。
只搬最外层那一张表必然留下悬空引用，而且症状统一是 openpyxl 抛
`IndexError` 或 `TypeError`，看不出是哪张表少了东西。
现由 `StylesMerger`（cellXf + cellStyleXfs）与 `mergeDxfsAndRemapSheet`
（dxfs + cellStyles 清单）分工负责，两条跨簿路径（`Merge` / `CopySheetTo`）共用。

**4. 同一张目标表只能有一个写入点。** dxf 合并与命名样式合并都改
`dstMap["xl/styles.xml"]`，若各自从头写就会互相覆盖 —— 我先写了独立的
`mergeNamedStyles`，结果它把刚合并好的 dxfs 冲掉，openpyxl 又抛 IndexError。
现在 `appendNamedStyles` 是纯函数（收字节、返回字节），由
`mergeDxfsAndRemapSheet` 串起来一次写回。

**5. Integer 属性无值时要整个省略。** `xfId=""` 这类空值属性会让消费方
`_convert("")` 失败抛 `TypeError: expected <class 'int'`，整份文件读不出来。

**6. 密码入口的语义必须显式。** OOXML 的 `password` 属性存的是**哈希值**。
本库的 `ProtectSheet` 历史上要求传已哈希值，而 `ProtectWorkbook` 收明文 ——
两者不一致必然导致误用（传真文会得到一个看似设了密码、实际在 Excel 里永远
解不开的保护）。现新增 `ProtectSheetWithPassword`（收明文，内部哈希）作为
推荐入口，旧入口保留兼容。哈希算法经 openpyxl 独立验证与 Excel 一致
（含中文与特殊字符）。

设计要点：

- **归一化到语义层**。两侧 XML 排布必然不同，比字节无意义；比"用户能观察到的语义"才有意义。
- **区分"真差异"与"能力差异"**。openpyxl 有些能力本库没有，也有本库刻意不同的行为。每个场景可带 `Exempts` 按**字段路径**精确豁免，而不是整场景跳过 —— 前者是真问题，后者是已知取舍。
- **豁免必须写理由**。例如空串：`excelgo` 写共享字符串并能读回 `""`，而 openpyxl 连自己写的空串都读不回来（实测其往返同样为 `None`），这是 openpyxl 的往返损失，本库行为更符合"空串是有效值"的语义。

### 测试夹具（零外部依赖）

所有测试夹具已内嵌于 `testfixtures/` 子包（`//go:embed` 打包 `full.xlsx`、`grid.xlsx`、`target.xlsx`、`src1.xlsx`、`src2.xlsx` 及 `ph_red.png`、`ph_blue.png`），并由 `excelgo` 包的 `TestMain` 在测试运行前自动物化到 `../testdata_tmp/`。

- **无需** 安装 Python / openpyxl，也 **无需** 预先准备任何样本文件或运行生成脚本；
- `go test ./...` 开箱即跑绿；`testdata_tmp/` 为测试瞬态产物，已被 `.gitignore` 忽略，不会进入版本库；
- `testfixtures/` 仅含内嵌资源，编译进测试二进制、不污染库本体。

> 说明：`TestCopySheetSharedVsIndependent` 仍会尝试读取仓库根目录的 `新建 XLSX 工作表.xlsx`，缺失时该用例自动跳过（不影响其余测试）。如需覆盖此用例，可把任意含浮动图片与打印区域的工作簿放到仓库根目录同名文件。

## XML 安全模型（重要）

### 为什么不用 XML 解析器

这不是"偏好字符串"的问题，而是 **`encoding/xml` 往返会破坏 OOXML**。实测数据
（用 `xml.Decoder` + `xml.Encoder` 对 WPS 风格片段做无损往返）：

| 往返前的写法 | 往返后 |
| --- | --- |
| `<mc:AlternateContent>` | `<AlternateContent>`（**前缀丢失**） |
| `<x14:picture>` | `<picture>` |
| `mc:Ignorable="x14ac"` | `x:Ignorable="x14ac"` + 前缀被改名 `_xmlns:x14ac` |

致命的是最后一条：`mc:Ignorable` 的**值是一个前缀名**。Go 的 `encoding/xml`
不保留原始前缀（它把前缀解析成 `Name.Space`，重新编码时自己生成前缀），
于是 `Ignorable` 指向的前缀不再存在 —— **MCE（Markup Compatibility）降级机制
直接失效**，Excel 与 WPS 对该文件的解释随之改变。用结构体 `Unmarshal`+`Marshal`
往返更糟：未在结构体中声明的元素会被整体丢弃（`definedNames`/`bookViews`/`calcPr`
全部消失），且 `r:id` 会被改写成 `rId` 之外的其它前缀。

因此本库**全程以字符串/正则编辑 XML 部件**，这也是保住 WPS 单元格内嵌图片
（`mc:AlternateContent` + `x14:picture`）、命名空间与版式的前提。

### 转义约定与护栏

代价是**每个写入点都必须自行转义**。约定：

- 所有用户数据写入 XML，**一律且只能**经 `safeText`（文本节点）/ `safeAttr`（属性值）出栈；
- 旧名 `escapeXML` / `escapeAttr` 保留为等价别名；
- 用户数据若要进入**正则**，必须用 `regexp.QuoteMeta`。

三层防护：

| 防护 | 说明 |
| --- | --- |
| 单一入口 | `safeText` / `safeAttr` 是唯一允许的转义出口；库内所有写入点已迁移 |
| 控制字符剔除 | XML 1.0 禁止 `U+0000~U+0008 / U+000B / U+000C / U+000E~U+001F`（保留 `\t \n \r`）。这类字符**无法用实体表达**，写入后会让整个部件变成 not-well-formed，Excel 直接判损坏。`safeText` 统一剔除 |
| 长度与字符校验 | 工作表名 ≤ 31 字符、拒绝 `\ / ? * [ ] :`（`AddSheet`/`NewSheet`/`RenameSheet`/`CopySheet`） |

**两道自动防线**（都不依赖人记得）：

1. `TestNoTaintedXMLWrite` —— AST 污点追踪护栏，**零人工豁免**。
   分析每个函数的数据流：函数入参、局部变量、结构体字段、range 变量、`append` 结果
   均视为"带污"，经 `safeText`/`safeAttr`/`strconv.*` 后污点清除；
   带污量被拼进 XML 标签（无论是裸拼接还是作为函数实参）即失败。

   **安全性由类型表达，而非清单**。以下具名类型被视为已认证 —— 因为它们只能经
   带校验的构造函数产生，认证发生在构造处：

   | 类型 | 构造 | 含义 |
   | --- | --- | --- |
   | `cellRef` | `newCellRef` | 已校验的单元格坐标（`A1` 形式） |
   | `rangeRef` | `newRangeRef` | 已校验的区域引用（`A1:C10`） |
   | `relID` | `newRelID` | 库内生成的关系 ID（`rIdN`） |
   | `styleIndex` | `idx` | 样式索引（`fontId`/`numFmtId` 等） |
   | `borderSideName` | — | 固定边框边名常量 |
   | `xmlTagName` | `tagName` | 固定 schema 标签名（`v`/`f`/`r`） |
   | `xmlFrag` | `frag` | 本次新建、内容已 `safe*` 的片段 |
   | `partXML` | `part` | 部件里**既有**的原始 XML（读取后复用） |

   因此**不存在"忘记往清单里加一条"这种漏检** —— 新增写入点默认受保护
   （裸 `string` 一律带污）。`TestGuardNoNameAllowlist` 作为元测试锁死这一约束：
   护栏内若重新出现按变量名的豁免清单，测试直接失败。
2. `TestXMLInjectionResistance` —— 端到端注入回归。6 类载荷
   （闭合标签、闭合加实体、属性引号逃逸、CDATA 伪装、注释伪装、控制字符）× 12 个
   写入点（单元格值/公式、共享字符串、批注文本与作者、超链接 URL 与显示文本、
   命名区域名、文档属性、表名、数据验证、条件格式、页眉页脚/标签颜色），
   断言所有 XML 部件**良构**（`encoding/xml` 严格解码器逐个校验）且未注入多余节点。

> ⚠️ **新增写入点时**：调用 `safeText`（文本节点）或 `safeAttr`（属性值），
> 拼进正则时用 `regexp.QuoteMeta`；若某值确实是"库内生成的安全值"，
> 正确做法是**为它建一个具名类型 + 带校验的构造函数**（如 `cellRef`/`newCellRef`），
> 而**不是**往清单里加一条变量名豁免 —— 那样只是把"靠人记得"换了个地方。

## 许可证

MIT
