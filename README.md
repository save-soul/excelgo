# excelgo

> 🇨🇳 中文文档 / Chinese version: [README.zh-CN.md](README.zh-CN.md)

A pure-Go Excel (`.xlsx`) toolkit whose goal is to cover the common scenarios —
worksheet copying, cross-workbook merging, read/write, styling, rows/columns, charts —
**without** depending on Microsoft Excel / WPS.

Implemented so far:

- **Worksheet copy**: copy a worksheet (floating pictures, WPS in-cell embedded
  pictures, charts, comments, print areas, etc.) as-is into a **new worksheet in the
  same workbook**.
- **Cross-workbook merge**: move worksheets from multiple source workbooks in as
  independent worksheets, preserving page layout, pictures, shapes, page breaks, print
  properties, and deduplicating identical styles.

The output is a standards-compliant OOXML `.xlsx` that opens normally in Excel / WPS.

## Features

**Worksheet copy & merge**

- Copy floating pictures (`xl/drawings/drawingN.xml`)
- Copy WPS "picture embedded in cell" inline pictures (`mc:AlternateContent` / `x14:picture`)
- Copy charts, comments, and other relationship-referenced parts
- Copy print areas (`_xlnm.Print_Area`), auto-remapped to the new sheet index and name
- Configurable media (picture) copy strategy: shared vs. independent

**In-place editing (without breaking styles / layout / data / pictures)**

- Worksheet lifecycle: create / delete / move / rename
- Cell and range read/write (automatic type inference, formulas, formula references
  auto-shifted with row/column operations)
- Row/column insert / delete (shifts merged regions, column widths, page breaks, print areas)
- Merge cells, column width / row height / hide, freeze panes
- Hyperlinks, auto filter, data validation (dropdown / generic), conditional formatting, table objects
- Comments, worksheet properties (tab color / gridlines / zoom), sheet protection, cell styles
  (font / fill / border / alignment / number format)
- Row/column outline grouping, find-and-replace within a sheet, document core properties

**Usage style**

- Pure Go: no CGO, no external Office dependency
- Provides both a global-function API and a `file → sheet → cell` chained object API
  (see "Object-Oriented API")

## Installation

```bash
go get github.com/save-soul/excelgo
```

## Using it as a library

```go
package main

import (
	"log"

	"github.com/save-soul/excelgo"
)

func main() {
	// Default: shared media (smaller file)
	err := excelgo.CopySheet("input.xlsx", "output.xlsx", "Sheet1")
	if err != nil {
		log.Fatal(err)
	}

	// Independent media: the copy owns its own picture files, editable/deletable
	// independently without affecting the original sheet
	err = excelgo.CopySheet("input.xlsx", "output2.xlsx", "Sheet1",
		excelgo.WithMedia(excelgo.MediaIndependent))
	if err != nil {
		log.Fatal(err)
	}

	// You can also use a 1-based worksheet index and a custom suffix
	err = excelgo.CopySheet("input.xlsx", "output3.xlsx", "1",
		excelgo.WithSuffix("_copy"))
	if err != nil {
		log.Fatal(err)
	}
}
```

### Media strategy

| Strategy | Constant | Behavior | Use case |
| --- | --- | --- | --- |
| Shared media | `excelgo.MediaShared` | Copy and original point to the same `xl/media/imageN.png` | Smallest size; lightweight semantics for WPS "embedded in cell" |
| Independent media | `excelgo.MediaIndependent` | Copy owns its own picture files; references and relationships are rewritten | Copy needs to be edited/deleted independently, decoupled from the original |

The default strategy is `MediaShared`.

## Command-line tool

The repo ships a CLI (`cmd/excelgo`) with two subcommands: `copy` (worksheet copy)
and `merge` (cross-workbook merge).

```bash
# Build
go build -o excelgo ./cmd/excelgo

# —— copy: worksheet copy ——
# Shared media (default)
excelgo copy input.xlsx output.xlsx Sheet1

# Independent-media copy
excelgo copy --independent input.xlsx output.xlsx Sheet1

# Custom suffix
excelgo copy --suffix _copy input.xlsx output.xlsx Sheet1

# —— merge: cross-workbook merge ——
# Move "数据A" from src1 and "数据B" from src2 into target
# (target must already exist, as the container)
excelgo merge target.xlsx "src1.xlsx:数据A" "src2.xlsx:数据B"

# Suffix appended on naming conflicts (default _merge)
excelgo merge target.xlsx "src1.xlsx:Sheet1" --suffix _merged
```

> Each merge source argument has the form `workbookPath:sheetNameOrIndex` (separated by
> `:`), e.g. `"src1.xlsx:数据A"` or `"src1.xlsx:1"`.

## API

```go
func CopySheet(src, dst, sheetRef string, opts ...Option) error
```

- `src` — source `.xlsx` path
- `dst` — output `.xlsx` path (overwritten)
- `sheetRef` — worksheet name (e.g. `"Sheet1"`) or 1-based numeric index (e.g. `"1"`)
- `opts` — optional configuration:
  - `WithMedia(excelgo.MediaStrategy)` — set the media strategy
  - `WithSuffix(string)` — set the new worksheet name suffix (default `"_copy"`)

### Merge worksheets (cross-workbook)

Move specified worksheets from multiple source workbooks into the **same** target
workbook as **independent worksheets** (the merged sheets remain independent of each
other afterward). Good for consolidating worksheets scattered across multiple files.

```go
func MergeWorkbook(dst string, sources []SourceRef, opts ...MergeOption) error
```

- `dst` — target workbook path (**must already exist**, used as the merge container, overwritten)
- `sources` — list of source worksheets, each with `Workbook` (source file path) and
  `Sheet` (worksheet name or 1-based index)
- `opts` — optional configuration:
  - `WithRename(name string)` — set a uniform name for the moved-in worksheet
    (auto-suffixes on conflict)
  - naming-conflict suffix defaults to `"_merge"`

Properties & guarantees:

- **Style merge-dedup**: when moving sheets across workbooks, compares by
  `cellXfs/fontId/fillId/borderId/numFmtId` semantics, reuses equivalent existing style
  entries in the target, and only appends genuinely missing styles — avoiding style-count
  overflow. The cell `s` index is remapped accordingly, so formats never cross.
- **Parts migrate with the sheet**: floating pictures, shapes (drawing), media, comments,
  charts and other related parts are copied, renumbered, and their relationship references
  rewritten.
- **Page properties preserved**: print area (`_xlnm.Print_Area`, with `localSheetId` and
  sheet name auto-remapped), page breaks (`rowBreaks/colBreaks`), page layout, etc. follow
  the worksheet.

```go
err := excelgo.MergeWorkbook("summary.xlsx", []excelgo.SourceRef{
    {Workbook: "jan.xlsx", Sheet: "Sheet1"},
    {Workbook: "feb.xlsx", Sheet: "data"},
})
```

### Worksheet lifecycle (in-place editing)

All of the following modify the original file directly, changing **only the necessary
parts** (workbook.xml / rels / content types / the relevant worksheet / media / drawing)
and never breaking the content, styles, layout, or pictures of other worksheets.

```go
// Create a worksheet, returning its 1-based index (auto-suffixes on duplicate name)
idx, _ := excelgo.NewSheet("wb.xlsx", "NewSheet")

// Delete a worksheet (cleans up its worksheet / drawing / exclusive media; shared pictures kept)
_ = excelgo.DeleteSheet("wb.xlsx", "Sheet1")        // name or 1-based index
_ = excelgo.DeleteSheet("wb.xlsx", "2")

// Move a worksheet to a position (toIndex is 1-based; 1 = front)
_ = excelgo.MoveSheet("wb.xlsx", "Sheet2", 1)

// Rename (auto-suffixes on duplicate name)
_ = excelgo.RenameSheet("wb.xlsx", "old", "new")
```

### Cell read/write (multiple formats)

Following excelize's fine-grained write style: automatic inference or precise typed
write, preserving the original cell `s` style index.

```go
// Read (shared strings auto-parsed to real text; formulas return result <v>,
// "=formula" returned when no result)
val, _ := excelgo.GetCell("wb.xlsx", "Sheet1", "A1")

// Auto type inference (string→shared string, int/float→number, bool→boolean)
_ = excelgo.SetCellValue("wb.xlsx", "Sheet1", "A1", "text")
_ = excelgo.SetCellValue("wb.xlsx", "Sheet1", "A2", 123)
_ = excelgo.SetCellValue("wb.xlsx", "Sheet1", "A3", true)

// Typed writes
_ = excelgo.SetCellStr("wb.xlsx", "Sheet1", "B1", "shared string")
_ = excelgo.SetCellInt("wb.xlsx", "Sheet1", "B2", 42)
_ = excelgo.SetCellNumeric("wb.xlsx", "Sheet1", "B3", 3.14)
_ = excelgo.SetCellBool("wb.xlsx", "Sheet1", "B4", true)
_ = excelgo.SetCellInline("wb.xlsx", "Sheet1", "B5", "inline string (no sharedStrings)")
_ = excelgo.SetCellFormula("wb.xlsx", "Sheet1", "B6", "A1*2", "precomputed result") // result may be empty
```

> Newly written cells are strictly placed inside the correct `<row r="rowNumber">`,
> conforming to the OOXML spec; the existing cell `s` style index is preserved, and cells
> without a written style default to style 0 (normal).

### Insert pictures

Two modes, neither breaking existing drawing objects on the sheet.

```go
// Floating picture: anchor cell, pixel offset, and scale ratio are configurable
_ = excelgo.AddPicture("wb.xlsx", "Sheet1", "photo.png", &excelgo.PictureOptions{
    ScaleX: 0.5, ScaleY: 0.5,                       // scale ratio, default 1
    Position: &excelgo.PicturePosition{Cell: "H1",  // anchor cell
        ColOffset: 10, RowOffset: 5},               // pixel offset relative to anchor
})
// Byte variant is also available: excelgo.AddPictureFromBytes(..., data, opts)

// WPS "picture embedded in cell": moves and scales with the cell (mc:AlternateContent / x14:picture)
_ = excelgo.AddCellPicture("wb.xlsx", "Sheet2", "B2", "photo.png")
// Byte variant: excelgo.AddCellPictureFromBytes(..., data)
```

> Floating pictures are written to `xl/drawings/` and `xl/media/` and reuse/create drawing
> parts; WPS in-cell pictures are written to the worksheet's `mc:AlternateContent` block
> (with `x14:picture` and a standard drawingML fallback), and the required namespaces
> (mc/x14/a/xdr/r) are completed so both Excel / WPS can recognize them.

### Worksheet query & whole-sheet read

```go
// Worksheet list (in workbook.xml order)
names, _ := excelgo.GetSheetList("wb.xlsx")
// 1-based index (returns 0 and error if not found)
idx, _ := excelgo.GetSheetIndex("wb.xlsx", "Sheet1")
// Whether it exists
exists, _ := excelgo.SheetExists("wb.xlsx", "Sheet1")
// Whole sheet read as a 2D array (shared strings parsed, formulas take results, empty cells "")
rows, _ := excelgo.GetSheetData("wb.xlsx", "Sheet1") // equivalent to GetRows
```

### Row/column insert & delete

Insert or delete rows or columns. Only affected elements are shifted (`<row r>`,
`<c r>`, merged region `mergeCell ref`, column width `col min/max`, page break `brk`,
print area `definedName`); the cell `s` style index and pictures are preserved, and other
data/layout is left intact. Row/column numbers are 1-based.

```go
// Insert 2 rows before row 3
_ = excelgo.InsertRows("wb.xlsx", "Sheet1", 3, 2)
// Delete 1 row starting at row 3
_ = excelgo.RemoveRows("wb.xlsx", "Sheet1", 3, 1)
// Insert 1 column before column 2 (column B)
_ = excelgo.InsertCols("wb.xlsx", "Sheet1", 2, 1)
// Delete 1 column starting at column 2
_ = excelgo.RemoveCols("wb.xlsx", "Sheet1", 2, 1)
```

### Range read/write (by rectangular block)

Read or write a rectangular region (e.g. `"A1:C10"`) as a whole, reusing a single
read / per-cell rewrite / single write-back over the in-memory map, preserving the
original `s` style index and layout.

```go
// Read a region as a 2D string array (shared strings parsed, formulas take results, empty cells "")
grid, _ := excelgo.GetRange("wb.xlsx", "Sheet1", "A1:C10")

// Write a region: a single-cell start (e.g. "A1") auto-expands by data size;
// a region containing ":" is validated against its bounds
_ = excelgo.SetRange("wb.xlsx", "Sheet1", "A1", [][]interface{}{
    {"Name", "Age", "Score"},
    {"Zhang", 18, 95.5},
    {"Li", 20, nil},                                   // nil skips that cell
    {excelgo.CellFormula{Formula: "B2+C2", Result: "113.5"}}, // formula cell
})
```

- `GetRange(filename, sheetRef, rangeRef)` → `([][]string, error)`, accepts a region or a single cell.
- `SetRange(filename, sheetRef, rangeRef, values)` supports `string / int / float64 / bool / nil /
  excelgo.CellFormula`; a string value starting with `=` is automatically written as a formula.
- A formula cell uses `excelgo.CellFormula{Formula: "A1+B1", Result: "result"}`; `Result` may be
  empty (only `<f>` is written, no cached `<v>`, so Excel recomputes on open).

### Style API

Programmatically set a cell's font / fill / border / alignment / number format, writing
to `xl/styles.xml`; `font / fill / border / cellXfs` are all deduplicated by content to
avoid style-table bloat, and the cell `s` index is reused accordingly.

```go
style := excelgo.Style{
    Font:      &excelgo.FontStyle{Bold: true, Color: "FFFF0000", Size: 14, Name: "Microsoft YaHei"},
    Fill:      &excelgo.FillStyle{Color: "FFFFFF00", PatternType: "solid"},
    Border:    &excelgo.BorderStyle{                       // each side / diagonal configurable
        Left:  excelgo.BorderSide{Style: "thin", Color: "FF000000"},
        Right: excelgo.BorderSide{Style: "thin", Color: "FF000000"},
    },
    Alignment: &excelgo.AlignmentStyle{Horizontal: "center", Vertical: "center", WrapText: true},
    NumFmt:    "0.00",                                     // number format code; built-in id reused, custom starts at 164
}

// Apply to a region (recommended); applying the same style again reuses the same cellXf index (dedup)
idx, _ := excelgo.SetCellStyleRange("wb.xlsx", "Sheet1", "A1:C3", style)

// Apply to a single cell (delegates to SetCellStyleRange)
idx, _ = excelgo.SetCellStyle("wb.xlsx", "Sheet1", "A1", style)
```

Option structs:

- `FontStyle{Bold, Italic, Underline("single|double|..."), Strike bool; Size float64; Name, Color string}`
- `FillStyle{Color string; PatternType string}` (`Color` is ARGB, e.g. `FFFFFF00`; `solid` used if omitted)
- `BorderStyle{Left, Right, Top, Bottom, Diagonal BorderSide}`、`BorderSide{Style, Color string}`
- `AlignmentStyle{Horizontal, Vertical string; WrapText bool}`
- `Style{Font, Fill, Border, Alignment *...; NumFmt string}`

> Only set the dimensions you use (e.g. pass only `Font`). Unset dimensions inherit style 0
> (normal) and will not clear other existing dimensions.

> **Note:** When inserting/deleting rows or columns, the A1 reference inside a formula
> cell's `<f>` is **auto-shifted** (a hand-written tokenizer parses the formula,
> distinguishing function names, string literals, `'SheetName'!` prefixes, absolute
> references `$`, and cross-sheet/range references) — no third-party library, and function
> names are never misjudged as cells. Rules:
> - Relative references shift with rows/columns (e.g. after inserting a row `A1+A2` →
>   `A2+A3`); absolute dimensions (`$A$1`/`$A1`/`A$1`) keep that dimension unchanged.
> - References falling into the deleted range become `#REF!` on delete (e.g. after deleting
>   columns 2~3, `A2+B2+C2` → `A2+#REF!+#REF!`).
> - R1C1 references are not supported (left as-is when detected); cross-sheet and
>   space-containing sheet names (`'My Sheet'!A1`) are correctly recognized.
>
> Floating picture anchors do **not** shift with row/column operations, to preserve the drawing.

### Object-oriented API (file → sheet → cell)

Besides the global functions above, a chained object model is provided, so you don't
repeat the "filename" everywhere, and can accumulate multiple edits on the same workbook
and **save once** (`Save`/`SaveAs` flush to disk; edits accumulate in the in-memory map).
The object layer and the global functions share the same underlying primitives and behave
identically, neither breaking styles / layout / data / pictures.

#### Lifecycle & locating

```go
// Open an existing workbook (reads the whole package into the in-memory map)
b, err := excelgo.Open("wb.xlsx")
if err != nil { log.Fatal(err) }
defer b.Save()                       // write back to original path; or b.SaveAs("out.xlsx") to save as new

// List all worksheet names (in workbook.xml order)
names := b.SheetNames()

// Get an existing worksheet; returns error if not found
s, err := b.Sheet("Sheet1")

// Create a worksheet (auto-suffixes on duplicate name, returns *WorkSheet)
s2, err := b.NewSheet("Feat")

// Other workbook-level operations (equivalent to global functions)
_ = b.RemoveSheet("OldSheet")             // delete worksheet (cleans worksheet/drawing/exclusive media; shared media kept)
_ = b.RenameSheet("old", "new")           // auto-suffixes on name conflict
_ = b.MoveSheet("Sheet2", 1)              // move to position 1 (1-based)
```

#### Cell

`s.Cell(ref)` returns a `*Cell` (a lightweight handle, not holding data directly); call
methods to read/write:

```go
c := s.Cell("A1")
c.SetStr("subcontractor")                          // shared string
c.Set("text or number or bool")                   // auto type inference: string→shared / int·float64→number / bool→bool
c.SetFormula("A2*2", "150")                       // formula + precomputed cache result (result may be empty; left empty → Excel recomputes on open)
c.SetStyle(excelgo.Style{Font: &excelgo.FontStyle{Bold: true}})  // set style (see "Style API")

// Read
_ = c.Get()          // shared string parsed to real text; formula returns result <v>, "=formula" if no result
_ = c.GetStr()       // equivalent to Get()
_ = c.GetNum()       // parsed to float64 (non-number → 0, no error on failure)
_ = c.GetFormula()   // returns formula string (without "="), "" if empty
_ = c.GetStyle()     // returns the cell's s style index (int)
_ = c.MergeTo("C1")  // merge from this cell (top-left) to C1
```

> All `Set*` first ensure the target cell's `<row>` exists and is placed in the correct
> `<row r="rowNumber">`, conforming to the OOXML spec; the existing cell `s` style index is
> preserved, and cells without a written style default to style 0 (normal).

#### Range

`s.Range(ref)` returns a `*Range` for block read/write and style / merge:

```go
r := s.Range("A1:C10")
grid, _ := r.Get()                 // 2D [][]string (shared strings parsed, formulas take result, empty cells "")
r.Set([][]interface{}{            // start = region top-left, auto-expands by data size
    {"Name", "Age", "Score"},
    {"Zhang", 18, 95.5},
    {"Li", 20, nil},              // nil skips that cell (keeps original content/style)
    {excelgo.CellFormula{Formula: "B2+C2", Result: "113.5"}}, // formula cell
})
r.SetStyle(excelgo.Style{Fill: &excelgo.FillStyle{Color: "FFFFFF00", PatternType: "solid"}})
r.Merge()                          // merge the whole ref region
```

- A region containing `:` (e.g. `"A1:C10"`) is validated against bounds; a single cell
  (e.g. `"A1"`) can also be passed as the write start.
- A string value starting with `=` is automatically written as a formula (equivalent to
  `CellFormula`), but using `excelgo.CellFormula` explicitly lets you attach the `Result`
  precomputed value.

#### Worksheet-level operations (WorkSheet methods)

The object layer puts "rows/cols / merge / freeze / column width / grouping / filter /
comment / table / conditional format / properties / protection / replace" on `*WorkSheet`,
each corresponding one-to-one with a global function:

```go
s.InsertRows(3, 1)                // insert 1 row before row 3
s.RemoveRows(3, 1)                // delete 1 row starting at row 3
s.InsertCols(2, 1)                // insert 1 column before column 2 (column B)
s.RemoveCols(2, 1)                // delete 1 column starting at column 2
s.MergeCells("A1:C1")             // merge region
s.UnmergeCells("A1:C1")           // unmerge
s.SetColWidth(1, 20)              // column 1 width 20
s.SetColWidthRange(1, 3, 20)      // columns 1~3 width 20
w := s.GetColWidth(1)             // read column 1 width (default 0)
s.SetRowHeight(1, 24)             // row 1 height 24
s.SetRowVisible(2, false)         // hide row 2
s.SetColVisible(2, false)         // hide column 2
s.FreezePanes("A2")               // freeze first row (A2=freeze top row, B1=freeze first col, B2=both)
s.AddHyperlink("D1", "https://example.com", "link")   // hyperlink (displayText optional)
s.AutoFilter("A1:C10")            // auto filter
s.AddDataValidationList("A4", []string{"Zhang", "Li", "Wang"}, true) // dropdown
s.AddDataValidation("B4", "whole", "between", "1", "100", true)      // generic data validation
s.SetConditionalFormat("C2:C10", "expression", "C2>70", 1,
    excelgo.Style{Fill: &excelgo.FillStyle{Color: "FFFF0000", PatternType: "solid"}})
s.AddTable("A1:B10", "MyTable")   // structured table
s.AddComment("A1", "note text", "Zhang")   // comment (author optional, default "Author")
_ = s.GetComment("A1")            // returns comment text
s.SetProps(excelgo.SheetProps{TabColor: "FFFF0000", ShowGridLines: boolPtr(false), Zoom: 120})
s.Protect("")                     // sheet protection (empty password → lock only locked cells)
s.GroupRows(2, 5, 1)              // group rows 2~5 at outline level 1
s.GroupCols(1, 3, 1)              // group columns 1~3 at outline level 1
n, _ := s.Replace("old", "new")   // find-and-replace within sheet, returns count
```

> **GroupCols behavior note**: the grouping level is written to the `<col>` `outlineLevel`
> attribute. If the target range was already covered by an earlier `SetColWidth` /
> `SetColWidthRange` (e.g. `SetColWidthRange(1,3,20)` produces `<col min=1 max=3 width=20>`),
> the library splits the overlapping `<col>` into "left original-property segment + overlapping
> segment (with outlineLevel) + right original-property segment", preserving the original
> column width while making grouping take effect — it won't lose grouping due to a mismatched
> coverage range (an early-version defect, since fixed).

### Merge cells / column width & row height / freeze / hyperlink / filter

```go
// Merge (region) and unmerge
_ = excelgo.MergeCells("wb.xlsx", "Sheet1", "A1:C1")
_ = excelgo.UnmergeCells("wb.xlsx", "Sheet1", "A1:C1")

// Column width / row height (both 1-based); hide row/col with Set*Visible
_ = excelgo.SetColWidth("wb.xlsx", "Sheet1", 1, 18)        // column 1 width 18
_ = excelgo.SetColWidthRange("wb.xlsx", "Sheet1", 1, 3, 18) // columns 1~3 width 18
_ = excelgo.GetColWidth("wb.xlsx", "Sheet1", 1)            // read column 1 width (default 0)
_ = excelgo.SetRowHeight("wb.xlsx", "Sheet1", 1, 24)       // row 1 height 24
_ = excelgo.SetRowVisible("wb.xlsx", "Sheet1", 2, false)   // hide row 2
_ = excelgo.SetColVisible("wb.xlsx", "Sheet1", 2, false)   // hide column 2

// Freeze panes: A2=freeze top row, B1=freeze first col, B2=both
_ = excelgo.FreezePanes("wb.xlsx", "Sheet1", "A2")

// Hyperlink (external URL; displayText optional, written to cell)
_ = excelgo.AddHyperlink("wb.xlsx", "Sheet1", "D1", "https://example.com", "link")

// Auto filter (region, must include the header row)
_ = excelgo.AutoFilter("wb.xlsx", "Sheet1", "A1:C10")
```

Parameters & notes:

- **MergeCells / UnmergeCells**: `ref` is a standard A1 region (e.g. `"A1:C1"`). Merge skips
  if an identical merge definition already exists; unmerge removes the matching region. Both
  modify only the `mergeCells` part, leaving cell content and styles intact.
- **SetColWidth / SetColWidthRange**: column number is 1-based; width unit is Excel
  character width. `SetColWidthRange(min, max, width)` splits and rewrites any existing
  `<col>` overlapping `[min,max]` by attribute, avoiding coverage conflicts (see the object
  layer "column grouping" note).
- **SetRowHeight**: row number is 1-based; `SetRowVisible(col, visible)` /
  `SetColVisible(col, visible)` with `visible=false` hides (writes `hidden="1"`).
- **FreezePanes**: `ref` is the top-left freeze point (`A2`/`B1`/`B2`), written to
  `<sheetViews><pane>`; existing pane/sheetView state is auto-cleared to avoid conflicts.
- **AddHyperlink**: writes the cell display text first (if `displayText` given), then the
  `<hyperlinks>` and corresponding rels; the URL is a new relation pointed to by `r:id`,
  leaving other cells intact.
- **AutoFilter**: writes `<autoFilter ref="...">` over `ref`; works directly with the
  "filter" dropdown when existing data isn't cleared.

### Data validation (dropdown) / conditional format / table / comment / sheet properties

```go
// Dropdown list (comma-separated options)
_ = excelgo.AddDataValidationList("wb.xlsx", "Sheet1", "A4", []string{"Zhang", "Li", "Wang"}, true)
// Generic data validation: type like list/whole/decimal/date, op like between, formula1/formula2 are conditions
_ = excelgo.AddDataValidation("wb.xlsx", "Sheet1", "B4", "whole", "between", "1", "100", true)

// Conditional format (cfType like expression/cellIs; formula is the condition; st is the hit style)
_ = excelgo.SetConditionalFormat("wb.xlsx", "Sheet1", "C2:C10", "expression", "C2>70", 1,
    excelgo.Style{Fill: &excelgo.FillStyle{Color: "FFFF0000", PatternType: "solid"}})

// Structured table (ListObject)
_ = excelgo.AddTable("wb.xlsx", "Sheet1", "A1:B10", "MyTable")

// Comment (author optional, default "Author")
_ = excelgo.AddComment("wb.xlsx", "Sheet1", "A1", "note text", "Zhang")
_ = excelgo.GetComment("wb.xlsx", "Sheet1", "A1")   // returns comment text

// Sheet properties: tab color (ARGB, written on <sheet> in workbook.xml), gridlines, zoom
_ = excelgo.SetSheetProps("wb.xlsx", "Sheet1", excelgo.SheetProps{
    TabColor: "FFFF0000", ShowGridLines: boolPtr(false), Zoom: 120})
```

Parameters & notes:

- **AddDataValidationList**: `values` is a string slice, internally joined into a
  comma-separated string written to `<dataValidation type="list">`; `allowBlank=true`
  allows empty values.
- **AddDataValidation**: `type` supports `list`/`whole`/`decimal`/`date`, etc.; `op`
  supports `between`/`notBetween`/`greaterThan`, etc.; `formula2` is required for
  two-value conditions (e.g. `between`). `allowBlank` controls whether empty cells are
  allowed.
- **SetConditionalFormat**: `cfType` is `expression` (formula) or `cellIs` (cell value
  comparison); `priority` is the priority ordinal (smaller = higher priority); the hit
  style is written via the `dxfs` differential format in `xl/styles.xml`, with Content_Types
  and style references auto-completed.
- **AddTable**: generates `tableN.xml` and the `xl/worksheets/_rels` relationship entry;
  the table name must be unique; the region's first row is usually the header.
- **AddComment / GetComment**: the comment is written to `xl/comments/commentsN.xml` with the
  `xl/worksheets/_rels` relationship completed; `GetComment` returns that cell's comment text
  (empty string if none).
- **SetSheetProps**: `TabColor` is ARGB (e.g. `FFFF0000`); `ShowGridLines` controls gridline
  display; `Zoom` is the display zoom percentage. `nil` fields mean "do not modify".

### Sheet protection / row-column grouping / find-replace / document properties

```go
// Sheet protection (empty password → no password, only lock locked cells)
_ = excelgo.ProtectSheet("wb.xlsx", "Sheet1", "")

// Row/column outline grouping (level is outline level, 1-based)
_ = excelgo.GroupRows("wb.xlsx", "Sheet1", 2, 5, 1)
_ = excelgo.GroupCols("wb.xlsx", "Sheet1", 1, 3, 1)

// Find-and-replace within a sheet (covers shared strings, inline strings, and numeric result text)
n, _ := excelgo.ReplaceText("wb.xlsx", "Sheet1", "old", "new")

// Document core properties (title/subject/creator/keywords/description/lastModifiedBy)
_ = excelgo.SetDocProps("wb.xlsx", map[string]string{"title": "G15 log", "creator": "newbie"})
```

Parameters & notes:

- **ProtectSheet**: when `password` is empty, only "locked cells" protection is enabled
  (cells not explicitly unlocked become uneditable); a non-empty string writes
  `<sheetProtection password="...">` using Excel's hash algorithm. The `protect` CLI command
  has the same semantics (the CLI takes an already-hashed password string).
- **GroupRows / GroupCols**: `r1,r2` or `c1,c2` are start/end row/col (1-based), `level` is
  the outline level (1-based, usually 1~7). Row grouping writes `<row outlineLevel="N">`;
  column grouping writes `<col outlineLevel="N">` (overlapping `<col>` are split/rewritten by
  attribute, see the object-layer note). Nested grouping accumulates the level step by step.
- **ReplaceText**: matches and replaces within the shared-string table, inline string `<is>`,
  and formula result `<v>`, returning the actual replacement count; note that if `new`
  contains the `old` substring, chained replacement may double-replace, so callers should
  ensure correct semantics.
- **SetDocProps**: writes the `coreProperties` in `docProps/core.xml` (title, subject,
  creator, keywords, description, lastModifiedBy). Only the passed fields are written; other
  existing properties are unaffected.

> All writes above follow "read → precisely rewrite the in-memory map → write back", never
> breaking existing styles / layout / data / pictures; newly created dependent parts
> (comments / table / dxfs differential format, etc.) all complete their Content_Types and
> relationship entries.

## Command-line tool

`cmd/excelgo` exposes all the capabilities above (worksheet lifecycle, cell read/write,
pictures, styles, rows/cols, merge, filter, data validation, conditional format, table,
comment, grouping, replace, document properties, plus copy/merge):

```bash
# Build
go build -o excelgo ./cmd/excelgo

excelgo list   <file.xlsx>
excelgo newsheet <file.xlsx> <newSheetName>
excelgo delsheet <file.xlsx> <sheetNameOrIndex>
excelgo movesheet <file.xlsx> <sheetNameOrIndex> <targetPos(1-based)>
excelgo renamesheet <file.xlsx> <oldName> <newName>
excelgo getcell <file.xlsx> <sheet> <cell>
excelgo setcell <file.xlsx> <sheet> <cell> <value> [--type str|num|bool|formula] [--result precomputed]
excelgo addpic <file.xlsx> <sheet> <image> [--cell A1] [--col-off N] [--row-off N] [--scale 1]
excelgo addcellpic <file.xlsx> <sheet> <cell> <image>
excelgo rows  insert|remove <file.xlsx> <sheet> <rowNum(1-based)> <count>
excelgo cols  insert|remove <file.xlsx> <sheet> <colNum(1-based)> <count>
excelgo rangeget <file.xlsx> <sheet> <region>          e.g. A1:C10
excelgo rangeset <file.xlsx> <sheet> <region> <data>   e.g. "a,b;1,2" (rows split by ;, cells by ,; value starting with = is a formula)
excelgo setstyle <file.xlsx> <sheet> <cellOrRegion> [style flags]   e.g. A1:C3 --bold --fill FFFFFF00 --align-h center --wrap
excelgo mergecells <file.xlsx> <sheet> <region>        merge cells (e.g. A1:C3)
excelgo unmerge <file.xlsx> <sheet> <region>           unmerge
excelgo colwidth <file.xlsx> <sheet> <col(1-based)> <width> [--max lastCol]
excelgo rowheight <file.xlsx> <sheet> <row(1-based)> <height>
excelgo freeze <file.xlsx> <sheet> <freezePoint>       e.g. A2(freeze top row) B1(freeze first col) B2(both)
excelgo hyperlink <file.xlsx> <sheet> <cell> <URL> [displayText]
excelgo autofilter <file.xlsx> <sheet> <region>
excelgo dropdown <file.xlsx> <sheet> <region> <optionsCSV>   dropdown list (e.g. a,b,c)
excelgo validation <file.xlsx> <sheet> <region> <type> <formula1> [formula2]
excelgo cformat <file.xlsx> <sheet> <region> <formula> [--fill ARGB] [--bold]
excelgo table <file.xlsx> <sheet> <region> <tableName>        structured table
excelgo comment <file.xlsx> <sheet> <cell> <text> [author]
excelgo sheetprops <file.xlsx> <sheet> [--tab-color ARGB] [--grid-lines 0|1] [--zoom N]
excelgo protect <file.xlsx> <sheet> [passwordHash]
excelgo grouprows <file.xlsx> <sheet> <startRow> <endRow> <level>
excelgo groupcols <file.xlsx> <sheet> <startCol> <endCol> <level>
excelgo replace <file.xlsx> <sheet> <oldText> <newText>
excelgo docprops <file.xlsx> [--title T] [--author A] [--subject S] [--keywords K]
excelgo copy  <input.xlsx> <output.xlsx> <sheet> [--independent] [--suffix suffix]
excelgo merge <target.xlsx> <src.xlsx>:<sheet> [src.xlsx>:<sheet> ...] [--suffix conflictSuffix]
```

> All flag-bearing subcommands (`setstyle`/`sheetprops`/`docprops`/`cformat`, etc.) accept
> flags anywhere (they also take effect after positional arguments), fixing the standard-lib
> `flag` issue where parsing stops after the first positional argument.

#### Subcommand ↔ library function mapping

| CLI subcommand | Library function | Description |
| --- | --- | --- |
| `list` | `GetSheetList` | list worksheet names |
| `newsheet` | `NewSheet` | create worksheet |
| `delsheet` | `DeleteSheet` | delete worksheet (cleans dependent parts) |
| `movesheet` | `MoveSheet` | reorder worksheet |
| `renamesheet` | `RenameSheet` | rename (auto-suffix on conflict) |
| `getcell` | `GetCell` | read cell |
| `setcell` | `SetCellValue` / `SetCellStr` / `SetCellInt` / `SetCellNumeric` / `SetCellBool` / `SetCellFormula` | write by `--type` |
| `addpic` | `AddPicture` | floating picture (supports `--cell/--col-off/--row-off/--scale`) |
| `addcellpic` | `AddCellPicture` | in-cell embedded picture |
| `rows insert\|remove` | `InsertRows` / `RemoveRows` | row insert/delete (formula refs auto-shift) |
| `cols insert\|remove` | `InsertCols` / `RemoveCols` | column insert/delete (formula refs auto-shift) |
| `rangeget` | `GetRange` | region read |
| `rangeset` | `SetRange` | region write (rows by `;`; cells by `,`; `=` prefix = formula) |
| `setstyle` | `SetCellStyleRange` | style (font/fill/border/alignment/number format) |
| `mergecells` | `MergeCells` | merge |
| `unmerge` | `UnmergeCells` | unmerge |
| `colwidth` | `SetColWidth` / `GetColWidth` | column width (`--max` sets last col for a range) |
| `rowheight` | `SetRowHeight` | row height |
| `freeze` | `FreezePanes` | freeze panes |
| `hyperlink` | `AddHyperlink` | hyperlink |
| `autofilter` | `AutoFilter` | auto filter |
| `dropdown` | `AddDataValidationList` | dropdown list |
| `validation` | `AddDataValidation` | generic data validation (`type` like list/whole/decimal/date) |
| `cformat` | `SetConditionalFormat` | conditional format (hit fill/bold) |
| `table` | `AddTable` | structured table |
| `comment` | `AddComment` / `GetComment` | comment read/write |
| `sheetprops` | `SetSheetProps` | tab color/gridlines/zoom |
| `protect` | `ProtectSheet` | sheet protection |
| `grouprows` / `groupcols` | `GroupRows` / `GroupCols` | row/column outline grouping |
| `replace` | `ReplaceText` | find-and-replace (returns count) |
| `docprops` | `SetDocProps` | document core properties |
| `copy` | `CopySheet` | worksheet copy |
| `merge` | `MergeWorkbook` | cross-workbook merge |

> The object API (`Open` → `Book.Sheet` → `WorkSheet.Cell/Range`) behaves exactly like the
> global functions above; you can accumulate edits in memory and flush once with
> `Save`/`SaveAs`.

The style flags for `setstyle` accept being written anywhere (also after positional args):

```bash
# Write a whole region as a block (rows split by ;, cells by ,; formula starts with =)
excelgo rangeset wb.xlsx Sheet1 A1 "Name,Age,Score;Zhang,18,95.5;Li,20,=B2+C2"

# Read a whole region as a block (tab-separated output)
excelgo rangeget wb.xlsx Sheet1 A1:C10

# Set style: font/fill/alignment/number format; flags may follow positional args
excelgo setstyle wb.xlsx Sheet1 A1:C3 --bold --italic --underline single \
    --size 14 --name "Microsoft YaHei" --color FFFF0000 \
    --fill FFFFFF00 --numfmt 0.00 \
    --align-h center --align-v center --wrap
# Equivalent (flags first): excelgo setstyle --bold --fill FFFFFF00 wb.xlsx Sheet1 A1:C3
```

## Testing

```bash
go test ./...
```

Unit tests (`excelgo/*_test.go`, `cmd/excelgo/*_test.go`) cover the object API, formula
reference shifting, copy/merge, style dedup, merge cells / column width / freeze / hyperlink
/ filter, data validation / conditional format / table / comment / sheet properties,
protection / grouping / replace / document properties, and more.

### Test fixtures (zero external dependency)

All test fixtures are embedded in the `testfixtures/` subpackage (`//go:embed` packages
`full.xlsx`, `grid.xlsx`, `target.xlsx`, `src1.xlsx`, `src2.xlsx`, and `ph_red.png`,
`ph_blue.png`), and materialized to `../testdata_tmp/` automatically by the `excelgo`
package's `TestMain` before tests run.

- **No** need to install Python / openpyxl, nor to prepare any sample files or run a
  generation script in advance;
- `go test ./...` runs green out of the box; `testdata_tmp/` is a transient test artifact,
  ignored by `.gitignore`, and never enters the repo;
- `testfixtures/` contains only embedded resources, compiled into the test binary without
  polluting the library itself.

> Note: `TestCopySheetSharedVsIndependent` still attempts to read `新建 XLSX 工作表.xlsx` at
> the repo root; when missing, that case auto-skips (doesn't affect the rest). To cover it,
> place any workbook with floating pictures and a print area at that same name in the repo root.

## License

MIT
