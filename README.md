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
	// Open a workbook (returns *excelgo.Book, alias *excelgo.File)
	f, err := excelgo.Open("./input.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Copy a worksheet out to another file (cross-workbook); shared media by default
	if err := f.CopySheetTo("output.xlsx", "Sheet1", ""); err != nil {
		log.Fatal(err)
	}

	// Independent media: the copy owns its own picture files, editable/deletable
	// independently without affecting the original sheet
	if err := f.CopySheetTo("output2.xlsx", "Sheet1", "",
		excelgo.WithMedia(excelgo.MediaIndependent)); err != nil {
		log.Fatal(err)
	}

	// You can also copy within the same workbook using a 1-based worksheet index
	// and a custom suffix; then persist with Save/SaveAs
	if _, err := f.CopySheet("1", "", excelgo.WithSuffix("_copy")); err != nil {
		log.Fatal(err)
	}
	if err := f.Save(); err != nil {
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

### Object-oriented API (excelize-style)

`Open` / `Create` return a `*Book` (alias `*File`) handle, and modifications are buffered in
memory until you call `Save` / `SaveAs` — the same flow as excelize. Worksheet copy
(`CopySheet` / `CopySheetTo`) and cross-workbook merge (`Merge` / `MergeWorkbook`) are
provided **both** as methods on `*Book` **and** as equivalent package-level convenience
functions (`CopySheet` / `MergeWorkbook`): pick the package-level form for a one-shot
open→operate→save on a single file, or the object form when you want to perform several
operations on the same workbook and save once.

```go
package main

import (
	"log"

	"github.com/save-soul/excelgo"
)

func main() {
	// Open a workbook (returns *excelgo.Book, alias *excelgo.File)
	f, err := excelgo.Open("./钢筋.xlsx")
	if err != nil {
		log.Fatal(err)
	}

	// Copy a worksheet within the same workbook (excelize-style); returns the new name
	newName, err := f.CopySheet("Sheet1", "")
	if err != nil {
		log.Fatal(err)
	}
	_ = newName

	// Copy a worksheet out to another file (cross-workbook)
	if err := f.CopySheetTo("output.xlsx", "Sheet1", "", excelgo.WithMedia(excelgo.MediaIndependent)); err != nil {
		log.Fatal(err)
	}

	// Merge worksheets from other workbooks into this one
	if err := f.Merge([]excelgo.SourceRef{
		{Workbook: "jan.xlsx", Sheet: "数据A"},
		{Workbook: "feb.xlsx", Sheet: "数据B"},
	}); err != nil {
		log.Fatal(err)
	}

	// Persist all in-memory changes (or use f.SaveAs("new.xlsx") for a new file)
	if err := f.Save(); err != nil {
		log.Fatal(err)
	}
}
```

Because changes are buffered in memory, a single `f` can be edited many times
(copy / merge / read / write cells) before a single `Save`.

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

# Set the copied sheet name explicitly (--name)
excelgo copy input.xlsx output.xlsx Sheet1 --name Summary

# Flags may appear before or after positional args — both are equivalent
excelgo copy --name Summary input.xlsx output.xlsx Sheet1

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
// Package-level convenience (one-shot open → copy → save, for single-file ops):
func CopySheet(src, dst, sheetRef, newName string, opts ...Option) error

// Object-oriented API (recommended; several ops then one Save):
func (b *Book) CopySheet(sheetRef, newName string, opts ...Option) (string, error)    // copy within the same workbook
func (b *Book) CopySheetTo(dstFile, sheetRef, newName string, opts ...Option) error   // copy out to another file
```

- `sheetRef` — worksheet name (e.g. `"Sheet1"`) or 1-based numeric index (e.g. `"1"`)
- `newName` — name of the copied worksheet. **Pass `""` to use "`<source name>` + `WithSuffix` suffix"**
  (default `_copy`); pass a non-empty string to use it verbatim. Name clashes get a `_1`/`_2` suffix.
- `dstFile` — (`CopySheetTo` only) output `.xlsx` path; empty or same as the current
  file → modify in memory / overwrite the source instead of writing a separate file
- `opts` — optional configuration:
  - `WithMedia(excelgo.MediaStrategy)` — set the media strategy
  - `WithSuffix(string)` — set the new worksheet name suffix (default `"_copy"`, only used when `newName` is empty)

```go
// Most common: name the copied sheet explicitly
excelgo.CopySheet("src.xlsx", "dst.xlsx", "Sheet1", "Summary")

// Don't care about the name → default suffix Sheet1_copy
excelgo.CopySheet("src.xlsx", "dst.xlsx", "Sheet1", "")
```

> **Cross-file semantics**: both `CopySheet(src, dst, ...)` and
> `(*Book).CopySheetTo(dst, ...)` move **only the requested worksheet(s)** into the
> target file — other worksheets of the source workbook are never included, all
> worksheets already in the target are preserved, and the source workbook is left
> untouched. This is *not* a whole-workbook merge, and the target file is never
> overwritten. Parts are relocated directly between two in-memory maps, with no
> temporary file involved.
>
> To keep formatting/pictures/print areas correct after the move, cross-workbook
> adaptation is applied: `styles.xml` is merged and `s` indexes remapped, shared
> strings are inlined, related parts (drawing / media / comments / charts) are
> relocated with `rId` remapping, and the print area is copied with its sheet name
> rewritten.
>
> If the target file does not exist, a minimal empty workbook is created first; if it
> does exist, the sheet is appended to it. If the target **exists but cannot be parsed**
> (corrupt / encrypted / not an xlsx / wrong path), an error is returned and the
> original file is **left untouched** — it is never silently replaced by an empty
> workbook. The new sheet is named
> "`<source name>` + `WithSuffix`" (default `_copy"); pass the 4th positional argument
> `newName` to set it directly (`excelgo.CopySheet(src, dst, "Sheet1", "Summary")`),
> or use `WithSuffix` to customize the suffix. Name clashes get a numeric
> suffix appended.

### Merge worksheets (cross-workbook)

Move specified worksheets from multiple source workbooks into the **same** target
workbook as **independent worksheets** (the merged sheets remain independent of each
other afterward). Good for consolidating worksheets scattered across multiple files.

```go
// Package-level convenience (one-shot open → merge → save; dst is the container, must exist):
func MergeWorkbook(dst string, sources []SourceRef, opts ...MergeOption) error

// Object-oriented API (recommended; several ops then one Save):
func (b *Book) Merge(sources []SourceRef, opts ...MergeOption) error            // merge into this workbook
```

- The merge container is the workbook `f` you already opened with `Open` (no separate
  target path is passed); `sources` specifies which sheets to pull in from other files.
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

> Print titles (`_xlnm.Print_Titles`, i.e. rows/columns repeated on every printed page) are
> carried over to the new sheet as well, with `localSheetId` and the sheet name remapped.

```go
f, err := excelgo.Open("summary.xlsx")
if err != nil {
    log.Fatal(err)
}
if err := f.Merge([]excelgo.SourceRef{
    {Workbook: "jan.xlsx", Sheet: "Sheet1"},
    {Workbook: "feb.xlsx", Sheet: "data"},
}); err != nil {
    log.Fatal(err)
}
if err := f.Save(); err != nil {
    log.Fatal(err)
}
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

// Rename (appends a numeric suffix on name clash, e.g. new_1)
_ = excelgo.RenameSheet("wb.xlsx", "old", "new")
```

> **Sheet name rules (following Excel)**: when creating, renaming or copying, the name may be
> at most **31 characters** and must not contain `\ / ? * [ ] :` (validated by `AddSheet` /
> `NewSheet` / `RenameSheet` / `CopySheet`). Names **may** contain XML-significant characters
> such as `& < > "` — they are escaped automatically when written into `workbook.xml`, and
> copy / merge / rename all handle them correctly.

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

> **Bounds checking**: the insert position may equal `lastRow + 1` (appending past the last
> row, matching Excel's behaviour); anything beyond that returns an error, as does a delete
> range extending past the last row. This avoids silently filling the file with meaningless
> empty rows/columns when a large number is passed by mistake. E.g.
> `InsertRows(wb, "Sheet1", 9999, 1)` returns
> `插入行 9999 超出表尾（当前最大行 N，最多可在 N+1 处插入）`.
> The last row/column is derived from actual content (`<row r>`, `<c r>`, `<col min/max>`,
> `<row spans>`); package-level functions and `(*WorkSheet)` methods behave identically.

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
> Floating picture anchors (`oneCellAnchor` / `twoCellAnchor`) shift with row/column
> operations, matching Excel's own behaviour:
> - On insert, pictures anchored at that position **or after it** shift by the same
>   row/column delta; `colOff`/`rowOff` pixel offsets are preserved.
> - On delete, anchors likewise move up; if a picture's `from` anchor falls inside the
>   deleted band, that picture is removed along with the content.
> - Both the `xdr:`-prefixed form (Excel/WPS) and the default-namespace form (written by
>   this library) are handled.

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

### Workbook protection & defined names

```go
// ---- Workbook structure protection (lock sheet add/delete/reorder) ----
_ = excelgo.ProtectWorkbook("wb.xlsx", "")            // no password
_ = excelgo.ProtectWorkbook("wb.xlsx", "secret")      // with password (Excel verifier hash)
_ = excelgo.UnprotectWorkbook("wb.xlsx")
wp, _ := excelgo.GetWorkbookProtection("wb.xlsx")      // wp.LockStructure / wp.PasswordHash ...

// ---- Workbook-level defined (named) ranges ----
_ = excelgo.SetDefinedName("wb.xlsx", "MyRange", "Sheet1!$A$1:$D$10")
v, _ := excelgo.GetDefinedName("wb.xlsx", "MyRange")   // "Sheet1!$A$1:$D$10"
_ = excelgo.DeleteDefinedName("wb.xlsx", "MyRange")
```

The object API mirrors the same (see `Book.ProtectWorkbook` / `Book.UnprotectWorkbook` /
`Book.GetWorkbookProtection` / `Book.SetDefinedName` / `Book.GetDefinedName` /
`Book.DeleteDefinedName`).

Parameters & notes:

- **ProtectWorkbook**: writes `<workbookProtection lockStructure="1">` into
  `xl/workbook.xml` (workbook structure protection only — distinct from `ProtectSheet`,
  which protects a single worksheet). With a non-empty `password`, the password is stored as
  Excel's standard 16-bit password-verifier hash (4 uppercase hex chars) in the
  `workbookPassword` attribute. `UnprotectWorkbook` removes the element.
- **GetWorkbookProtection**: returns the current state; `LockStructure` etc. are `false` and
  `PasswordHash` is `""` when the workbook is unprotected.
- **SetDefinedName / GetDefinedName / DeleteDefinedName**: manage *workbook-global* named
  ranges (no `localSheetId`), stored under `<definedNames>` in `xl/workbook.xml`. The reference
  string is stored verbatim (e.g. `Sheet1!$A$1:$D$10`); sheet-scoped names such as the print
  area (`_xlnm.Print_Area`) are handled separately by the page-setup API and are written with
  a `localSheetId`.
- All of the above pass the openpyxl strict-load cross-check (workbook protection is read back
  as `wb.security.lock_structure`; a print area round-trips as `ws.print_area`).

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
excelgo copy  <input.xlsx> <output.xlsx> <sheet> [--name newname] [--independent] [--suffix suffix]
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
| `renamesheet` | `RenameSheet` | rename (auto numeric suffix on conflict) |
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
| `copy` | `CopySheet` / `(*Book).CopySheetTo` | worksheet copy (package / object) |
| `merge` | `MergeWorkbook` / `(*Book).Merge` | cross-workbook merge (package / object) |

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
go test -timeout 30m ./...
```

> **A note on the timeout.** The suite takes roughly **10 minutes**, because every differential
> test spawns a Python/openpyxl subprocess to act as an independent oracle. That exceeds
> `go test`'s default 10-minute limit, hence the explicit `-timeout 30m` above. Without it you
> get `panic: test timed out after 10m0s` — which reads like a hang but is only the ceiling
> being reached. To run a single case: `go test -run TestName -v .`

Unit tests (`excelgo/*_test.go`, `cmd/excelgo/*_test.go`) cover the object API, formula
reference shifting, copy/merge, style dedup, merge cells / column width / freeze / hyperlink
/ filter, data validation / conditional format / table / comment / sheet properties,
protection / grouping / replace / document properties, and more.

### Differential testing (openpyxl as an external oracle)

Unit tests have a structural blind spot: they verify that *this library's output matches this
library's expectations*. If the expectation itself is wrong, nothing catches it.

`difftest/` uses openpyxl (an independent implementation, widely validated in the Excel
ecosystem) as an external oracle: the same set of documents, the same set of operations,
executed by both sides, then normalized into **semantic snapshots** and compared field by
field. It catches two classes of problem existing tests miss:

- this library produces things other tools can't read, or read differently (real bugs)
- this library can't read openpyxl's normal output (compatibility gaps)

| File | Role |
| --- | --- |
| `difftest/snapshot.py` | normalizes an xlsx into comparable JSON (semantic layer, not XML bytes) |
| `difftest/op.py` | executes the same operations via openpyxl from JSON instructions |
| `difftest_test.go` | harness: drives both sides, compares, exempts by field path |
| `difftest_write_test.go` | 34 write-side scenarios (values/formulas/styles/rows+cols/layout/sheets) |
| `difftest_read_test.go` | 12 read-side scenarios (openpyxl writes → excelgo reads) + round-trip |
| `difftest_composite_test.go` | Composite ops: `CopySheet` / `CopySheetTo` / `MergeWorkbook` fidelity |
| `difftest_picture_test.go` | Pictures: WPS inline pictures, floating pictures, media strategies, `ExtractPicture` |
| `difftest_style_test.go` | Style enum details: underline / border style / alignment / number formats + rejection of invalid values |
| `difftest_chart_test.go` | Charts: transport and closure of the **multi-level chain** (sheet→drawing→chart) |
| `difftest_protect_test.go` | Sheet/workbook protection: hash compatibility with Excel, retention after copy/merge |
| `difftest_combo_test.go` | **Feature combinations (transporting existing files)**: 6 openpyxl-generated combos × 3 operations |
| `difftest_selfcombo_test.go` | **Feature combinations (excelgo-authored)**: chart creation + named style + table + array formula + conditional formats + data validation + outline + merged cells all on one sheet, including "copy the copy again" |
| `difftest_readsem_test.go` | Read semantics: values/types, formulas, shared strings, boundary coords, date divergence |

Composite operations cannot be compared by "both sides run the same operation": openpyxl
has no `MergeWorkbook`, and `copy_worksheet` is a shallow copy (it carries no pictures /
conditional formats / data validations). There openpyxl instead acts as a **result
validator**: excelgo performs the composite operation, then openpyxl independently opens
the output and checks that semantics match and that **associated parts came along**
(drawing / media / tables / comments) — the latter being the core of "layout fidelity" and
something a plain semantic snapshot cannot see at all.

```bash
go test -run TestDiff ./...        # ~2.5 min
```

When openpyxl isn't installed these tests skip automatically (no hard CI dependency). Point
at a specific interpreter with the `EXCELGO_PYTHON` environment variable.

### Real bugs it caught

The value of differential testing: these defects pass every unit test, because they can
only surface through an outside implementation.

| Defect | Symptom | Root cause |
| --- | --- | --- |
| Column-width ranges aren't interoperable | on the very same file, excelgo reads 3 columns as 12 while openpyxl reads only D | it wrote `<col min="4" max="6">` (legal, but consumers treat column defs as a sparse map). openpyxl expands per column itself, so this library now does too |
| Copying a sheet with a table makes the whole file unreadable | openpyxl raises `Table with name X already exists` | part copying is byte-level, so the copy carries the same `id` and `name`/`displayName` — both of which are **workbook-scoped unique** |
| Wrong conditional-format type makes the whole file unreadable | openpyxl raises `Unable to read workbook` | `cfType` is written into XML verbatim. The library's own doc comment says `cellIs`, but passing `cell` (the intuitive spelling) produces a broken file with **no validation at all** |
| Same for data-validation type/operator | same | `typ`/`op` are likewise written verbatim, unvalidated |
| Fill pattern / underline / border style / alignment | same | one class of defect: enum values treated as free text |
| Copying/merging a sheet with charts loses them and corrupts the file | openpyxl raises `There is no item named 'xl/charts/chart1.xml'` | part copying only handled **one level** of rels. The chain is multi-level (sheet→drawing→chart), so drawing's rels got copied verbatim and nothing further — charts were never copied |
| Merging a cross-workbook file with conditional formats corrupts it | openpyxl raises `IndexError: list index out of range` (reading `differential_styles[dxfId]`) | `dxfId` is a **workbook-scoped index**. Only `cellXf` was merged, not `<dxfs>`, leaving every `dxfId` dangling |
| Merging a cross-workbook file with charts/tables corrupts it | openpyxl raises `Unknown relationship: rId1` | **double-write conflict**: the part transport rewrote the sheet's rIds and wrote the map, then the caller overwrote it with its own copy — losing the rId rewrite |
| `CopySheetTo` of a cross-workbook file with conditional formats corrupts it | openpyxl raises `IndexError: list index out of range` | same class but a **different code path** (`copySheetAcrossMaps`, not `mergeOneSheet`). Fixing one and missing the other is why `mergeDxfsAndRemapSheet` is now shared by both |
| Transporting named styles across workbooks corrupts the file | openpyxl raises `TypeError: expected <class 'int'>` | `<cellStyle xfId>` points at **cellStyleXfs** (not cellXfs), and those xfs reference fontId/fillId in turn. Moving only the `<cellStyles>` manifest leaves xfId dangling. Now handled by `StylesMerger.mergeStyleXfs` |
| Integer attribute written as an empty string | openpyxl raises `TypeError: expected <class 'int'>` | `buildXF` emitted `xfId=""` when the source xf had none. **Integer attributes must be omitted entirely when absent**, never written empty |

**A general lesson**: the write side must validate OOXML enum domains. Escaping cannot
catch these — they aren't "wrong values", they're "the whole file becomes unreadable by
consumers", and **the write returns nil and the file is produced**; only Excel/WPS
reports "the file is corrupt" at open time, which is extremely hard to trace back to one
field.

Guard: centralized validation via `validateStyleEnums` / `validatePatternType` /
`validCFTypes` / `validDVTypes`, returning explicit errors that list the correct values
(passing `cell` now says "should be cellIs").

### Known semantic differences vs openpyxl

A few differences are **intentional** — each implementation is self-consistent on its own,
but code migrated from openpyxl will trip on them.

| Scenario | openpyxl | this library | Note |
| --- | --- | --- | --- |
| Reading a date cell (number + date format) | `cell.value` is a `datetime` | `GetCellValue` gives the raw serial; `Cell.GetTime()` converts to `time.Time` | this library never loses the raw value; conversion is explicit |
| Writing formulas | writes no cached result | writes `result` | this library can produce formulas with a cached result |
| Writing an empty string | round-trips to `None` | round-trips to `""` | openpyxl round-trip loss; `""` being a valid value is more correct here |
| Conditional-format type name | `cellIs` | same (plus alias hints) | aligned |
| Column-width ranges | expanded per column | expanded per column | aligned (merging is legal but consumers use a sparse map) |

`TestDateSemanticsDivergence` locks the date behaviour down: if someone changes it to
auto-convert to `datetime`, the test fails and points at this table.

### Capability status vs openpyxl

The gaps listed earlier have now been **filled in one by one**:

| Capability | Status | API |
| --- | --- | --- |
| Chart creation | ✅ 12 types (bar/line/pie/scatter/area/doughnut/radar/bubble/surface + 3D) | `AddChart` / `(*WorkSheet).AddChart` |
| Array formulas | ✅ emits `t="array"` + `ref`, CSE-capable | `SetCellArrayFormula` |
| Multi-rule conditional formats | ✅ 10 types incl. ColorScale / DataBar / IconSet / top10 / aboveAverage | `SetConditionalFormatRules` |
| Move region | ✅ values + styles, formulas remapped automatically | `MoveRange` |
| Named styles | ✅ workbook-level `cellStyles` templates, shareable across sheets | `NewNamedStyle` / `SetNamedStyle` / `GetNamedStyles` |
| Outline collapse | ✅ writes both `hidden` and `collapsed`, summary position configurable | `CollapseRows` / `ExpandRows` / `SetOutlineSummary` |
| Read-only streaming | ✅ reads sheet XML on demand, skips unrelated parts | `OpenReader` / `StreamRows` |
| write_only streaming write | ✅ appends row by row, memory doesn't grow with data | `NewStreamWriter` / `StreamSheet.Append` |
| Printer-settings binary | ✅ read/write/transport, page-setup changes keep `r:id` | `GetPrinterSettings` / `SetPrinterSettings` |

Still missing (low priority; openpyxl's own support is limited too):

| Missing | Notes |
| --- | --- |
| Pivot tables | unsupported (openpyxl also only reads/writes definitions) |
| Chartsheets | unsupported (charts must live in a normal worksheet) |
| Synthesizing printer settings | this library can read, transport and write back the binary, but cannot **synthesize** a DEVMODE (it encodes a specific printer driver's capabilities). Same as openpyxl |
| Advanced chart features | data tables / trendlines / secondary axis / combo charts not exposed |

### Performance: write_only streaming write

The regular write path rescans the entire sheet XML on **every cell write**, i.e.
**O(cells × document length)**. Measured: 3000 rows × 4 columns takes **over 4 minutes**;
streaming 20000 rows takes **224 ms** — roughly 3000× faster.

```go
w, _ := excelgo.NewStreamWriter("out.xlsx")
head, _ := w.AddStyle(excelgo.Style{Font: &excelgo.FontStyle{Bold: true}})
w.SetColWidth(1, 3, 18)

ws, _ := w.NewSheet("data")
ws.Append([]interface{}{excelgo.StreamCell{Value: "Name", Style: head}, "Qty"})
for i := 1; i <= 20000; i++ {
    ws.Append([]interface{}{fmt.Sprintf("item-%d", i), i})
}
w.Save()
```

Same constraints as openpyxl's `write_only`:

- **Declare sheet names before writing rows** — the sheet-name ↔ `sheetN.xml` mapping
  is fixed in `workbook.xml`; renaming after writing leaves data nowhere to go
- **Rows must be appended in order** — `<row>` is a sequential stream; random
  back-filling requires rescanning the whole document
- **Styles must be registered up front via `AddStyle`** — you can't back-fill a
  style index after the cell has been written
- Strings use `t="inlineStr"` (no shared-string dedup — dedup means rescanning the
  whole table, exactly what streaming exists to avoid); the cost is a slightly larger file

One **hard format constraint** worth remembering: `archive/zip`'s `Writer` allows
only one open entry at a time — `Create` on a new entry closes the previous one.
So multiple sheets can't be written interleaved: rows go to temp files first,
then get emitted as zip entries in order.

### Two lessons

**1. Relation chains must be walked recursively.** OOXML part relations are multi-level
(`sheet → drawing → chart → embedded workbook`); handling only the first level leaves
dangling relationships. The library now has `copyDownstreamParts` (same workbook) and
`copyDownstreamPartsCross` (cross-workbook) for the recursion, and they place each
downstream part in **its own canonical directory** (charts belong in `xl/charts/`, never
in `xl/drawings/`).

**2. Part transport and sheet write-back must be a single write point.**
`copySheetPartsToTarget` does two things: moves parts and **rewrites the sheet's rId
references**. It used to write the map itself, then the caller overwrote it with its
own copy — which lacked the rId rewrite, so the sheet still pointed at the source's
old rIds while the rels had already been renumbered. It now **returns** the rewritten
content; the caller writes it back once, after every rewrite (style index / dxfId /
shared strings / rId) is done. `TestPartTransportSingleWritePoint` guards this.

**3. Cross-workbook transport must account for every workbook-scoped index table.**
In `styles.xml` each table is indexed workbook-wide: `s` → cellXfs, `dxfId` → dxfs,
`xfId` → cellStyleXfs — and those xfs reference fontId / fillId / borderId / numFmtId
in turn: **a whole indirect chain**. Moving only the outermost table necessarily leaves
dangling references, and the symptom is always an openpyxl `IndexError` or `TypeError`
that says nothing about which table is short. Now `StylesMerger` (cellXf +
cellStyleXfs) and `mergeDxfsAndRemapSheet` (dxfs + the cellStyles manifest) split the
work, shared by both cross-workbook paths (`Merge` / `CopySheetTo`).

**4. One write point per destination part.** Both dxf merging and named-style merging
write `dstMap["xl/styles.xml"]`; if each starts from the original they overwrite each
other — my first attempt used a standalone `mergeNamedStyles` that clobbered the
freshly-merged dxfs, and openpyxl threw IndexError again. `appendNamedStyles` is now a
pure function (bytes in, bytes out), sequenced inside `mergeDxfsAndRemapSheet`.

**5. Omit Integer attributes entirely when absent.** Values like `xfId=""` make
consumers fail `_convert("")` with `TypeError: expected <class 'int'>`, and the whole
workbook becomes unreadable.

**6. Password entry points must be explicit about their semantics.** OOXML's `password`
attribute holds a **hash**. Historically this library's `ProtectSheet` required an
already-hashed value while `ProtectWorkbook` took plaintext — an inconsistency that
invites misuse (passing plaintext yields protection that looks real but can never be
unlocked in Excel). `ProtectSheetWithPassword` (plaintext in, hashes internally) is now
the recommended entry point; the old one stays for compatibility. The hash algorithm was
independently verified against openpyxl, including Chinese text and special characters.

Design notes:

- **Normalize to the semantic layer.** The two sides necessarily lay out XML differently;
  comparing bytes is meaningless, comparing "what a user can observe" is not.
- **Separate real differences from capability differences.** openpyxl has features this
  library lacks, and there are deliberate behavioural differences. Each case may carry
  `Exempts` to skip specific **field paths** rather than the whole scenario — the former is
  a real problem, the latter a known trade-off.
- **Every exemption carries a reason.** For empty strings: excelgo writes a shared string and
  reads back `""`, while openpyxl can't read back even its own empty string (measured: its
  round-trip also yields `None`). That's a round-trip loss in openpyxl; this library's
  behaviour better matches "an empty string is a valid value".


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

## XML security model (important)

### Why not use an XML parser

This isn't a stylistic preference — **`encoding/xml` round-trips break OOXML**.
Measured with a lossless `xml.Decoder` + `xml.Encoder` round-trip on a WPS-style fragment:

| Before | After round-trip |
| --- | --- |
| `<mc:AlternateContent>` | `<AlternateContent>` (**prefix lost**) |
| `<x14:picture>` | `<picture>` |
| `mc:Ignorable="x14ac"` | `x:Ignorable="x14ac"`, prefix renamed to `_xmlns:x14ac` |

The last row is the killer: **the value of `mc:Ignorable` is a prefix name**. Go's
`encoding/xml` does not preserve original prefixes (it resolves a prefix into
`Name.Space` and generates its own on re-encode), so the prefix `Ignorable` points at no
longer exists — **MCE (Markup Compatibility) fallback silently stops working**, changing how
Excel and WPS interpret the file. Typed `Unmarshal`+`Marshal` is worse: any element not
declared in the struct is dropped entirely (`definedNames` / `bookViews` / `calcPr` all
disappear) and `r:id` gets rewritten to a different prefix.

Hence this library edits XML parts **as strings/regex** — that is what preserves WPS
inline pictures (`mc:AlternateContent` + `x14:picture`), namespaces and layout.

### Escaping convention and guards

The cost is that **every write site must escape by itself**. Convention:

- all user data written into XML goes through **and only through** `safeText`
  (text nodes) / `safeAttr` (attribute values);
- the old names `escapeXML` / `escapeAttr` remain as equivalent aliases;
- user data entering a **regex** must go through `regexp.QuoteMeta`.

Protections in place:

| Protection | Notes |
| --- | --- |
| Single entry point | `safeText` / `safeAttr` are the only sanctioned escaping exits; all write sites migrated |
| Control-character stripping | XML 1.0 forbids `U+0000–U+0008 / U+000B / U+000C / U+000E–U+001F` (keeping `\t \n \r`). These **cannot be expressed as entities**; writing them makes a whole part not-well-formed and Excel reports the file as damaged. `safeText` strips them centrally |
| Name validation | Worksheet names ≤ 31 chars, rejecting `\ / ? * [ ] :` (`AddSheet` / `NewSheet` / `RenameSheet` / `CopySheet`) |

**Two automated guards** (neither relies on anyone remembering):

1. `TestNoTaintedXMLWrite` — an AST taint-tracking guard with **zero manual exemptions**.
   It analyses data flow per function: parameters, local variables, struct fields,
   `range` variables and `append` results start out tainted; the taint is cleared by
   `safeText`/`safeAttr`/`strconv.*`. A tainted value reaching an XML tag — whether as a
   bare concatenation or as a function argument — fails the test.

   **Safety is expressed by type, not by a list.** These named types count as certified
   because they can only be produced by a validating constructor:

   | Type | Constructor | Meaning |
   | --- | --- | --- |
   | `cellRef` | `newCellRef` | validated cell reference (`A1`) |
   | `rangeRef` | `newRangeRef` | validated range reference (`A1:C10`) |
   | `relID` | `newRelID` | library-generated relationship id (`rIdN`) |
   | `styleIndex` | `idx` | style index (`fontId`/`numFmtId`, …) |
   | `borderSideName` | — | fixed border side name |
   | `xmlTagName` | `tagName` | fixed schema tag name (`v`/`f`/`r`) |
   | `xmlFrag` | `frag` | newly built fragment whose content passed `safe*` |
   | `partXML` | `part` | **pre-existing** part XML (reused, not newly written) |

   So there is **no "forgot to add a list entry" failure mode** — new write sites are
   protected by default (a bare `string` is always tainted). `TestGuardNoNameAllowlist`
   locks this invariant: if a per-variable exemption list ever reappears in the guard,
   that test fails.
2. `TestXMLInjectionResistance` — end-to-end injection regression. 6 payload classes
   (tag closing, closing + entities, attribute-quote escape, CDATA masquerade, comment
   masquerade, control characters) × 12 write sites (cell values/formulas, shared strings,
   comment text & author, hyperlink URL & display text, defined-name, doc properties,
   table name, data validation, conditional format, header-footer / tab color). It asserts
   every XML part is **well-formed** (Go's strict `encoding/xml` decoder) and that no extra
   node was injected.

> ⚠️ **When adding a new write site**: call `safeText` (text nodes) or `safeAttr`
> (attribute values), and use `regexp.QuoteMeta` for anything entering a regex. If a
> value really is "generated safely by the library", the right move is to **give it a
> named type plus a validating constructor** (like `cellRef`/`newCellRef`) — not to add
> another per-name exemption, which just relocates the "remember to update the list"
> problem.

## License

MIT
