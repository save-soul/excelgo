package excelgo

// 工作表 / 工作簿保护的差分验证。
//
// 保护有个容易踩的语义陷阱：OOXML 的 `password` 属性存的是**哈希值**而非明文。
// 本库 Protect(password) 要求调用方自己传已哈希的字符串，而 openpyxl 的
// SheetProtection(password=...) 传的是**明文**、由它内部算哈希。
//
// 两侧参数语义不同，但最终写进 XML 的哈希值必须一致 —— 这才是可比的地方。
// 差分测试要验证的正是这一点，以及：
//   - 哈希算法与 Excel 兼容（否则用户按明文设的密码在 Excel 里解不开）
//   - 保护设置在复制 / 合并 / 改名 / 删表后仍然有效
//   - 无密码保护（只锁结构）在两侧行为一致

import (
	"path/filepath"
	"regexp"
	"strings"
	"testing"
)

// sheetProtectionAttrs 提取 sheet XML 里的 <sheetProtection .../> 属性。
func sheetProtectionAttrs(t *testing.T, p, sheetName string) map[string]string {
	t.Helper()
	sheetFile := findSheetFileFor(t, p, sheetName)
	data := readPartStr(t, p, sheetFile)
	m := regexp.MustCompile(`<sheetProtection\b([^>]*)/?>`).FindStringSubmatch(data)
	if m == nil {
		return nil
	}
	attrs := map[string]string{}
	for _, kv := range regexp.MustCompile(`([\w:]+)="([^"]*)"`).FindAllStringSubmatch(m[1], -1) {
		attrs[kv[1]] = kv[2]
	}
	return attrs
}

// workbookProtectionAttrs 提取 workbook.xml 里的 <workbookProtection .../>。
func workbookProtectionAttrs(t *testing.T, p string) map[string]string {
	t.Helper()
	data := readPartStr(t, p, "xl/workbook.xml")
	m := regexp.MustCompile(`<workbookProtection\b([^>]*)/?>`).FindStringSubmatch(data)
	if m == nil {
		return nil
	}
	attrs := map[string]string{}
	for _, kv := range regexp.MustCompile(`([\w:]+)="([^"]*)"`).FindAllStringSubmatch(m[1], -1) {
		attrs[kv[1]] = kv[2]
	}
	return attrs
}

// TestDiffSheetProtect 验证工作表保护：两侧的哈希值一致且 Excel 兼容。
func TestDiffSheetProtect(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()

	// 本库侧：传已哈希值（OOXML legacy hash，Excel 用同一算法）
	gotFile := filepath.Join(dir, "excelgo.xlsx")
	if err := Create2(gotFile); err != nil {
		t.Fatal(err)
	}
	// "1234" 的 Excel legacy hash 是 CC3D（由 openpyxl hash_password 独立算出）
	if err := ProtectSheet(gotFile, "Sheet1", "CC3D"); err != nil {
		t.Fatal(err)
	}

	// openpyxl 侧：传真文，由它内部算哈希
	wantFile := filepath.Join(dir, "openpyxl.xlsx")
	runOpenpyxlOps(t, wantFile,
		`{"out": `+quoteJSON(wantFile)+`, "ops": [{"op":"protect","sheet":"Sheet1","password":"1234"}]}`)

	gotAttrs := sheetProtectionAttrs(t, gotFile, "Sheet1")
	wantAttrs := sheetProtectionAttrs(t, wantFile, "Sheet1")
	if gotAttrs == nil {
		t.Fatalf("本库侧未写入 sheetProtection")
	}
	if wantAttrs == nil {
		t.Fatalf("openpyxl 侧未写入 sheetProtection")
	}

	// 1) 密码哈希必须一致（这是两侧唯一的公共语言）
	if gotAttrs["password"] != wantAttrs["password"] {
		t.Errorf("密码哈希不一致：本库=%q openpyxl=%q\n"+
			"本库要求传**已哈希**值（Excel legacy hash），openpyxl 传明文由内部计算。\n"+
			"若本库的值与 Excel 不兼容，用户设的密码在 Excel 里解不开",
			gotAttrs["password"], wantAttrs["password"])
	}

	// 2) 保护开关必须生效
	if gotAttrs["sheet"] != "1" {
		t.Errorf("sheet 属性 = %q，期望 1（保护未生效）", gotAttrs["sheet"])
	}

	// 3) openpyxl 必须能读回
	runSnapshot(t, gotFile)
	assertNoDanglingRels(t, gotFile)
}

// TestDiffSheetProtectNoPassword 验证无密码保护：两侧都应只锁结构、不写 password。
func TestDiffSheetProtectNoPassword(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "nopw.xlsx")
	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	if err := ProtectSheet(p, "Sheet1", ""); err != nil {
		t.Fatal(err)
	}
	attrs := sheetProtectionAttrs(t, p, "Sheet1")
	if attrs == nil {
		t.Fatal("未写入 sheetProtection")
	}
	if attrs["sheet"] != "1" {
		t.Errorf("sheet = %q，期望 1", attrs["sheet"])
	}
	if pw, ok := attrs["password"]; ok && pw != "" {
		t.Errorf("无密码保护不应写 password 属性，实际 = %q", pw)
	}
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	runSnapshot(t, p)
}

// TestDiffProtectSurvivesCopy 验证保护设置在复制工作表后仍然存在。
//
// 这是实际使用中容易出错的点：保护是"表级"设置，复制时若只搬 sheetData，
// 副本就会变成未保护 —— 用户以为复制了一份受控表，实际可以随意编辑。
func TestDiffProtectSurvivesCopy(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	p := filepath.Join(dir, "protsheet.xlsx")
	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(p, "Sheet1", "受控表"))
	must(t, SetCellStr(p, "受控表", "A1", "数据"))
	if err := ProtectSheet(p, "受控表", "CC3D"); err != nil {
		t.Fatal(err)
	}

	g, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := g.CopySheet("受控表", "受控表副本"); err != nil {
		t.Fatal(err)
	}
	if err := g.Save(); err != nil {
		t.Fatal(err)
	}

	// 源表与副本都应保持保护
	for _, name := range []string{"受控表", "受控表副本"} {
		attrs := sheetProtectionAttrs(t, p, name)
		if attrs == nil {
			t.Errorf("%s 复制后丢失 sheetProtection —— "+
				"用户以为复制了受控表，实际副本可随意编辑", name)
			continue
		}
		if attrs["sheet"] != "1" {
			t.Errorf("%s 的 sheet = %q，期望 1", name, attrs["sheet"])
		}
		if attrs["password"] != "CC3D" {
			t.Errorf("%s 的密码哈希 = %q，期望 CC3D（复制时应保持一致）",
				name, attrs["password"])
		}
	}
	runSnapshot(t, p)
	assertNoDanglingRels(t, p)
}

// TestDiffWorkbookProtect 验证工作簿结构保护在两侧一致，且复制/合并后仍有效。
func TestDiffWorkbookProtect(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()

	// 本库侧
	gotFile := filepath.Join(dir, "wb_excelgo.xlsx")
	if err := Create2(gotFile); err != nil {
		t.Fatal(err)
	}
	// 注意 API 语义差异：ProtectWorkbook 收**明文**（内部哈希），
	// 而 ProtectSheet 要求传**已哈希**值。这个不一致本身值得记录。
	if err := ProtectWorkbook(gotFile, "1234"); err != nil {
		t.Fatal(err)
	}

	// openpyxl 侧同样传真文 "1234" —— 故两侧应得到同一哈希
	wantFile := filepath.Join(dir, "wb_openpyxl.xlsx")
	runOpenpyxlOps(t, wantFile,
		`{"out": `+quoteJSON(wantFile)+`, "ops": [{"op":"protect_workbook","password":"1234"}]}`)

	gotAttrs := workbookProtectionAttrs(t, gotFile)
	wantAttrs := workbookProtectionAttrs(t, wantFile)
	if gotAttrs == nil {
		t.Fatalf("本库侧未写入 workbookProtection")
	}
	if gotAttrs["lockStructure"] != "1" {
		t.Errorf("lockStructure = %q，期望 1（结构保护未生效）", gotAttrs["lockStructure"])
	}
	if wantAttrs["lockStructure"] != "1" {
		t.Errorf("openpyxl 侧 lockStructure = %q，期望 1", wantAttrs["lockStructure"])
	}
	// 密码哈希应一致
	if gp, wp := gotAttrs["workbookPassword"], wantAttrs["workbookPassword"]; gp != wp {
		t.Errorf("工作簿密码哈希不一致：本库=%q openpyxl=%q", gp, wp)
	}

	// 解除保护后应能重新读出未保护状态
	if err := UnprotectWorkbook(gotFile); err != nil {
		t.Fatal(err)
	}
	if attrs := workbookProtectionAttrs(t, gotFile); attrs != nil {
		if attrs["lockStructure"] == "1" {
			t.Errorf("UnprotectWorkbook 后 lockStructure 仍为 1: %v", attrs)
		}
	}
	runSnapshot(t, gotFile)
}

// TestDiffProtectAfterMerge 验证合并进来的受保护工作表保留其保护设置。
func TestDiffProtectAfterMerge(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()
	base := filepath.Join(dir, "base.xlsx")
	src := filepath.Join(dir, "src.xlsx")

	if err := Create2(base); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(base, "Sheet1", "Base"))

	if err := Create2(src); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(src, "Sheet1", "受控源表"))
	must(t, SetCellStr(src, "受控源表", "A1", "受控数据"))
	if err := ProtectSheet(src, "受控源表", "CC3D"); err != nil {
		t.Fatal(err)
	}

	if err := MergeWorkbook(base, []SourceRef{{Workbook: src, Sheet: "受控源表"}}); err != nil {
		t.Fatal(err)
	}

	attrs := sheetProtectionAttrs(t, base, "受控源表")
	if attrs == nil {
		t.Fatalf("合并后受控源表丢失 sheetProtection: %v", partNames(t, base))
	}
	if attrs["sheet"] != "1" || attrs["password"] != "CC3D" {
		t.Errorf("合并后保护属性异常: %v", attrs)
	}
	runSnapshot(t, base)
	assertNoDanglingRels(t, base)
}

// TestDiffProtectRenameSheet 验证重命名工作表不影响保护设置。
//
// 保护是写在 sheet XML 里的，与表名无关；但若重命名实现误伤了 sheet XML，
// 保护就会静默丢失。
func TestDiffProtectRenameSheet(t *testing.T) {
	dir := t.TempDir()
	p := filepath.Join(dir, "ren.xlsx")
	if err := Create2(p); err != nil {
		t.Fatal(err)
	}
	must(t, RenameSheet(p, "Sheet1", "原名"))
	if err := ProtectSheet(p, "原名", "CC3D"); err != nil {
		t.Fatal(err)
	}
	if err := RenameSheet(p, "原名", "新名"); err != nil {
		t.Fatal(err)
	}
	attrs := sheetProtectionAttrs(t, p, "新名")
	if attrs == nil {
		t.Fatalf("重命名后保护丢失")
	}
	if attrs["password"] != "CC3D" || attrs["sheet"] != "1" {
		t.Errorf("重命名后保护属性异常: %v", attrs)
	}
	if strings.Contains(readPartStr(t, p, findSheetFileFor(t, p, "新名")), "原名") {
		t.Log("提示：sheet XML 里仍含旧表名（若表名出现在定义名称里属正常）")
	}
}

// TestDiffProtectPasswordEntryPoints 验证两个密码入口的语义与 Excel 兼容性。
//
// 背景：ProtectSheet 要求传**已哈希**值（历史约定），而 ProtectWorkbook 收
// **明文**。这个不一致必然导致误用 —— 用户给 ProtectSheet 传真文，会得到一个
// 看似设了密码、实际在 Excel 里永远解不开的保护。
//
// 故新增 ProtectSheetWithPassword（收明文），与 ProtectWorkbook 语义对齐。
// 本测试断言：明文入口算出的哈希与 openpyxl（独立的 Excel legacy 实现）一致。
func TestDiffProtectPasswordEntryPoints(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	dir := t.TempDir()

	for _, pw := range []string{"1234", "password", "中文密码", "a", "P@ssw0rd!"} {
		pw := pw
		t.Run(pw, func(t *testing.T) {
			// 本库：明文入口
			got := filepath.Join(dir, "g"+pw+".xlsx")
			if err := Create2(got); err != nil {
				t.Fatal(err)
			}
			if err := ProtectSheetWithPassword(got, "Sheet1", pw); err != nil {
				t.Fatal(err)
			}
			// openpyxl：同样是明文，内部独立算哈希
			want := filepath.Join(dir, "w"+pw+".xlsx")
			ops := `{"op":"protect","sheet":"Sheet1","password":` + quoteJSON(pw) + `}`
			runOpenpyxlOps(t, want, `{"out": `+quoteJSON(want)+`, "ops": [`+ops+`]}`)

			g := sheetProtectionAttrs(t, got, "Sheet1")
			w := sheetProtectionAttrs(t, want, "Sheet1")
			if g == nil || w == nil {
				t.Fatalf("保护属性缺失: 本库=%v openpyxl=%v", g, w)
			}
			if g["password"] != w["password"] {
				t.Errorf("明文 %q 的哈希不一致: 本库=%q openpyxl=%q\n"+
					"哈希算法必须与 Excel 兼容，否则用户设的密码在 Excel 里解不开",
					pw, g["password"], w["password"])
			}
			if g["password"] == "" {
				t.Errorf("明文 %q 未产生密码哈希", pw)
			}
			// 旧入口（已哈希）应写入调用方给的值，不做二次哈希
			legacy := filepath.Join(dir, "l"+pw+".xlsx")
			if err := Create2(legacy); err != nil {
				t.Fatal(err)
			}
			if err := ProtectSheet(legacy, "Sheet1", "CC3D"); err != nil {
				t.Fatal(err)
			}
			la := sheetProtectionAttrs(t, legacy, "Sheet1")
			if la == nil || la["password"] != "CC3D" {
				t.Errorf("旧入口应原样写入已哈希值，实际 = %v", la)
			}
		})
	}
}
