package excelgo

import (
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// TestFeatureCombinations 在**真实组合**文件上跑 excelgo 的复合操作。
//
// 为什么单独一层：至今抓到的 bug 几乎全是组合场景暴露的 ——
// 单项功能各自都对，叠在一起就出问题（表格 id 重复、多级关系链断、
// pageSetup 的 r:id 被抹）。这层测试的价值就在"叠"。
//
// 输入由 difftest/probe_combos.py 用 openpyxl 生成 —— 库自身造不出
// 这么复杂的组合（有 openpyxl 当外部实现才有意义）。
func TestFeatureCombinations(t *testing.T) {
	if pythonExe() == "" {
		t.Skip("未找到 Python/openpyxl")
	}
	comboDir := comboFixtureDir(t)

	cases := []struct {
		file string
		desc string
		// 需要在该文件里存在的要素（复制/合并后逐项验证）
		wantCharts  int
		wantTables  int
		wantCondFmt bool
	}{
		{"probe_combo1.xlsx", "图表+命名样式+条件格式+冻结", 1, 0, true},
		{"probe_combo2.xlsx", "数据验证+表格+条件格式", 0, 1, true},
		{"probe_combo3.xlsx", "图表+打印区域+页眉", 1, 0, false},
		{"probe_combo4.xlsx", "多图表+多表格", 3, 3, false},
		{"probe_combo5.xlsx", "隐藏行列+大纲折叠+条件格式", 0, 0, true},
		{"probe_combo6.xlsx", "超链接+批注+保护", 0, 0, false},
	}

	for _, c := range cases {
		c := c
		t.Run(c.desc, func(t *testing.T) {
			src := filepath.Join(comboDir, c.file)

			t.Run("CopySheet", func(t *testing.T) {
				dst := t.TempDir() + "/cp.xlsx"
				copyFileForTest(t, src, dst)
				g, err := Open(dst)
				mustNoErr(t, err)
				names := g.SheetNames()
				if len(names) == 0 {
					t.Fatal("没有工作表")
				}
				if _, err := g.CopySheet(names[0], "副本"); err != nil {
					t.Fatalf("CopySheet 失败: %v", err)
				}
				mustNoErr(t, g.Save())

				// 1) openpyxl 必须能严格解析（组合错乱通常在这里暴露）
				snap := runSnapshot(t, dst)
				sheets, _ := snap["sheets"].(map[string]interface{})
				if _, ok := sheets["副本"]; !ok {
					t.Errorf("副本未出现在快照里: %v", sheets)
				}
				// 2) 关联部件数量：源表与副本各一份
				if c.wantCharts > 0 {
					got := chartCountOfSheet(t, dst, "副本")
					if got != c.wantCharts {
						t.Errorf("副本的图表数 = %d，期望 %d", got, c.wantCharts)
					}
				}
				if c.wantTables > 0 {
					if n := countParts(t, dst, "xl/tables/table"); n < c.wantTables*2 {
						t.Errorf("表格部件数 = %d，期望 ≥%d（源表与副本各一份）",
							n, c.wantTables*2)
					}
					// 表格的 id 与 name 必须工作簿级唯一
					assertTableIdentityUnique(t, dst)
				}
				// 3) 条件格式必须跟着搬
				if c.wantCondFmt {
					assertCondFmtSurvived(t, dst, "副本", c.desc)
				}
				// 4) 关系不得悬空
				assertNoDanglingRels(t, dst)
			})

			t.Run("Merge", func(t *testing.T) {
				dst := t.TempDir() + "/merged.xlsx"
				b, _ := Create()
				mustNoErr(t, b.SaveAs(dst))
				srcNames := sheetNamesOf(t, src)
				if len(srcNames) == 0 {
					t.Fatal("源文件没有工作表")
				}
				refs := make([]SourceRef, 0, len(srcNames))
				for _, n := range srcNames {
					refs = append(refs, SourceRef{Workbook: src, Sheet: n})
				}
				if err := b.Merge(refs); err != nil {
					t.Fatalf("Merge 失败: %v", err)
				}
				mustNoErr(t, b.Save())

				runSnapshot(t, dst) // openpyxl 严格解析
				if c.wantTables > 0 {
					assertTableIdentityUnique(t, dst)
				}
				assertNoDanglingRels(t, dst)
			})

			t.Run("StreamReader", func(t *testing.T) {
				// 流式读必须能处理组合文件（含大部件）
				sb, err := OpenReader(src)
				mustNoErr(t, err)
				defer sb.Close()
				names := sb.GetSheetList()
				if len(names) == 0 {
					t.Fatal("流式读不到工作表列表")
				}
				it, err := sb.StreamRows(names[0])
				mustNoErr(t, err)
				n := 0
				for it.Next() {
					n++
				}
				mustNoErr(t, it.Error())
				it.Close()
				if n == 0 {
					t.Error("流式迭代到 0 行")
				}
			})
		})
	}
}

// comboFixtureDir 跑 probe_combos.py 生成组合输入，返回所在目录。
func comboFixtureDir(t *testing.T) string {
	t.Helper()
	dir := t.TempDir()
	script := filepath.Join(difftestDir(t), "probe_combos.py")
	if o, err := execCommand(pythonExe(), script).CombinedOutput(); err != nil {
		t.Fatalf("生成组合输入失败: %v\n%s", err, o)
	} else if !strings.Contains(string(o), "OK") {
		t.Fatalf("生成组合输入未全部成功: %s", o)
	}
	// 探针脚本把产物写到 difftest/ 的**父目录**（库根），
	// 这样它也能被手工直接运行查看 —— 测试里再搬进临时目录。
	src := filepath.Dir(difftestDir(t))
	moved := 0
	for i := 1; i <= 6; i++ {
		name := "probe_combo" + strconv.Itoa(i) + ".xlsx"
		from := filepath.Join(src, name)
		if _, err := os.Stat(from); err != nil {
			continue
		}
		to := filepath.Join(dir, name)
		copyFileForTest(t, from, to)
		moved++
	}
	if moved == 0 {
		t.Fatal("未找到任何组合输入文件")
	}
	return dir
}

// assertTableIdentityUnique 断言所有表格部件的 id 与 name 工作簿级唯一。
//
// 这是"复制带表格的表 → 整个文件打不开"的直接守卫：id 与 name
// 重复会让 openpyxl 抛 "Table with name X already exists"。
func assertTableIdentityUnique(t *testing.T, p string) {
	t.Helper()
	seenID := map[string]string{}
	seenName := map[string]string{}
	seenDisplay := map[string]string{}
	for _, name := range zipPartNamesOf(p) {
		if !strings.HasPrefix(name, "xl/tables/") ||
			!strings.HasSuffix(name, ".xml") {
			continue
		}
		data := readZipPart(t, p, name)
		id := xmlAttrOf(data, "id")
		n := xmlAttrOf(data, "name")
		d := xmlAttrOf(data, "displayName")
		if prev, dup := seenID[id]; dup && id != "" {
			t.Errorf("表格 id 重复：%s 与 %s 都是 %q", prev, name, id)
		}
		seenID[id] = name
		if prev, dup := seenName[n]; dup && n != "" {
			t.Errorf("表格 name 重复：%s 与 %s 都是 %q"+
				"（openpyxl 判定重名用的是 name）", prev, name, n)
		}
		seenName[n] = name
		if prev, dup := seenDisplay[d]; dup && d != "" {
			t.Errorf("表格 displayName 重复：%s 与 %s 都是 %q", prev, name, d)
		}
		seenDisplay[d] = name
	}
}

// assertCondFmtSurvived 断言指定工作表的条件格式仍然存在且规则数 > 0。
func assertCondFmtSurvived(t *testing.T, p, sheet, desc string) {
	t.Helper()
	snap := runSnapshot(t, p)
	sheets, _ := snap["sheets"].(map[string]interface{})
	ws, _ := sheets[sheet].(map[string]interface{})
	if ws == nil {
		t.Errorf("%s: 快照缺少工作表 %s", desc, sheet)
		return
	}
	cf, _ := ws["conditional_formats"].([]interface{})
	if len(cf) == 0 {
		t.Errorf("%s: 工作表 %s 的条件格式丢失（源文件里存在）", desc, sheet)
	}
}

// sheetNamesOf 读出文件里的工作表名列表。
func sheetNamesOf(t *testing.T, p string) []string {
	t.Helper()
	g, err := Open(p)
	if err != nil {
		t.Fatalf("打开 %s 失败: %v", p, err)
	}
	return g.SheetNames()
}
