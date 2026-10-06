package excelgo

// 差分测试：同一组文档、同一组操作，分别用 excelgo 与 openpyxl 执行，
// 再把两侧产物归一化为语义快照并逐字段比对。
//
// 为什么需要这个：
// 既有测试大多验证"本库自己的输出符合本库的预期"，若预期本身写错就查不出来。
// openpyxl 是独立实现（且被 Excel 生态广泛验证），把它当**外部 oracle** 能抓出
// 两类问题：
//   1. 本库产出了 openpyxl/Excel 读不出或读成别的值的东西（真 bug）；
//   2. 本库读不出 openpyxl 正常产物（兼容性缺口）。
//
// 关键设计——区分"真差异"与"能力差异"：
// openpyxl 有一些能力本库没有（如图表、透视表），也有本库刻意不同的行为。
// 这些若不显式标注，会淹没真差异。所以每个场景可带 Notes 声明已知的能力差异，
// 差分时按"字段路径"精确豁免，而不是整场景跳过。

import (
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strings"
	"testing"
)

// pythonExe 定位 Python 解释器。找不到时相关测试自动跳过而非失败 ——
// 交叉校验是增强手段，不应成为硬依赖。
func pythonExe() string {
	candidates := []string{
		`C:\Users\null\.workbuddy\binaries\python\envs\validate\Scripts\python.exe`,
		`C:\Users\null\.workbuddy\binaries\python\envs\default\Scripts\python.exe`,
	}
	if p := os.Getenv("EXCELGO_PYTHON"); p != "" {
		candidates = append([]string{p}, candidates...)
	}
	for _, c := range candidates {
		if _, err := os.Stat(c); err == nil {
			return c
		}
	}
	return ""
}

func difftestDir(t *testing.T) string {
	t.Helper()
	wd, err := os.Getwd()
	if err != nil {
		t.Fatal(err)
	}
	return filepath.Join(wd, "difftest")
}

// runSnapshot 调 Python 快照工具，返回归一化后的语义表示。
func runSnapshot(t *testing.T, xlsx string) map[string]interface{} {
	t.Helper()
	py := pythonExe()
	if py == "" {
		t.Skipf("未找到 Python/openpyxl，跳过差分测试")
	}
	script := filepath.Join(difftestDir(t), "snapshot.py")
	// 必须分离 stdout / stderr：快照 JSON 走 stdout，openpyxl 的诊断警告走 stderr。
	// 早先用 CombinedOutput 把两者混在一起，导致警告被当成快照内容而误判失败。
	cmd := exec.Command(py, script, xlsx)
	var stdout, stderr strings.Builder
	cmd.Stdout = &stdout
	cmd.Stderr = &stderr
	if err := cmd.Run(); err != nil {
		t.Fatalf("快照失败(%s): %v\nstderr: %s", xlsx, err, stderr.String())
	}
	line := strings.TrimSpace(stdout.String())
	if !strings.HasPrefix(line, "SNAPSHOT_OK ") {
		t.Fatalf("快照未返回 OK(%s): %s\nstderr: %s", xlsx, line, stderr.String())
	}
	if w := strings.TrimSpace(stderr.String()); w != "" {
		t.Logf("openpyxl 诊断(%s): %s", filepath.Base(xlsx), w)
	}
	var m map[string]interface{}
	if err := json.Unmarshal([]byte(strings.TrimPrefix(line, "SNAPSHOT_OK ")), &m); err != nil {
		t.Fatalf("快照 JSON 解析失败: %v", err)
	}
	return m
}

// runOpenpyxlOps 让 openpyxl 执行一组操作并产出文件。
func runOpenpyxlOps(t *testing.T, outFile string, opsJSON string) {
	t.Helper()
	py := pythonExe()
	if py == "" {
		t.Skipf("未找到 Python/openpyxl，跳过差分测试")
	}
	script := filepath.Join(difftestDir(t), "op.py")
	cmd := exec.Command(py, script)
	cmd.Stdin = strings.NewReader(opsJSON)
	if out, err := cmd.CombinedOutput(); err != nil {
		t.Fatalf("openpyxl 执行失败: %v\n%s", err, out)
	}
	_ = outFile
}

// diffSnapshot 比较两份快照，返回人类可读的差异描述。
// exempts 里的字段路径会被跳过（用于声明已知的"能力差异"）。
func diffSnapshot(path string, got, want interface{}, exempts map[string]bool) []string {
	var out []string
	compare(path, got, want, exempts, &out)
	sort.Strings(out)
	return out
}

func compare(path string, got, want interface{}, exempts map[string]bool, out *[]string) {
	if exempts[path] {
		return
	}
	switch w := want.(type) {
	case map[string]interface{}:
		g, ok := got.(map[string]interface{})
		if !ok {
			*out = append(*out, fmt.Sprintf("%s: 类型不同 got=%T want=%T", path, got, want))
			return
		}
		keys := map[string]bool{}
		for k := range g {
			keys[k] = true
		}
		for k := range w {
			keys[k] = true
		}
		var sorted []string
		for k := range keys {
			sorted = append(sorted, k)
		}
		sort.Strings(sorted)
		for _, k := range sorted {
			compare(path+"."+k, g[k], w[k], exempts, out)
		}
	case []interface{}:
		g, ok := got.([]interface{})
		if !ok {
			*out = append(*out, fmt.Sprintf("%s: 类型不同 got=%T want=%T", path, got, want))
			return
		}
		if len(g) != len(w) {
			*out = append(*out, fmt.Sprintf("%s: 长度不同 got=%d(%v) want=%d(%v)",
				path, len(g), g, len(w), w))
			return
		}
		for i := range g {
			compare(fmt.Sprintf("%s[%d]", path, i), g[i], w[i], exempts, out)
		}
	default:
		// 标量：数值按容差比较，避免浮点抖动造成假差异
		gf, wok := toFloat(got)
		wf, wok2 := toFloat(want)
		if wok && wok2 {
			if !nearlyEqual(gf, wf) {
				*out = append(*out, fmt.Sprintf("%s: got=%v want=%v", path, got, want))
			}
			return
		}
		if !scalarEqual(got, want) {
			*out = append(*out, fmt.Sprintf("%s: got=%v want=%v", path, got, want))
		}
	}
}

func toFloat(v interface{}) (float64, bool) {
	switch n := v.(type) {
	case float64:
		return n, true
	case int:
		return float64(n), true
	case int64:
		return float64(n), true
	}
	return 0, false
}

func nearlyEqual(a, b float64) bool {
	d := a - b
	if d < 0 {
		d = -d
	}
	// 列宽/行高在两侧经过 XML 与 UI 往返后常有 1e-3 级偏差
	return d < 0.01
}

func scalarEqual(a, b interface{}) bool {
	// JSON 解出来都是 float64/string/bool/nil，直接比
	as, aok := a.(string)
	bs, bok := b.(string)
	if aok && bok {
		return as == bs
	}
	ab, aok2 := a.(bool)
	bb, bok2 := b.(bool)
	if aok2 && bok2 {
		return ab == bb
	}
	if a == nil && b == nil {
		return true
	}
	af, aok3 := toFloat(a)
	bf, bok3 := toFloat(b)
	if aok3 && bok3 {
		return af == bf
	}
	return fmt.Sprint(a) == fmt.Sprint(b)
}

// difftestCase 一个差分场景：两侧执行同一组操作，产物必须语义一致。
type difftestCase struct {
	Name string
	// ExcelGo 在空文件上执行的操作（返回值即 excelgo 调用）
	ExcelGo func(t *testing.T, f string)
	// Ops 是发给 openpyxl 的等价操作（JSON 片段）
	Ops string
	// Exempts 声明已知的"能力差异"字段路径（真差异不在此列）
	Exempts []string
	// SkipReason 非空则跳过（说明为何不可比）
	SkipReason string
}

// runDiffCase 执行一个差分场景并比对两侧快照。
func runDiffCase(t *testing.T, c difftestCase) {
	t.Helper()
	if c.SkipReason != "" {
		t.Skipf("不可比：%s", c.SkipReason)
	}
	dir := t.TempDir()
	gotFile := filepath.Join(dir, "excelgo.xlsx")
	wantFile := filepath.Join(dir, "openpyxl.xlsx")

	// excelgo 侧
	if err := Create2(gotFile); err != nil {
		t.Fatalf("创建 excelgo 基准文件失败: %v", err)
	}
	c.ExcelGo(t, gotFile)

	// openpyxl 侧
	ops := fmt.Sprintf(`{"out": %q, "ops": [%s]}`, wantFile, c.Ops)
	runOpenpyxlOps(t, wantFile, ops)

	got := runSnapshot(t, gotFile)
	want := runSnapshot(t, wantFile)

	exempts := map[string]bool{}
	for _, e := range c.Exempts {
		exempts[e] = true
	}
	if diffs := diffSnapshot("", got, want, exempts); len(diffs) > 0 {
		for _, d := range diffs {
			t.Errorf("[%s] %s", c.Name, d)
		}
	}
}

// Create2 创建一个只含默认空表的 xlsx（差分测试的共同起点）。
func Create2(path string) error {
	b, err := Create()
	if err != nil {
		return err
	}
	return b.SaveAs(path)
}

// execCommand 是 exec.Command 的薄包装，让测试文件不必各自 import os/exec。
func execCommand(name string, args ...string) *exec.Cmd {
	return exec.Command(name, args...)
}
