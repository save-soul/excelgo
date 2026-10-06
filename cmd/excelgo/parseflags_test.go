package main

import (
	"flag"
	"strings"
	"testing"
	"time"
)

// newTestFS 构造一个与实际子命令同构的 FlagSet：
// name/suffix 为字符串 flag，independent/bold 为布尔 flag。
func newTestFS() (*flag.FlagSet, *string, *string, *bool) {
	fs := flag.NewFlagSet("t", flag.ContinueOnError)
	fs.SetOutput(&strings.Builder{})
	name := fs.String("name", "", "")
	suffix := fs.String("suffix", "_copy", "")
	ind := fs.Bool("independent", false, "")
	return fs, name, suffix, ind
}

func TestParseFlagsOrders(t *testing.T) {
	cases := []struct {
		desc     string
		args     []string
		wantPos  []string
		wantName string
		wantSuf  string
		wantInd  bool
	}{
		{
			desc:    "flags 在后",
			args:    []string{"a.xlsx", "b.xlsx", "Sheet1", "--name", "汇总表"},
			wantPos: []string{"a.xlsx", "b.xlsx", "Sheet1"}, wantName: "汇总表",
		},
		{
			desc:    "flags 在前",
			args:    []string{"--name", "汇总表", "a.xlsx", "b.xlsx", "Sheet1"},
			wantPos: []string{"a.xlsx", "b.xlsx", "Sheet1"}, wantName: "汇总表",
		},
		{
			desc:    "等号形式",
			args:    []string{"a.xlsx", "b.xlsx", "--name=汇总表"},
			wantPos: []string{"a.xlsx", "b.xlsx"}, wantName: "汇总表",
		},
		{
			desc:    "单横线",
			args:    []string{"a.xlsx", "-name", "x"},
			wantPos: []string{"a.xlsx"}, wantName: "x",
		},
		{
			desc:    "布尔 flag 不吞值",
			args:    []string{"a.xlsx", "--independent", "b.xlsx"},
			wantPos: []string{"a.xlsx", "b.xlsx"}, wantInd: true,
		},
		{
			desc:    "布尔 flag 显式 =false",
			args:    []string{"a.xlsx", "--independent=false"},
			wantPos: []string{"a.xlsx"}, wantInd: false,
		},
		{
			desc:    "多 flag 混排",
			args:    []string{"a.xlsx", "--name", "N", "--suffix", "_s", "--independent", "b.xlsx"},
			wantPos: []string{"a.xlsx", "b.xlsx"}, wantName: "N", wantSuf: "_s", wantInd: true,
		},
		{
			desc:    "-- 之后全为位置参数",
			args:    []string{"a.xlsx", "--", "--name", "x"},
			wantPos: []string{"a.xlsx", "--name", "x"},
		},
		{
			desc:    "无 flag",
			args:    []string{"a.xlsx", "b.xlsx"},
			wantPos: []string{"a.xlsx", "b.xlsx"},
		},
		{
			desc:    "空参数",
			args:    nil,
			wantPos: nil,
		},
		{
			desc:    "负数位置参数（行号可为负？至少不应被当 flag 吞掉）",
			args:    []string{"a.xlsx", "Sheet1", "-1"},
			wantPos: []string{"a.xlsx", "Sheet1", "-1"},
		},
		{
			desc:    "负数 flag 值",
			args:    []string{"a.xlsx", "--name", "-5"},
			wantPos: []string{"a.xlsx"}, wantName: "-5",
		},
		{
			desc:    "值以横线开头",
			args:    []string{"a.xlsx", "--suffix", "-副本"},
			wantPos: []string{"a.xlsx"}, wantSuf: "-副本",
		},
		{
			desc:    "空字符串位置参数",
			args:    []string{"", "a.xlsx"},
			wantPos: []string{"", "a.xlsx"},
		},
		{
			desc:    "单个横线视为位置参数",
			args:    []string{"-", "a.xlsx"},
			wantPos: []string{"-", "a.xlsx"},
		},
	}

	for _, c := range cases {
		// 未显式传 --suffix 时应保持 FlagSet 的默认值 "_copy"
		wantSuf := c.wantSuf
		if wantSuf == "" {
			wantSuf = "_copy"
		}
		t.Run(c.desc, func(t *testing.T) {
			fs, name, suffix, ind := newTestFS()
			if err := parseFlags(fs, c.args); err != nil {
				t.Fatalf("parseFlags 报错: %v", err)
			}
			got := fs.Args()
			if strings.Join(got, "\x00") != strings.Join(c.wantPos, "\x00") {
				t.Errorf("位置参数 = %q，期望 %q", got, c.wantPos)
			}
			if *name != c.wantName {
				t.Errorf("--name = %q，期望 %q", *name, c.wantName)
			}
			if *suffix != wantSuf {
				t.Errorf("--suffix = %q，期望 %q", *suffix, wantSuf)
			}
			if *ind != c.wantInd {
				t.Errorf("--independent = %v，期望 %v", *ind, c.wantInd)
			}
		})
	}
}

// TestParseFlagsUnknownFlag 未注册的双横线 token（疑似打错字）应报错，
// 而不是被静默当作位置参数让用户以为生效了。
func TestParseFlagsUnknownFlag(t *testing.T) {
	fs, _, _, _ := newTestFS()
	err := parseFlags(fs, []string{"a.xlsx", "--typo", "x"})
	if err == nil {
		t.Error("未注册的双横线 token 应返回错误（提示用户打错了），而非静默吞掉")
	}
}

// TestParseFlagsSingleDashUnknownKeptAsPositional 单横线且未注册的 token
// （如 -5、-副本）保留为位置参数，避免 setcell A1 -5 这类合法负数值报错。
func TestParseFlagsSingleDashUnknownKeptAsPositional(t *testing.T) {
	fs, _, _, _ := newTestFS()
	args := []string{"a.xlsx", "Sheet1", "-5"}
	if err := parseFlags(fs, args); err != nil {
		t.Fatalf("单横线负数不应报错: %v", err)
	}
	got := fs.Args()
	if strings.Join(got, "\x00") != strings.Join(args, "\x00") {
		t.Errorf("位置参数 = %q，期望原样保留 %q", got, args)
	}
}

// TestParseFlagsDuplicateFlag 重复传同一 flag 时以最后一个为准（标准库语义）。
func TestParseFlagsDuplicateFlag(t *testing.T) {
	fs, name, _, _ := newTestFS()
	if err := parseFlags(fs, []string{"a.xlsx", "--name", "first", "--name", "second"}); err != nil {
		t.Fatalf("parseFlags: %v", err)
	}
	if *name != "second" {
		t.Errorf("--name = %q，期望 second（最后一个生效）", *name)
	}
}

// TestParseFlagsSetcellStyle 锁定 setcell 子命令的 flag 组合：
// flag 在前/在后、负数位置参数、布尔式字符串值都必须正确解析。
// （旧实现用独立的 extractFlags，在 --type 缺值时死循环，此处一并回归。）
func TestParseFlagsSetcellStyle(t *testing.T) {
	newFS := func() (*flag.FlagSet, *string, *string) {
		fs := flag.NewFlagSet("setcell", flag.ContinueOnError)
		fs.SetOutput(&strings.Builder{})
		typ := fs.String("type", "str", "")
		result := fs.String("result", "", "")
		return fs, typ, result
	}
	cases := []struct {
		desc       string
		args       []string
		wantPos    []string
		wantType   string
		wantResult string
	}{
		{
			desc:     "flag 在后",
			args:     []string{"f.xlsx", "Base", "A1", "你好", "--type", "str"},
			wantPos:  []string{"f.xlsx", "Base", "A1", "你好"},
			wantType: "str",
		},
		{
			desc:     "flag 在前",
			args:     []string{"--type", "num", "f.xlsx", "Base", "A1", "42"},
			wantPos:  []string{"f.xlsx", "Base", "A1", "42"},
			wantType: "num",
		},
		{
			desc:     "负数单元格值不被当作 flag",
			args:     []string{"f.xlsx", "Base", "A1", "-3.5", "--type", "num"},
			wantPos:  []string{"f.xlsx", "Base", "A1", "-3.5"},
			wantType: "num",
		},
		{
			desc:     "布尔式字符串值不被当作 flag",
			args:     []string{"f.xlsx", "Base", "A1", "true", "--type", "bool"},
			wantPos:  []string{"f.xlsx", "Base", "A1", "true"},
			wantType: "bool",
		},
		{
			desc:     "formula + result",
			args:     []string{"f.xlsx", "Base", "A4", "1+2", "--type", "formula", "--result", "3"},
			wantPos:  []string{"f.xlsx", "Base", "A4", "1+2"},
			wantType: "formula", wantResult: "3",
		},
		{
			desc:     "等号形式",
			args:     []string{"f.xlsx", "Base", "A1", "x", "--type=bool"},
			wantPos:  []string{"f.xlsx", "Base", "A1", "x"},
			wantType: "bool",
		},
		{
			desc:     "无 flag 用默认 str",
			args:     []string{"f.xlsx", "Base", "A1", "x"},
			wantPos:  []string{"f.xlsx", "Base", "A1", "x"},
			wantType: "str",
		},
	}
	for _, c := range cases {
		t.Run(c.desc, func(t *testing.T) {
			fs, typ, result := newFS()
			if err := parseFlags(fs, c.args); err != nil {
				t.Fatalf("parseFlags 报错: %v", err)
			}
			if got := fs.Args(); strings.Join(got, "\x00") != strings.Join(c.wantPos, "\x00") {
				t.Errorf("位置参数 = %q，期望 %q", got, c.wantPos)
			}
			if *typ != c.wantType {
				t.Errorf("--type = %q，期望 %q", *typ, c.wantType)
			}
			if *result != c.wantResult {
				t.Errorf("--result = %q，期望 %q", *result, c.wantResult)
			}
		})
	}
}

// TestParseFlagsNoInfiniteLoop 缺值 flag 必须立即返回错误，绝不能挂死。
// 回归：旧 extractFlags 在 `--type` 缺值时 switch 落空、i++ 不执行，CLI 永久挂起。
func TestParseFlagsNoInfiniteLoop(t *testing.T) {
	done := make(chan error, 1)
	go func() {
		fs := flag.NewFlagSet("setcell", flag.ContinueOnError)
		fs.SetOutput(&strings.Builder{})
		fs.String("type", "str", "")
		fs.String("result", "", "")
		done <- parseFlags(fs, []string{"f.xlsx", "Base", "A1", "x", "--type"})
	}()
	select {
	case err := <-done:
		if err == nil {
			t.Error("--type 缺值应返回错误")
		}
	case <-time.After(3 * time.Second):
		t.Fatal("parseFlags 在 flag 缺值时挂死（无限循环）")
	}
}
