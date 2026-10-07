package excelgo

import (
	"fmt"
	"testing"
	"time"
)

func TestSSTPerf(t *testing.T) {
	p := genPath("zz_sst.xlsx")
	b0, _ := Create()
	_ = b0.SaveAs(p)
	f, _ := Open(p)
	sh, _ := f.Sheet("Sheet1")

	start := time.Now()
	for i := 0; i < 2000; i++ {
		_ = sh.SetCellValue(fmt.Sprintf("A%d", i+1), fmt.Sprintf("value-%d", i))
	}
	d := time.Since(start)
	fmt.Printf("2000 次 SetCellValue(string) 耗时: %v\n", d)
	_ = f.Save()

	// 校验值正确
	f2, _ := Open(p)
	sh2, _ := f2.Sheet("Sheet1")
	for _, tc := range []struct{ ref, want string }{{"A1", "value-0"}, {"A1000", "value-999"}, {"A2000", "value-1999"}} {
		got, err := sh2.GetCellValue(tc.ref)
		if err != nil {
			t.Fatalf("%s: %v", tc.ref, err)
		}
		if fmt.Sprint(got) != tc.want {
			t.Errorf("%s = %v, want %s", tc.ref, got, tc.want)
		}
	}

	// 去重：重复写同一字符串应复用索引
	fm := readMap(t, p)
	sst := string(fm["xl/sharedStrings.xml"])
	n := len(extractSharedStrings(sst))
	fmt.Printf("SST 条目数 = %d (2000 格全唯一，应为 2000)\n", n)
	if n != 2000 {
		t.Errorf("SST 条目数 = %d, 期望 2000", n)
	}
	if !sstCountAttrRe.MatchString(sst) {
		t.Error("count/uniqueCount 属性缺失或格式异常")
	}
}

func TestSSTDedupe(t *testing.T) {
	p := genPath("zz_sst_dedup.xlsx")
	b0, _ := Create()
	_ = b0.SaveAs(p)
	f, _ := Open(p)
	sh, _ := f.Sheet("Sheet1")
	for i := 0; i < 100; i++ {
		_ = sh.SetCellValue(fmt.Sprintf("A%d", i+1), "重复值")
	}
	_ = f.Save()
	fm := readMap(t, p)
	sst := string(fm["xl/sharedStrings.xml"])
	n := len(extractSharedStrings(sst))
	fmt.Printf("100 格同值 -> SST 条目数 = %d (应为 1)\n", n)
	if n != 1 {
		t.Errorf("SST 条目数 = %d, 期望 1", n)
	}
}

func TestSSTCacheInvalidate(t *testing.T) {
	// 缓存不得因外部改动部件而给出错误索引
	fm := map[string][]byte{
		"xl/sharedStrings.xml": []byte(buildSharedStrings([]string{"a", "b"})),
	}
	sc := &sharedStringsCache{}
	// 首次：慢路径建索引（部件已存在）
	if idx, _ := ensureSharedString(fm, "c", sc); idx != 2 {
		t.Fatalf("期望索引 2，实际 %d", idx)
	}
	// 外部重置部件，模拟 CopySheet/合并重写
	fm["xl/sharedStrings.xml"] = []byte(buildSharedStrings([]string{"x", "y"}))
	// 指纹失配 -> 必须重解析，"x" 应得索引 0 而非旧缓存里的结果
	idx, _ := ensureSharedString(fm, "x", sc)
	if idx != 0 {
		t.Errorf("缓存未失效：期望索引 0，实际 %d", idx)
	}
}

// TestSSTEmptyCacheNilMapRegression 回归：裸 &sharedStringsCache{} 的 index 为 nil，
// 且部件尚不存在时 src 与 data 同为 nil，bytes.Equal(nil,nil)==true —— 若快速路径
// 不校验 data != nil / index != nil，会在写 index 时 panic。
// 真实触发路径：Open() 一份不含 sharedStrings.xml 的文件（openpyxl 产物常见形态）后写字符串。
func TestSSTEmptyCacheNilMapRegression(t *testing.T) {
	p := genPath("zz_sst_nilmap.xlsx")
	b0, _ := Create()
	_ = b0.SaveAs(p)

	f, err := Open(p)
	if err != nil {
		t.Fatal(err)
	}
	sh, err := f.NewSheet("S") // 走 NewSheet 而非 Sheet，覆盖另一条构造路径
	if err != nil {
		t.Fatal(err)
	}
	for _, ref := range []string{"A1", "A4", "B2", "B3", "G7"} {
		if err := sh.SetCellValue(ref, "val-"+ref); err != nil {
			t.Fatalf("写%s: %v", ref, err)
		}
	}
	if err := f.Save(); err != nil {
		t.Fatal(err)
	}
	fm := readMap(t, p)
	if _, ok := fm["xl/sharedStrings.xml"]; !ok {
		t.Error("应生成 sharedStrings.xml")
	}
	// 值可读回
	f2, _ := Open(p)
	sh2, _ := f2.Sheet("S")
	for _, ref := range []string{"A1", "A4", "B2", "B3", "G7"} {
		got, err := sh2.GetCellValue(ref)
		if err != nil {
			t.Fatalf("读%s: %v", ref, err)
		}
		if got != "val-"+ref {
			t.Errorf("%s = %v, want val-%s", ref, got, ref)
		}
	}
}

// TestSSTCacheHitExistingFirstWrite 首次写入即命中已有项时应建缓存并可继续快速追加。
func TestSSTCacheHitExistingFirstWrite(t *testing.T) {
	fm := map[string][]byte{
		"xl/sharedStrings.xml": []byte(buildSharedStrings([]string{"dup", "other"})),
	}
	sc := &sharedStringsCache{}
	// 第一次写已存在的值：慢路径 + 建缓存
	if idx, _ := ensureSharedString(fm, "dup", sc); idx != 0 {
		t.Fatalf("期望索引 0，实际 %d", idx)
	}
	if sc.index == nil {
		t.Fatal("命中已有项后仍应建立索引缓存")
	}
	// 第二次：新值应走快速路径（部件字节未变）
	before := len(sc.items)
	if idx, _ := ensureSharedString(fm, "new", sc); idx != 2 {
		t.Fatalf("期望索引 2，实际 %d", idx)
	}
	if len(sc.items) != before+1 {
		t.Error("缓存未在快速路径累加")
	}
}
