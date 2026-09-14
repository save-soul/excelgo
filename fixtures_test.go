package excelgo

import (
	"os"
	"path/filepath"
	"testing"

	"github.com/save-soul/excelgo/testfixtures"
)

// TestMain 在运行任何测试前，将内嵌的测试夹具物化到仓库根的 ../testdata_tmp/ 目录。
// 这样无论本机是否存在样本文件、是否安装 Python/openpyxl，go test 都能直接跑绿。
func TestMain(m *testing.M) {
	base := filepath.Join("..", "testdata_tmp")
	if err := os.MkdirAll(base, 0o755); err != nil {
		panic("create testdata_tmp: " + err.Error())
	}
	for _, name := range testfixtures.Names() {
		data, err := testfixtures.Bytes(name)
		if err != nil {
			panic("missing embedded fixture: " + name)
		}
		if err := os.WriteFile(filepath.Join(base, name), data, 0o644); err != nil {
			panic("write fixture " + name + ": " + err.Error())
		}
	}
	os.Exit(m.Run())
}
