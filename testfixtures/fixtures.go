// Package testfixtures 内嵌 ExcelGo 的测试夹具，使其在测试运行时由 TestMain 物化到
// ../testdata_tmp/，从而无需任何外部样本文件或 Python 生成脚本即可运行 go test。
//
// 夹具最初由 openpyxl 生成（与库的非破坏式编辑约定一致），此处仅作为二进制资源打包，
// 不引入任何运行时依赖。
package testfixtures

import "embed"

//go:embed full.xlsx grid.xlsx target.xlsx src1.xlsx src2.xlsx ph_red.png ph_blue.png
var assets embed.FS

// Names 返回所有内嵌夹具文件名。
func Names() []string {
	return []string{
		"full.xlsx", "grid.xlsx", "target.xlsx", "src1.xlsx", "src2.xlsx",
		"ph_red.png", "ph_blue.png",
	}
}

// Bytes 返回指定名称夹具的原始字节。
func Bytes(name string) ([]byte, error) {
	return assets.ReadFile(name)
}
