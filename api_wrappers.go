package excelgo

// api_wrappers.go 为 sheet_extra.go 中的对象方法提供「文件名 + 表名」形式的全局函数包装，
// 便于不引入对象 API 的场景或 CLI 直接使用。每个包装都遵循「读盘 -> 修改 -> 写回」，
// 与全库其余全局函数一致。

// withSheet 打开文件、定位工作表、执行 fn、保存。供多数只写操作复用。
func withSheet(filename, sheetRef string, fn func(*WorkSheet) error) error {
	b, err := Open(filename)
	if err != nil {
		return err
	}
	s, err := b.Sheet(sheetRef)
	if err != nil {
		return err
	}
	if err := fn(s); err != nil {
		return err
	}
	return b.Save()
}

// MergeCells 合并 ref 区域。
func MergeCells(filename, sheetRef, ref string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.MergeCells(ref) })
}

// UnmergeCells 取消 ref 区域合并。
func UnmergeCells(filename, sheetRef, ref string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.UnmergeCells(ref) })
}

// SetColWidth 设置单列宽度。
func SetColWidth(filename, sheetRef string, col int, width float64) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.SetColWidth(col, width) })
}

// SetColWidthRange 设置连续列宽度。
func SetColWidthRange(filename, sheetRef string, min, max int, width float64) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.SetColWidthRange(min, max, width) })
}

// GetColWidth 读取单列自定义宽度（未设置返回 0）。
func GetColWidth(filename, sheetRef string, col int) (float64, error) {
	b, err := Open(filename)
	if err != nil {
		return 0, err
	}
	s, err := b.Sheet(sheetRef)
	if err != nil {
		return 0, err
	}
	return s.GetColWidth(col), nil
}

// SetRowHeight 设置行高。
func SetRowHeight(filename, sheetRef string, row int, height float64) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.SetRowHeight(row, height) })
}

// SetRowVisible 设置行可见性。
func SetRowVisible(filename, sheetRef string, row int, visible bool) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.SetRowVisible(row, visible) })
}

// SetColVisible 设置列可见性。
func SetColVisible(filename, sheetRef string, col int, visible bool) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.SetColVisible(col, visible) })
}

// FreezePanes 冻结窗格。
func FreezePanes(filename, sheetRef, ref string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.FreezePanes(ref) })
}

// AddHyperlink 添加超链接。
func AddHyperlink(filename, sheetRef, cell, url, displayText string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.AddHyperlink(cell, url, displayText) })
}

// AutoFilter 添加自动筛选。
func AutoFilter(filename, sheetRef, ref string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.AutoFilter(ref) })
}

// AddDataValidation 添加数据验证。
func AddDataValidation(filename, sheetRef, ref, typ, op, formula1, formula2 string, allowBlank bool) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.AddDataValidation(ref, typ, op, formula1, formula2, allowBlank)
	})
}

// AddDataValidationList 添加下拉列表数据验证。
func AddDataValidationList(filename, sheetRef, ref string, values []string, allowBlank bool) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.AddDataValidationList(ref, values, allowBlank)
	})
}

// SetConditionalFormat 添加条件格式。
func SetConditionalFormat(filename, sheetRef, ref, cfType, formula string, priority int, st Style) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error {
		return s.SetConditionalFormat(ref, cfType, formula, priority, st)
	})
}

// AddTable 创建结构化表格。
func AddTable(filename, sheetRef, ref, name string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.AddTable(ref, name) })
}

// AddComment 添加批注。
func AddComment(filename, sheetRef, cell, text, author string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.AddComment(cell, text, author) })
}

// GetComment 读取批注。
func GetComment(filename, sheetRef, cell string) (string, error) {
	b, err := Open(filename)
	if err != nil {
		return "", err
	}
	s, err := b.Sheet(sheetRef)
	if err != nil {
		return "", err
	}
	return s.GetComment(cell), nil
}

// SetSheetProps 设置工作表属性。
func SetSheetProps(filename, sheetRef string, opts SheetProps) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.SetProps(opts) })
}

// ProtectSheet 保护工作表。
func ProtectSheet(filename, sheetRef, password string) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.Protect(password) })
}

// GroupRows 行分组大纲。
func GroupRows(filename, sheetRef string, r1, r2, level int) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.GroupRows(r1, r2, level) })
}

// GroupCols 列分组大纲。
func GroupCols(filename, sheetRef string, c1, c2, level int) error {
	return withSheet(filename, sheetRef, func(s *WorkSheet) error { return s.GroupCols(c1, c2, level) })
}

// ReplaceText 工作表内查找替换。
func ReplaceText(filename, sheetRef, old, new string) (int, error) {
	b, err := Open(filename)
	if err != nil {
		return 0, err
	}
	s, err := b.Sheet(sheetRef)
	if err != nil {
		return 0, err
	}
	n, err := s.Replace(old, new)
	if err != nil {
		return 0, err
	}
	if err := b.Save(); err != nil {
		return 0, err
	}
	return n, nil
}

// SetDocProps 设置文档核心属性。
func SetDocProps(filename string, props map[string]string) error {
	b, err := Open(filename)
	if err != nil {
		return err
	}
	if err := b.SetDocProps(props); err != nil {
		return err
	}
	return b.Save()
}
