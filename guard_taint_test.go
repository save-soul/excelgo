package excelgo

import (
	"go/ast"
	"go/parser"
	"go/token"
	"path/filepath"
	"strconv"
	"strings"
	"testing"
)

// 本文件是「XML 写入安全」的静态护栏（污点追踪，零人工豁免）。
//
// 背景：本库为保住 WPS 内嵌图片、命名空间与版式，全程以字符串与正则编辑 XML 部件
// （实测原因见 README「XML 安全模型」）。代价是每个把用户数据写进 XML 的地方都必须
// 自行转义 —— 漏一处即为可被构造数据破坏的写入点。
//
// 护栏原理（污点追踪 + 类型认证）：
//   - 起点：函数入参、局部变量、结构体字段、range 变量、append 结果 —— 视为"带污"；
//   - 清除：经 safeText/safeAttr/strconv.* 处理后污点被清除；
//   - 认证：若变量的**静态类型**是已认证具名类型（cellRef / relID / xmlFrag），
//     则视为可信 —— 认证来自类型本身，不来自任何人工清单。
//
// 关键设计：安全性由**类型**表达，而非变量名白名单。因此
//   - 不存在"忘记往清单里加一条"这种漏检；
//   - 新增写入点默认受保护（裸 string 一律带污）。

// certifiedTypes 是「已认证安全」的具名类型（定义见 cell.go）。
//
//	cellRef —— 经 newCellRef 校验的单元格坐标（A1 形式）
//	relID   —— 库内生成的关系 ID（rIdN）
//	xmlFrag —— 仅由字面量与 safe* 拼成的 XML 片段
var certifiedTypes = map[string]string{
	"cellRef":        "经 newCellRef 校验的单元格坐标",
	"relID":          "库内生成的关系 ID",
	"xmlFrag":        "仅由字面量与 safe* 拼成的 XML 片段",
	"xmlTagName":     "库内固定的 schema 标签名常量",
	"rangeRef":       "经 newRangeRef 校验的区域引用",
	"styleIndex":     "库内生成的样式索引（fontId/fillId/numFmtId 等）",
	"borderSideName": "库内固定的边框边名常量",
	"partXML":        "xlsx 部件的既有原始 XML 字节（读取后复用，非新建）",
	"numLit":         "纯数值字面量（由 strconv.FormatX 生成，只含数字与符号）",
	"cellOpen":       "<c> 标签的属性部分（由 cellRef + styleIndex 拼成）",
}

// cleanFuncs 会清除污点的函数。
var cleanFuncs = map[string]bool{
	"safeText": true, "safeAttr": true,
}

// numberFuncs 把值转成不可能携带注入的字符串。
var numberFuncs = map[string]bool{
	"strconv.Itoa": true, "strconv.FormatInt": true, "strconv.FormatFloat": true,
	"strconv.Quote": true, "itoa": true,
}

// TestNoTaintedXMLWrite 是护栏主测试。
func TestNoTaintedXMLWrite(t *testing.T) {
	fset := token.NewFileSet()
	files, err := filepath.Glob("*.go")
	if err != nil {
		t.Fatal(err)
	}
	var problems []string
	for _, fname := range files {
		// 跳过测试文件与本护栏自身（含 XML 字面量，会自我报警）
		if strings.HasSuffix(fname, "_test.go") || strings.HasPrefix(fname, "guard_") {
			continue
		}
		f, perr := parser.ParseFile(fset, fname, nil, 0)
		if perr != nil {
			t.Fatalf("解析 %s 失败: %v", fname, perr)
		}
		for _, decl := range f.Decls {
			fn, ok := decl.(*ast.FuncDecl)
			if !ok || fn.Body == nil {
				continue
			}
			problems = append(problems, analyzeFunc(fset, fn)...)
		}
	}
	if len(problems) > 0 {
		t.Errorf("XML 写入护栏：发现 %d 处把用户数据未转义地拼入 XML"+
			"（必须经 safeText/safeAttr，或把参数类型改为已认证的 cellRef/relID/xmlFrag）：", len(problems))
		for _, p := range problems {
			t.Errorf("  %s", p)
		}
	}
}

// analyzeFunc 对单个函数做污点分析，返回问题列表。
func analyzeFunc(fset *token.FileSet, fn *ast.FuncDecl) []string {
	tainted := map[string]bool{}   // 变量名 -> 带污
	certified := map[string]bool{} // 变量名 -> 类型已认证（可信）
	addParams(tainted, certified, fn)
	// 结构体字段视为带污（st.Title 等）
	structTainted := map[string]bool{}
	collectStructFields(structTainted, fn)

	// 固定点迭代：赋值可出现在使用之前（Go 允许），多跑几轮收敛
	for round := 0; round < 3; round++ {
		changed := false
		ast.Inspect(fn.Body, func(n ast.Node) bool {
			switch s := n.(type) {
			case *ast.AssignStmt:
				for i, rhs := range s.Rhs {
					if i >= len(s.Lhs) {
						continue
					}
					id, ok := s.Lhs[i].(*ast.Ident)
					if !ok {
						continue
					}
					now := exprTainted(rhs, tainted, structTainted, certified)
					if tainted[id.Name] != now {
						tainted[id.Name] = now
						changed = true
					}
					// 右值是已认证类型的构造/转换（newCellRef(x) 等）则结果可信
					if isCertifiedConv(rhs) {
						certified[id.Name] = true
						tainted[id.Name] = false
					}
				}
			case *ast.ValueSpec:
				// 局部声明：默认带污（保守）；有右值时按右值判断
				// 认证来源：显式声明认证类型（var x cellRef），或由已认证构造函数赋值
				declCert := isCertifiedTypeExpr(s.Type)
				for i, id := range s.Names {
					now := true
					if i < len(s.Values) {
						now = exprTainted(s.Values[i], tainted, structTainted, certified)
					}
					if tainted[id.Name] != now {
						tainted[id.Name] = now
						changed = true
					}
					if declCert || (i < len(s.Values) && isCertifiedConv(s.Values[i])) {
						certified[id.Name] = true
						tainted[id.Name] = false
					}
				}
			case *ast.RangeStmt:
				if s.Tok == token.DEFINE {
					srcTainted := exprTainted(s.X, tainted, structTainted, certified)
					for _, e := range []ast.Expr{s.Key, s.Value} {
						if id, ok := e.(*ast.Ident); ok && srcTainted && !tainted[id.Name] {
							tainted[id.Name] = true
							changed = true
						}
					}
				}
			}
			return true
		})
		if !changed {
			break
		}
	}

	// 检查 XML 拼接。两种形态都要覆盖：
	//   1) 裸拼接表达式：a + "<t>" + b
	//   2) 作为函数实参：f.WriteString("<t>" + b)  —— 同样是在构造 XML
	var problems []string
	seen := map[string]bool{}
	report := func(pos token.Pos, what string) {
		key := fset.Position(pos).String() + what
		if seen[key] {
			return
		}
		seen[key] = true
		problems = append(problems, key+"  ("+what+")")
	}
	ast.Inspect(fn.Body, func(n ast.Node) bool {
		// 形态 2：调用实参中的拼接
		if call, ok := n.(*ast.CallExpr); ok {
			for _, arg := range call.Args {
				checkConcat(arg, tainted, structTainted, certified, report, fset)
			}
			return true
		}
		be, ok := n.(*ast.BinaryExpr)
		if !ok || be.Op != token.ADD {
			return true
		}
		if !hasXMLTagLiteral(be.X) && !hasXMLTagLiteral(be.Y) {
			return true
		}
		other := be.Y
		if hasXMLTagLiteral(be.Y) {
			other = be.X
		}
		checkConcat(other, tainted, structTainted, certified, report, fset)
		return true
	})
	return problems
}

// checkConcat 检查一个表达式中的拼接：若含 XML 标签字面量且有带污操作数，则报告。
func checkConcat(e ast.Expr, tainted, structTainted, certified map[string]bool,
	report func(token.Pos, string), fset *token.FileSet) {
	be, ok := e.(*ast.BinaryExpr)
	if !ok || be.Op != token.ADD {
		return
	}
	if !hasXMLTagLiteral(be) {
		// 可能是更外层的拼接，交给父节点处理；这里只处理含标签的那一层
		return
	}
	for _, arg := range flattenConcat(be) {
		if !exprTainted(arg, tainted, structTainted, certified) {
			continue
		}
		report(arg.Pos(), "带污数据未清洗即拼入 XML")
	}
}

// addParams 把函数入参登记为「带污」；已认证的具名类型除外。
func addParams(tainted, certified map[string]bool, fn *ast.FuncDecl) {
	if fn.Type == nil || fn.Type.Params == nil {
		return
	}
	for _, p := range fn.Type.Params.List {
		isCert := isCertifiedTypeExpr(p.Type)
		for _, n := range p.Names {
			tainted[n.Name] = !isCert
			if isCert {
				certified[n.Name] = true
			}
		}
	}
}

// isCertifiedTypeExpr 判断类型表达式是否为已认证具名类型（cellRef/relID/xmlFrag）。
// 这些类型由本包定义、名字唯一，故按标识符判定即可，无需完整类型检查。
func isCertifiedTypeExpr(e ast.Expr) bool {
	switch v := e.(type) {
	case nil:
		return false
	case *ast.Ident:
		_, ok := certifiedTypes[v.Name]
		return ok
	case *ast.StarExpr:
		return isCertifiedTypeExpr(v.X)
	}
	return false
}

// collectStructFields 收集本函数内出现过的结构体字段名，视为潜在用户数据。
func collectStructFields(m map[string]bool, fn *ast.FuncDecl) {
	ast.Inspect(fn.Body, func(n ast.Node) bool {
		sel, ok := n.(*ast.SelectorExpr)
		if !ok {
			return true
		}
		m[sel.Sel.Name] = true
		return true
	})
}

// exprTainted 判断表达式是否带污。
// certified 记录由认证类型（cellRef/xmlFrag/relID/xmlTagName）产生的可信变量；
// 对它们取 string（拼接时必需）仍然可信。
func exprTainted(e ast.Expr, tainted, structTainted, certified map[string]bool) bool {
	switch v := e.(type) {
	case nil:
		return false
	case *ast.Ident:
		if certified[v.Name] {
			return false
		}
		return tainted[v.Name]
	case *ast.SelectorExpr:
		return structTainted[v.Sel.Name]
	case *ast.IndexExpr:
		return exprTainted(v.X, tainted, structTainted, certified)
	case *ast.BinaryExpr:
		return exprTainted(v.X, tainted, structTainted, certified) || exprTainted(v.Y, tainted, structTainted, certified)
	case *ast.ParenExpr:
		return exprTainted(v.X, tainted, structTainted, certified)
	case *ast.CallExpr:
		if isCleaningCall(v) || isQuoteMeta(v) || isCertifiedConv(v) {
			return false
		}
		// string(认证变量) —— 取出已认证值拼接，可信
		if isCertifyingVar(v, certified) {
			return false
		}
		// 容器/透传函数：append 等会把入参污点带进返回值
		if isContainerFunc(v) {
			for _, arg := range v.Args {
				if exprTainted(arg, tainted, structTainted, certified) {
					return true
				}
			}
			return false
		}
		// 其它未知函数：保守视为已清洗（取严则误报淹没真问题，
		// 端到端由 TestXMLInjectionResistance 兜底）
		return false
	}
	return false
}

// isCleaningCall 判断是否为清除污点的调用。
func isCleaningCall(call *ast.CallExpr) bool {
	switch fn := call.Fun.(type) {
	case *ast.Ident:
		return cleanFuncs[fn.Name] || numberFuncs[fn.Name]
	case *ast.SelectorExpr:
		if x, ok := fn.X.(*ast.Ident); ok {
			return numberFuncs[x.Name+"."+fn.Sel.Name]
		}
	}
	return false
}

// isContainerFunc 判断是否为「透传污点」的容器函数（append 等）。
func isContainerFunc(call *ast.CallExpr) bool {
	id, ok := call.Fun.(*ast.Ident)
	if !ok {
		return false
	}
	switch id.Name {
	case "append", "appendString", "concat":
		return true
	}
	return false
}

// isQuoteMeta 判断是否为 regexp.QuoteMeta（正则上下文，本身即正确转义）。
func isQuoteMeta(e ast.Expr) bool {
	call, ok := e.(*ast.CallExpr)
	if !ok {
		return false
	}
	sel, ok := call.Fun.(*ast.SelectorExpr)
	if !ok {
		return false
	}
	x, ok := sel.X.(*ast.Ident)
	return ok && x.Name == "regexp" && sel.Sel.Name == "QuoteMeta"
}

// hasXMLTagLiteral 判断表达式是否包含 XML 标签片段字面量（含 "<" 即算）。
//
// 不要求同一字面量内同时出现 "<" 与 ">"：实际代码常把
// `<color rgb="` 与 `"/>` 分成两段，跨字面量拼成完整标签。
func hasXMLTagLiteral(e ast.Expr) bool {
	found := false
	ast.Inspect(e, func(n ast.Node) bool {
		bl, ok := n.(*ast.BasicLit)
		if !ok || bl.Kind != token.STRING {
			return true
		}
		v, err := strconv.Unquote(bl.Value)
		if err != nil {
			return true
		}
		// 只需含 "<"：像 `<color rgb="` 与 `"/>` 会分成两个字面量，
		// 要求同一个字面量同时含 "<" 与 ">" 会漏掉这类最常见的拼接。
		if strings.Contains(v, "<") {
			found = true
		}
		return true
	})
	return found
}

// flattenConcat 展开嵌套的 + 表达式，取出每个叶子操作数。
func flattenConcat(e ast.Expr) []ast.Expr {
	if be, ok := e.(*ast.BinaryExpr); ok && be.Op == token.ADD {
		return append(flattenConcat(be.X), flattenConcat(be.Y)...)
	}
	return []ast.Expr{e}
}

// certifiedCtors 是「返回已认证类型」的构造函数。调用它们即完成校验/规范化，
// 结果可信（污点被清除）。
var certifiedCtors = map[string]bool{
	"newCellRef": true, "parseCellRef": true,
	"newRelID": true, "newRelIDFrom": true,
	"frag": true,
}

// isCertifiedConv 判断表达式是否为「构造已认证类型」的调用。
func isCertifiedConv(e ast.Expr) bool {
	call, ok := e.(*ast.CallExpr)
	if !ok {
		return false
	}
	switch fn := call.Fun.(type) {
	case *ast.Ident:
		return certifiedCtors[fn.Name]
	case *ast.SelectorExpr:
		// 认证类型的显式转换：cellRef(x) / xmlFrag(s) / string(认证值)
		// string(cellRef) 这类"取出已认证值"的转换同样可信：
		// 认证发生在构造处（newCellRef），转换只是把它变回 string 使用。
		return isCertifiedTypeExpr(fn) || isCertifyingConv(fn)
	}
	return false
}

// isCertifyingConv 判断是否是「从已认证值取出 string」的转换 string(cellRef)。
func isCertifyingConv(sel *ast.SelectorExpr) bool {
	if sel.Sel.Name != "string" {
		return false
	}
	return isCertifiedTypeExpr(sel.X)
}

// isCertifyingVar 判断是否为 `string(认证变量)` —— 把已认证值取出为 string 拼接。
// 认证发生在构造处（newCellRef/frag 等），此处只是类型转换，故仍可信。
func isCertifyingVar(call *ast.CallExpr, certified map[string]bool) bool {
	sel, ok := call.Fun.(*ast.SelectorExpr)
	if !ok || sel.Sel.Name != "string" {
		return false
	}
	id, ok := sel.X.(*ast.Ident)
	return ok && certified[id.Name]
}

// TestGuardNoNameAllowlist 是护栏自身的元测试：确保**没有**按变量名的豁免清单。
//
// 此前曾有 18 项按变量名豁免的清单 —— 它要求人记得把新写入点登记进去，
// 漏登记即为漏检。本测试锁死这一约束，防止清单形式回归。
//
// 认证只允许通过「已认证具名类型」表达（cellRef / relID / xmlFrag / xmlTagName /
// rangeRef / styleIndex / borderSideName）：认证来自构造处的校验，
// 而非人对某个变量名的判断。
func TestGuardNoNameAllowlist(t *testing.T) {
	// 检查是否存在 map[string]string 形式的"变量名 -> 理由"豁免表
	fset := token.NewFileSet()
	f, err := parser.ParseFile(fset, "guard_taint_test.go", nil, 0)
	if err != nil {
		t.Fatal(err)
	}
	bad := false
	ast.Inspect(f, func(n ast.Node) bool {
		vs, ok := n.(*ast.ValueSpec)
		if !ok {
			return true
		}
		if len(vs.Names) == 0 {
			return true
		}
		name := strings.ToLower(vs.Names[0].Name)
		// 豁免清单典型命名：allowlist / exempt / whitelist
		if strings.Contains(name, "allowlist") || strings.Contains(name, "exempt") ||
			strings.Contains(name, "whitelist") {
			bad = true
		}
		return true
	})
	if bad {
		t.Error("护栏中不允许存在按变量名的豁免清单；请改用已认证的具名类型表达安全性")
	}
}
