# 中文 | English

如果觉得这个库好用，请点个星星 ⭐ 支持一下！🙇‍

## 设计目标
本库旨在重新封装 EPPlus 4.5 的 API，使其更加易用。

---

## 使用指南

### 读取 Excel 示例

```csharp
// 先将文件读取到内存中，避免 EPPlus 锁定 Excel 文件。
using var ms = new MemoryStream();
using (var fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    await fs.CopyToAsync(ms);
}

ms.Seek(0, SeekOrigin.Begin);
using ExcelPackage excel = new ExcelPackage(ms);
var ebook = new LinqForEEPlus.Book(excel.Workbook, file.FullName);

// 访问第一个 Sheet，支持 IEnumerable<Sheet> 接口，可使用 LINQ 运算符
var esheet = ebook[0];
var row = esheet[0]; // 支持 IEnumerable<Row> 接口，可使用 LINQ 运算符
var cell = esheet[rowNum, colNum]; // 注意：在 EPPlus 中，colNum 范围为 [1,+∞)，第一列的 colNum 是 1 而不是 0

// 获取单元格值
var cellStr = cell.StrValue; // 返回值已进行 Trim 操作，并将 null 替换为 ""
var cellFlo = cell.Flo; // 请使用 IsFlo 检查或 TryGetFlo 方法，否则非 float 值会抛出异常

// 获取列号对应的 Excel 字母
var colLet = Sheet.GetColumnLetter(colNum); // 将列号 123 转换为 Excel 中的列号 "ABC"
```

---

## 常见问题

**Q：为什么使用 EPPlus 4.5，而不是最新版本？**  
**A：** 1. 新版本的 EPPlus 修改了开源协议，使用不便。  
2. 4.5 版本功能已足够满足需求。

**Q：还会维护这个库吗？**  
**A：** 会，但更新频率不高，因为已经很久没有发现 bug 了 😀。

**Q：性能如何？兼容性如何？功能如何？**  
*A：* **性能**：破烂  **兼容性**：支持 Unity。   **功能**：仅支持读取操作。

---

如果需要更多功能或有任何建议，欢迎提交 Issue 或 PR！
