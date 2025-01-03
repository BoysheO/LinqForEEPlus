[中文](README.md) | [English](README_EN.md)

If you find this library useful, please give it a star ⭐ to show your support! 🙇‍

## Design Goals
This library aims to rewrap the EPPlus 4.5 API to make it more convenient to use.

---

## User Guide

### Example: Reading Excel

```csharp
// First, load the file into memory to avoid EPPlus locking the Excel file.
using var ms = new MemoryStream();
using (var fs = new FileStream(file.FullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
{
    await fs.CopyToAsync(ms);
}

ms.Seek(0, SeekOrigin.Begin);
using ExcelPackage excel = new ExcelPackage(ms);
var ebook = new LinqForEEPlus.Book(excel.Workbook, file.FullName);

// Access the first sheet, supports IEnumerable<Sheet> interface for LINQ operations
var esheet = ebook[0];
var row = esheet[0]; // Supports IEnumerable<Row> interface for LINQ operations
var cell = esheet[rowNum, colNum]; // Note: In EPPlus, colNum starts from 1 (not 0), valid range is [1,+∞)

// Get cell values
var cellStr = cell.StrValue; // Returns a trimmed value and replaces null with an empty string ""
var cellFlo = cell.Flo; // Use IsFlo to check or TryGetFlo method, otherwise, non-float values will throw exceptions

// Get Excel column letters from column number
var colLet = Sheet.GetColumnLetter(colNum); // Converts column number (e.g., 123) to Excel column letters (e.g., "ABC")
```

---

## FAQ

**Q: Why use EPPlus 4.5 instead of the latest version?**  
**A:**  
1. Later versions of EPPlus changed their open-source license, making them less convenient to use.  
2. Version 4.5 provides sufficient functionality for our needs.

**Q: Will this library still be maintained?**  
**A:**  
Yes, but updates will not be frequent as no bugs have been discovered for a long time 😀.

**Q: How is the performance, compatibility, and functionality?**  
**A:**  
- **Performance**: Average, not optimized for high performance.  
- **Compatibility**: Supports Unity.  
- **Functionality**: Only supports reading operations.

---

If you need more features or have suggestions, feel free to submit an Issue or PR!

