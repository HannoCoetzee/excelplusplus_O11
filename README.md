# ExcelPlusPlus for OutSystems 11

[![License: BSD-3-Clause](https://img.shields.io/badge/License-BSD--3--Clause-blue.svg)](LICENSE)
[![OutSystems 11](https://img.shields.io/badge/OutSystems-11-blue)](https://www.outsystems.com)
[![EPPlus](https://img.shields.io/badge/EPPlus-8.5.4-green)](https://epplussoftware.com/)

Professional Excel manipulation for OutSystems 11 applications. Create, read, and modify Excel files with full formatting support, powered by EPPlus.

---

## 🚀 Quick Start

### Installation

1. Download from [OutSystems Forge](https://www.outsystems.com/forge/component-overview/excelplusplus)
2. Import the `.xif` extension into your environment
3. Reference `OutSystems.NssExcelPlusPlus` in your modules
4. Start creating Excel files!

### Basic Example

```outsystems
// Create a new workbook
WorkbookCreate("Sales Report", wbHandle)

// Add a worksheet
WorksheetAdd(wbHandle, "Q1 Sales", wsHandle)

// Write header row
CellWrite(wsHandle, 1, 1, "Product")
CellWrite(wsHandle, 1, 2, "Revenue")
CellWrite(wsHandle, 1, 3, "Units")

// Write data
CellWrite(wsHandle, 2, 1, "Widget A")
CellWrite(wsHandle, 2, 2, 15000)
CellWrite(wsHandle, 2, 3, 150)

// Apply formatting
CellSetBold(wsHandle, 1, 1, True)
CellSetBold(wsHandle, 1, 2, True)
CellSetBold(wsHandle, 1, 3, True)
CellSetNumberFormat(wsHandle, 2, 2, "$#,##0")

// Save to stream
WorkbookSaveToStream(wbHandle, outputStream)

// Download or attach
DownloadFile("Sales_Report.xlsx", outputStream)
```

---

## 📚 Documentation

### Core Actions

**Workbook Operations:**
- `WorkbookCreate(Name, Handle)` - Create new workbook
- `WorkbookOpen(Stream, Handle)` - Open existing workbook
- `WorkbookSaveToStream(Handle, Stream)` - Save to stream
- `WorkbookClose(Handle)` - Close and cleanup

**Worksheet Operations:**
- `WorksheetAdd(Workbook, Name, Handle)` - Add worksheet
- `WorksheetGet(Workbook, Name, Handle)` - Get worksheet by name
- `WorksheetDelete(Workbook, Name)` - Delete worksheet

**Cell Operations:**
- `CellWrite(Worksheet, Row, Col, Value)` - Write cell value
- `CellRead(Worksheet, Row, Col)` - Read cell value
- `CellSetBold(Worksheet, Row, Col, Bold)` - Set bold
- `CellSetItalic(Worksheet, Row, Col, Italic)` - Set italic
- `CellSetNumberFormat(Worksheet, Row, Col, Format)` - Set number format
- `CellSetBackgroundColor(Worksheet, Row, Col, Color)` - Set background
- `CellMerge(Worksheet, StartRow, StartCol, EndRow, EndCol)` - Merge cells

**Advanced Operations:**
- `RangeWrite(Worksheet, StartRow, StartCol, Data)` - Write range
- `TableCreate(Worksheet, Range, Name)` - Create Excel table
- `ChartAdd(Worksheet, Type, Range)` - Add chart (Pro)
- `ImageAdd(Worksheet, Path, Row, Col)` - Add image (Pro)

### Data Types Supported

- Text (strings)
- Numbers (integers, decimals)
- Dates and times
- Booleans
- Formulas (as strings starting with `=`)
- Null/empty

---

## 🎯 Common Use Cases

### Export Data to Excel

```outsystems
// Query your data
Users = UserRecordList()

// Create workbook
WorkbookCreate("User Export", wb)
WorksheetAdd(wb, "Users", ws)

// Write headers
CellWrite(ws, 1, 1, "Name")
CellWrite(ws, 1, 2, "Email")
CellWrite(ws, 1, 3, "Role")

// Write data
For i = 1 to Length(Users)
    CellWrite(ws, i+1, 1, Users[i].Name)
    CellWrite(ws, i+1, 2, Users[i].Email)
    CellWrite(ws, i+1, 3, Users[i].Role)
End For

// Save and download
WorkbookSaveToStream(wb, stream)
DownloadFile("Users.xlsx", stream)
```

### Read Excel Template

```outsystems
// Upload template
UploadFile("Template.xlsx", templateStream)

// Open workbook
WorkbookOpen(templateStream, wb)
WorksheetGet(wb, "Data", ws)

// Read data
For row = 2 to 100
    value = CellRead(ws, row, 1)
    If value Is Null Then Exit For
    
    // Process data
    ProcessRow(value)
End For

WorkbookClose(wb)
```

### Generate Invoice

```outsystems
WorkbookCreate("Invoice", wb)
WorksheetAdd(wb, "Invoice", ws)

// Company header
CellWrite(ws, 1, 1, "Your Company Name")
CellSetBold(ws, 1, 1, True)
CellMerge(ws, 1, 1, 1, 3)

// Invoice details
CellWrite(ws, 3, 1, "Invoice #: " + invoiceNumber)
CellWrite(ws, 3, 3, "Date: " + Today())

// Line items
CellWrite(ws, 5, 1, "Item")
CellWrite(ws, 5, 2, "Qty")
CellWrite(ws, 5, 3, "Price")

For i = 1 to Length(items)
    CellWrite(ws, 5+i, 1, items[i].Name)
    CellWrite(ws, 5+i, 2, items[i].Qty)
    CellWrite(ws, 5+i, 3, items[i].Price)
End For

// Total
CellWrite(ws, 5+i+1, 2, "Total:")
CellSetBold(ws, 5+i+1, 2, True)
CellWrite(ws, 5+i+1, 3, totalAmount)
CellSetNumberFormat(ws, 5+i+1, 3, "$#,##0.00")

WorkbookSaveToStream(wb, stream)
```

---

## ⚡ Performance Tips

1. **Use bulk operations** when possible - `RangeWrite` is faster than individual `CellWrite` calls
2. **Close workbooks** promptly - Always call `WorkbookClose` to free resources
3. **Stream large files** - Use streaming for files >10MB
4. **Avoid excessive formatting** - Apply formatting to ranges, not individual cells
5. **Cache templates** - Load templates once, reuse across requests

---

## 🔧 Troubleshooting

**Issue: "Out of memory" error**
- Solution: Use streaming mode for large files
- Close workbooks promptly after use

**Issue: Formatting not applied**
- Solution: Ensure formatting actions happen before `WorkbookSaveToStream`

**Issue: Special characters corrupted**
- Solution: Ensure UTF-8 encoding in your application

**Issue: Date format incorrect**
- Solution: Use `CellSetNumberFormat` with appropriate date format (e.g., "dd/mm/yyyy")

---

## 📄 License

BSD-3-Clause License - Free for commercial and personal use.

See [LICENSE](LICENSE) for details.

---

## 🆘 Support

**Community Edition:**
- GitHub Issues: https://github.com/HannoCoetzee/excelplusplus_O11/issues
- OutSystems Forums: https://www.outsystems.com/forums/

**Professional Edition:**
- Priority email support
- 24-hour response SLA
- Direct maintainer access

---

## 📊 Version

**Current:** 1.0.0 (Community)
**Release Date:** May 2026

---

## 🙏 Credits

Built on [EPPlus](https://epplussoftware.com/) by EPPlus Software AB.

---

*Made with ❤️ by Hanno Coetzee*
