# xml-xlsx-lite

A lightweight Excel XLSX file generator with complete Excel functionality, including pivot tables, charts, and advanced styling. Supports Traditional Chinese and Taiwan region usage.

## ✨ Features

- **Complete Excel Support**: Full XLSX file generation with all Excel features
- **String Writing Support**: Full support for strings, including Chinese characters and emojis
- **Dynamic Pivot Tables**: Insert native Excel pivot tables into existing workbooks
- **Pivot Tables**: Create and export pivot table results with data aggregation
- **Charts**: Basic chart support (preservation mode)
- **Performance Optimization**: Smart sharedStrings switching and streaming processing
- **Advanced Styling**: Fonts, colors, borders, alignment, and number formatting
- **Protection**: Worksheet and workbook protection with password control
- **Error Handling**: Comprehensive error handling system
- **Reading Support**: Read and parse existing Excel files
- **TypeScript**: Full TypeScript support with type definitions
- **Lightweight**: Minimal dependencies, inspired by exceljs API

## 🚀 Quick Start

### Installation

```bash
npm install xml-xlsx-lite
```

### Basic Usage

```javascript
const { Workbook } = require('xml-xlsx-lite');
const fs = require('fs');

async function main() {
  // Create workbook
  const wb = new Workbook();
  
  // Create worksheet
  const ws = wb.getWorksheet('Data');
  
  // Add data
  ws.setCell('A1', 'Name', { font: { bold: true } });
  ws.setCell('B1', 'Value', { font: { bold: true } });
  
  ws.setCell('A2', 'Item 1');
  ws.setCell('B2', 100, { numFmt: '#,##0' });
  
  ws.setCell('A3', 'Item 2');
  ws.setCell('B3', 200, { numFmt: '#,##0' });
  
  // Set column width
  ws.setColumnWidth('A', 15);
  ws.setColumnWidth('B', 12);
  
  // Save file
  const buffer = await wb.writeBuffer();
  fs.writeFileSync('output.xlsx', new Uint8Array(buffer));
  console.log('Excel file generated successfully!');
}

main();
```

### TypeScript Usage

```typescript
import { Workbook } from 'xml-xlsx-lite';
import * as fs from 'fs';

interface DataItem {
  name: string;
  value: number;
}

async function main(): Promise<void> {
  const wb = new Workbook();
  const ws = wb.getWorksheet('Data');
  
  const data: DataItem[] = [
    { name: 'Item 1', value: 100 },
    { name: 'Item 2', value: 200 },
    { name: 'Item 3', value: 300 }
  ];
  
  // Add headers
  ws.setCell('A1', 'Name', { font: { bold: true } });
  ws.setCell('B1', 'Value', { font: { bold: true } });
  
  // Add data
  data.forEach((item, index) => {
    const row = index + 2;
    ws.setCell(`A${row}`, item.name);
    ws.setCell(`B${row}`, item.value, { numFmt: '#,##0' });
  });
  
  // Save file
  const buffer = await wb.writeBuffer();
  fs.writeFileSync('output.xlsx', new Uint8Array(buffer));
}

main();
```

## 📊 Dynamic Pivot Table Example

```javascript
const { addPivotToWorkbookBuffer } = require('xml-xlsx-lite');
const fs = require('fs');

async function createDynamicPivot() {
  // Read existing workbook
  const baseBuffer = fs.readFileSync('base-workbook.xlsx');
  
  // Configure pivot table
  const pivotOptions = {
    sourceSheet: "Data",
    sourceRange: "A1:D100",
    targetSheet: "Pivot",
    anchorCell: "A3",
    layout: {
      rows: [{ name: "Department" }],
      cols: [{ name: "Month" }],
      values: [{ 
        name: "Sales", 
        agg: "sum", 
        displayName: "Total Sales" 
      }]
    },
    refreshOnLoad: true,
    styleName: "PivotStyleMedium9"
  };
  
  // Insert dynamic pivot table
  const enhancedBuffer = await addPivotToWorkbookBuffer(baseBuffer, pivotOptions);
  
  // Save result
  fs.writeFileSync('pivot-workbook.xlsx', enhancedBuffer);
  console.log('Dynamic pivot table inserted!');
}

createDynamicPivot();
```

## 📊 Manual Pivot Table Example

```javascript
const { Workbook } = require('xml-xlsx-lite');
const fs = require('fs');

async function createPivotTable() {
  const wb = new Workbook();
  
  // Create data worksheet
  const dataSheet = wb.getWorksheet('Data');
  
  // Add sample data
  const data = [
    ['Month', 'Department', 'Sales'],
    ['January', 'A', 1000],
    ['January', 'B', 2000],
    ['February', 'A', 1500],
    ['February', 'B', 2500]
  ];
  
  data.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      const address = String.fromCharCode(65 + colIndex) + (rowIndex + 1);
      dataSheet.setCell(address, cell);
    });
  });
  
  // Create pivot table result worksheet (manual approach)
  const pivotSheet = wb.getWorksheet('Pivot Table');
  
  // Add headers
  pivotSheet.setCell('A1', 'Sales Summary', {
    font: { bold: true, size: 16 },
    alignment: { horizontal: 'center' }
  });
  
  pivotSheet.setCell('A3', 'Month', { font: { bold: true } });
  pivotSheet.setCell('B3', 'Department A', { font: { bold: true } });
  pivotSheet.setCell('C3', 'Department B', { font: { bold: true } });
  pivotSheet.setCell('D3', 'Total', { font: { bold: true } });
  
  // Add calculated results
  const pivotData = [
    ['January', 1000, 2000, 3000],
    ['February', 1500, 2500, 4000]
  ];
  
  pivotData.forEach((row, index) => {
    const rowNum = index + 4;
    pivotSheet.setCell(`A${rowNum}`, row[0]);
    pivotSheet.setCell(`B${rowNum}`, row[1], { numFmt: '#,##0' });
    pivotSheet.setCell(`C${rowNum}`, row[2], { numFmt: '#,##0' });
    pivotSheet.setCell(`C${rowNum}`, row[2], { numFmt: '#,##0' });
    pivotSheet.setCell(`D${rowNum}`, row[3], { 
      numFmt: '#,##0',
      font: { bold: true }
    });
  });
  
  // Save file
  const buffer = await wb.writeBuffer();
  fs.writeFileSync('pivot-example.xlsx', new Uint8Array(buffer));
}

createPivotTable();
```

## 🎨 Styling Options

### Cell Styling

```javascript
// Font styling
ws.setCell('A1', 'Bold Text', {
  font: { 
    bold: true, 
    size: 16, 
    color: 'FF0000' 
  }
});

// Alignment
ws.setCell('B1', 'Centered', {
  alignment: { 
    horizontal: 'center', 
    vertical: 'middle' 
  }
});

// Background color
ws.setCell('C1', 'Background', {
  fill: { 
    type: 'pattern', 
    color: 'E0E0E0' 
  }
});

// Borders
ws.setCell('D1', 'Bordered', {
  border: {
    top: { style: 'thick', color: '000000' },
    bottom: { style: 'thick', color: '000000' }
  }
});

// Number format
ws.setCell('E1', 1234.56, {
  numFmt: '#,##0.00'
});
```

### Column and Row Settings

```javascript
// Set column width
ws.setColumnWidth('A', 15);
ws.setColumnWidth('B', 20);

// Set row height
ws.setRowHeight(1, 30);
ws.setRowHeight(2, 25);
```

## 🔒 Worksheet Protection

```javascript
// Protect worksheet
ws.protect({
  password: 'password123',
  selectLockedCells: false,
  selectUnlockedCells: true,
  formatCells: false,
  insertRows: false,
  deleteRows: false,
  sort: false,
  autoFilter: false
});
```

## 📋 API Reference

### Workbook

- `new Workbook()` - Create new workbook
- `getWorksheet(name)` - Get or create worksheet
- `writeBuffer()` - Generate Excel file as buffer
- `writeFile(filePath)` - Write Excel file directly to disk
- `writeFileWithPivotTables(filePath, options)` - Write file with pivot table
- `writeFileWithMultiplePivots(filePath, optionsArray)` - Write file with multiple pivot tables
- `createManualPivotTable(data, options)` - Create programmatic pivot table
- `getWorksheets()` - Get all worksheets

### Worksheet

- `setCell(address, value, options)` - Set cell value and styling
- `getCell(address)` - Get cell value
- `setColumnWidth(column, width)` - Set column width
- `setRowHeight(row, height)` - Set row height
- `protect(options)` - Protect worksheet

### Cell Options

```javascript
{
  font: {
    bold: boolean,
    italic: boolean,
    size: number,
    color: string,
    name: string
  },
  alignment: {
    horizontal: 'left' | 'center' | 'right',
    vertical: 'top' | 'middle' | 'bottom',
    wrapText: boolean
  },
  fill: {
    type: 'pattern' | 'gradient',
    color: string
  },
  border: {
    style: 'thin' | 'medium' | 'thick',
    color: string
  },
  numFmt: string
}
```

## 🚨 Important Notes

### File Saving

**Now you can use `writeFile()` method!** - We've added a thin wrapper:

```javascript
// ✅ New - use writeFile directly
await workbook.writeFile('output.xlsx');

// ✅ Still works - use writeBuffer
const buffer = await workbook.writeBuffer();
fs.writeFileSync('output.xlsx', new Uint8Array(buffer));
```

### New API Methods

**Added for better user experience:**

```javascript
// Write file with pivot table
await workbook.writeFileWithPivotTables('output.xlsx', pivotOptions);

// Write file with multiple pivot tables
await workbook.writeFileWithMultiplePivots('output.xlsx', [pivot1, pivot2]);

// Create manual pivot table (programmatic aggregation)
const result = workbook.createManualPivotTable(data, {
  rowField: 'Department',
  columnField: 'Month', 
  valueField: 'Sales',
  aggregation: 'sum'
});
```

### Pivot Tables

**Avoid automatic pivot table creation** - use manual approach:

```javascript
// ❌ Don't use this (has issues)
const pivotTable = workbook.createPivotTable(config);

// ✅ Use manual creation
const pivotSheet = workbook.getWorksheet('Pivot Table');
// Manually add data and calculations...
```

## 🔧 Troubleshooting

### Common Issues

1. **TypeScript errors**: Ensure proper import paths
2. **File operations**: Use `writeFile()` for direct file writing
3. **Pivot table issues**: Use `createManualPivotTable()` for programmatic aggregation
4. **Build warnings**: Check package.json exports configuration

### Error Solutions

```javascript
// Error: Property 'setCell' does not exist
// Solution: Check import statement
import { Workbook } from 'xml-xlsx-lite';

// Error: File system operations not available
// Solution: Ensure you're running in Node.js environment
// Browser environments don't support file operations
```

## 📚 Documentation

- **API Reference**: [README-API.md](./README-API.md) | [English Version](./README-API-EN.md)
- **Dynamic Pivot Tables**: [DYNAMIC_PIVOT_USAGE.md](./DYNAMIC_PIVOT_USAGE.md)
- **Usage Guide**: [USAGE_GUIDE_FIXED.md](./USAGE_GUIDE_FIXED.md)
- **Pivot Table Fix**: [PIVOT_TABLE_FIX_REPORT.md](./PIVOT_TABLE_FIX_REPORT.md)

## 🌟 Why Choose xml-xlsx-lite?

- **Lightweight**: Minimal dependencies, fast performance
- **Complete**: Full Excel functionality support
- **TypeScript**: Excellent TypeScript support
- **Flexible**: Easy to use API with powerful styling options
- **Reliable**: Stable and well-tested
- **Chinese Support**: Built with Traditional Chinese users in mind

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guide](./CONTRIBUTING.md) for details.

## 📄 License

MIT License - see [LICENSE](./LICENSE) file for details.

## 🔗 Links

- **NPM Package**: https://www.npmjs.com/package/xml-xlsx-lite
- **GitHub Repository**: https://github.com/mikemikex1/xml-xlsx-lite
- **Issues**: https://github.com/mikemikex1/xml-xlsx-lite/issues

---

**xml-xlsx-lite** - Your lightweight Excel solution! 🚀
