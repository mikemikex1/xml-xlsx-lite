# 🚀 xml-xlsx-lite

> **Lightweight Excel XLSX generator with full Excel features: dynamic pivot tables, charts, styles, and Chinese support. Fast, TypeScript-friendly Excel file creation library.**

[![npm version](https://img.shields.io/npm/v/xml-xlsx-lite.svg)](https://www.npmjs.com/package/xml-xlsx-lite)
[![npm downloads](https://img.shields.io/npm/dm/xml-xlsx-lite.svg)](https://www.npmjs.com/package/xml-xlsx-lite)
[![License](https://img.shields.io/npm/l/xml-xlsx-lite.svg)](https://github.com/mikemikex1/xml-xlsx-lite/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.0+-blue.svg)](https://www.typescriptlang.org/)

## 📋 Table of Contents

- [🚀 Features](#-features)
- [📦 Installation](#-installation)
- [🎯 Quick Start](#-quick-start)
- [📚 Complete Guide](#-complete-guide)
  - [1. Create Excel Files](#1-create-excel-files)
  - [2. Basic Cell Operations](#2-basic-cell-operations)
  - [3. Styling and Formatting](#3-styling-and-formatting)
  - [4. Worksheet Management](#4-worksheet-management)
  - [5. Read Excel Files](#5-read-excel-files)
  - [6. Pivot Tables](#6-pivot-tables)
  - [7. Chart Support](#7-chart-support)
  - [8. Advanced Features](#8-advanced-features)
- [🔧 API Reference](#-api-reference)
- [📖 Examples](#-examples)
- [🚨 Important Notes](#-important-notes)
- [📊 Feature Matrix](#-feature-matrix)
- [🌐 Browser Support](#-browser-support)
- [🤝 Contributing](#-contributing)
- [📄 License](#-license)

## 🚀 Features

- ✨ **Complete Excel Support**: Full XLSX format support with all Excel features
- 🔄 **Dynamic Pivot Tables**: Insert refreshable pivot tables into existing workbooks
- 📊 **Chart Support**: Create and preserve Excel charts
- 🎨 **Rich Styling**: Comprehensive cell formatting, borders, colors, and fonts
- 🌏 **Chinese Support**: Full Traditional Chinese character support
- ⚡ **High Performance**: Optimized for large files with streaming support
- 🔧 **TypeScript Ready**: Built with TypeScript, includes type definitions
- 📱 **Cross Platform**: Works in Node.js and modern browsers
- 🚀 **Lightweight**: Minimal dependencies, fast build times

## 📦 Installation

```bash
npm install xml-xlsx-lite
```

## 🎯 Quick Start

```typescript
import { Workbook } from 'xml-xlsx-lite';

// Create workbook
const workbook = new Workbook();

// Get worksheet
const worksheet = workbook.getWorksheet('Sheet1');

// Set cells
worksheet.setCell('A1', 'Hello World');
worksheet.setCell('B1', 42);

// Save file
const buffer = await workbook.writeBuffer();
```

## 📚 Complete Guide

### 1. Create Excel Files

#### 1.1 Basic Workbook Creation

```typescript
import { Workbook } from 'xml-xlsx-lite';

// Create new workbook
const workbook = new Workbook();

// Get default worksheet
const worksheet = workbook.getWorksheet('Sheet1');

// Set title
worksheet.setCell('A1', 'Product Sales Report');
worksheet.setCell('A2', '2024 Annual');

// Set column headers
worksheet.setCell('A3', 'Product Name');
worksheet.setCell('B3', 'Sales Quantity');
worksheet.setCell('C3', 'Unit Price');
worksheet.setCell('D3', 'Total Amount');

// Set data
worksheet.setCell('A4', 'Laptop');
worksheet.setCell('B4', 10);
worksheet.setCell('C4', 35000);
worksheet.setCell('D4', 350000);

worksheet.setCell('A5', 'Mouse');
worksheet.setCell('B5', 50);
worksheet.setCell('C5', 500);
worksheet.setCell('D5', 25000);

// Save file
const buffer = await workbook.writeBuffer();
```

#### 1.2 Multi-Worksheet Workbook

```typescript
const workbook = new Workbook();

// Create multiple worksheets
const summarySheet = workbook.getWorksheet('Summary');
const detailSheet = workbook.getWorksheet('Detailed Data');
const chartSheet = workbook.getWorksheet('Charts');

// Set data in different worksheets
summarySheet.setCell('A1', 'Sales Summary');
detailSheet.setCell('A1', 'Detailed Sales Data');
chartSheet.setCell('A1', 'Sales Charts');
```

### 2. Basic Cell Operations

#### 2.1 Cell Value Setting

```typescript
const worksheet = workbook.getWorksheet('Sheet1');

// Different types of data
worksheet.setCell('A1', 'Text');                    // String
worksheet.setCell('B1', 123);                       // Number
worksheet.setCell('C1', true);                      // Boolean
worksheet.setCell('D1', new Date());                // Date
worksheet.setCell('E1', null);                      // Null
worksheet.setCell('F1', '');                        // Empty string

// Using coordinates
worksheet.setCell('G1', 'Using A1 coordinates');
worksheet.setCell(1, 8, 'Using row-column coordinates'); // Row 1, Column 8
```

#### 2.2 Cell Range Operations

```typescript
// Set cells in a range
for (let row = 1; row <= 10; row++) {
    for (let col = 1; col <= 5; col++) {
        const value = `R${row}C${col}`;
        worksheet.setCell(row, col, value);
    }
}

// Set entire row
for (let col = 1; col <= 5; col++) {
    worksheet.setCell(1, col, `Title${col}`);
}

// Set entire column
for (let row = 1; row <= 10; row++) {
    worksheet.setCell(row, 1, `Item${row}`);
}
```

### 3. Styling and Formatting

#### 3.1 Basic Styling

```typescript
// Font styling
worksheet.setCell('A1', 'Bold Title', {
    font: {
        bold: true,
        size: 16,
        color: 'FF0000'  // Red
    }
});

// Alignment styling
worksheet.setCell('B1', 'Center Aligned', {
    alignment: {
        horizontal: 'center',
        vertical: 'middle'
    }
});

// Border styling
worksheet.setCell('C1', 'With Borders', {
    border: {
        top: { style: 'thin', color: '000000' },
        bottom: { style: 'double', color: '000000' },
        left: { style: 'thin', color: '000000' },
        right: { style: 'thin', color: '000000' }
    }
});

// Fill styling
worksheet.setCell('D1', 'With Background', {
    fill: {
        type: 'solid',
        color: 'FFFF00'  // Yellow
    }
});
```

#### 3.2 Number Formatting

```typescript
// Currency format
worksheet.setCell('A1', 1234.56, {
    numFmt: '¥#,##0.00'
});

// Percentage format
worksheet.setCell('B1', 0.1234, {
    numFmt: '0.00%'
});

// Date format
worksheet.setCell('C1', new Date(), {
    numFmt: 'yyyy-mm-dd'
});

// Custom format
worksheet.setCell('D1', 42, {
    numFmt: '0 "items"'
});
```

#### 3.3 Merge Cells

```typescript
// Merge cells
worksheet.mergeCells('A1:D1');
worksheet.setCell('A1', 'Merged Title');

// Merge multiple rows
worksheet.mergeCells('A2:A5');
worksheet.setCell('A2', 'Vertical Merge');
```

### 4. Worksheet Management

#### 4.1 Column Width and Row Height

```typescript
// Set column width
worksheet.setColumnWidth('A', 20);      // Column A width 20
worksheet.setColumnWidth(2, 15);        // Column B width 15

// Set row height
worksheet.setRowHeight(1, 30);          // Row 1 height 30
worksheet.setRowHeight(2, 25);          // Row 2 height 25
```

#### 4.2 Freeze Panes

```typescript
// Freeze first row and first column
worksheet.freezePanes(2, 2);

// Freeze only first row
worksheet.freezePanes(2);

// Freeze only first column
worksheet.freezePanes(undefined, 2);

// Unfreeze panes
worksheet.unfreezePanes();
```

#### 4.3 Worksheet Protection

```typescript
// Protect worksheet
worksheet.protect('password123', {
    selectLockedCells: false,
    selectUnlockedCells: true,
    formatCells: false,
    formatColumns: false,
    formatRows: false
});

// Check protection status
const isProtected = worksheet.isProtected();
```

### 5. Read Excel Files

#### 5.1 Basic Reading

```typescript
import { Workbook } from 'xml-xlsx-lite';

// Read from file
const workbook = await Workbook.readFile('existing-file.xlsx');

// Read from Buffer
const fs = require('fs');
const buffer = fs.readFileSync('existing-file.xlsx');
const workbook = await Workbook.readBuffer(buffer);
```

#### 5.2 Read Worksheet Data

```typescript
// Get worksheet
const worksheet = workbook.getWorksheet('Sheet1');

// Convert to 2D array
const arrayData = worksheet.toArray();
console.log('Array data:', arrayData);

// Convert to JSON object array
const jsonData = worksheet.toJSON({ headerRow: 1 });
console.log('JSON data:', jsonData);

// Get specific range
const rangeData = worksheet.getRange('A1:D10');
console.log('Range data:', rangeData);
```

#### 5.3 Reading Options

```typescript
const workbook = await Workbook.readFile('file.xlsx', {
    enableSharedStrings: true,      // Enable shared strings optimization
    preserveStyles: true,           // Preserve style information
    preserveFormulas: true,         // Preserve formulas
    preservePivotTables: true,      // Preserve pivot tables
    preserveCharts: true            // Preserve charts
});
```

### 6. Pivot Tables

#### 6.1 Manual Pivot Table Creation

```typescript
// Create manual pivot table
const pivotData = [
    { department: 'IT', month: 'Jan', sales: 1000 },
    { department: 'IT', month: 'Feb', sales: 1200 },
    { department: 'HR', month: 'Jan', sales: 800 },
    { department: 'HR', month: 'Feb', sales: 900 }
];

const pivotSheet = workbook.getWorksheet('Pivot');
workbook.createManualPivotTable(pivotData, {
    rowField: 'department',
    columnField: 'month',
    valueField: 'sales',
    aggregation: 'sum',
    numberFormat: '#,##0'
});
```

#### 6.2 Dynamic Pivot Tables

```typescript
// Create base workbook
const workbook = new Workbook();
const dataSheet = workbook.getWorksheet('Data');

// Fill in data
const data = [
    ['Department', 'Month', 'Sales'],
    ['IT', 'Jan', 1000],
    ['IT', 'Feb', 1200],
    ['HR', 'Jan', 800],
    ['HR', 'Feb', 900]
];

data.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
        const address = String.fromCharCode(65 + colIndex) + (rowIndex + 1);
        dataSheet.setCell(address, value);
    });
});

// Save base file
const baseBuffer = await workbook.writeBuffer();

// Dynamically insert pivot table
import { addPivotToWorkbookBuffer } from 'xml-xlsx-lite';

const enhancedBuffer = await addPivotToWorkbookBuffer(baseBuffer, {
    sourceSheet: 'Data',
    sourceRange: 'A1:C100',
    targetSheet: 'Pivot',
    anchorCell: 'A3',
    layout: {
        rows: [{ name: 'Department' }],
        cols: [{ name: 'Month' }],
        values: [{ 
            name: 'Sales', 
            agg: 'sum', 
            displayName: 'Total Sales' 
        }]
    },
    refreshOnLoad: true,
    styleName: 'PivotStyleMedium9'
});
```

#### 6.3 Pivot Table Configuration Options

```typescript
const pivotOptions = {
    sourceSheet: 'Data',           // Source worksheet
    sourceRange: 'A1:C100',        // Source range
    targetSheet: 'Pivot',          // Target worksheet
    anchorCell: 'A3',              // Anchor cell
    
    layout: {
        rows: [                     // Row fields
            { name: 'Department' },
            { name: 'Product' }     // Multi-level row fields
        ],
        cols: [                     // Column fields
            { name: 'Month' },
            { name: 'Year' }
        ],
        values: [                   // Value fields
            { 
                name: 'Sales', 
                agg: 'sum',         // Aggregation: sum, avg, count, max, min
                displayName: 'Total Sales',
                numberFormat: '#,##0'
            },
            { 
                name: 'Quantity', 
                agg: 'count',
                displayName: 'Order Count'
            }
        ]
    },
    
    refreshOnLoad: true,            // Auto-refresh on open
    styleName: 'PivotStyleMedium9', // Pivot table style
    showGrandTotals: true,          // Show grand totals
    showSubTotals: true,            // Show subtotals
    enableDrilldown: true           // Enable drill-down
};
```

### 7. Chart Support

#### 7.1 Basic Charts

```typescript
// Create chart worksheet
const chartSheet = workbook.getWorksheet('Charts');

// Set chart data
chartSheet.setCell('A1', 'Month');
chartSheet.setCell('B1', 'Sales');
chartSheet.setCell('A2', 'Jan');
chartSheet.setCell('B2', 1000);
chartSheet.setCell('A3', 'Feb');
chartSheet.setCell('B3', 1200);
chartSheet.setCell('A4', 'Mar');
chartSheet.setCell('B4', 1100);

// Add chart (basic support)
chartSheet.addChart({
    type: 'bar',
    title: 'Monthly Sales Chart',
    dataRange: 'A1:B4',
    position: { x: 100, y: 100, width: 400, height: 300 }
});
```

#### 7.2 Chart Types

```typescript
// Supported chart types
const chartTypes = [
    'bar',          // Bar chart
    'line',         // Line chart
    'pie',          // Pie chart
    'column',       // Column chart
    'area',         // Area chart
    'scatter'       // Scatter chart
];

chartTypes.forEach((type, index) => {
    const row = index + 1;
    chartSheet.setCell(`A${row}`, `${type} Chart`);
    chartSheet.addChart({
        type: type,
        title: `${type} Chart Example`,
        dataRange: 'A1:B4',
        position: { x: 100, y: 100 + index * 100, width: 300, height: 200 }
    });
});
```

### 8. Advanced Features

#### 8.1 Formula Support

```typescript
// Set formulas
worksheet.setFormula('D4', '=B4*C4');           // Multiplication
worksheet.setFormula('D5', '=B5*C5');           // Multiplication
worksheet.setFormula('D6', '=SUM(D4:D5)');     // Sum
worksheet.setFormula('B6', '=SUM(B4:B5)');     // Quantity sum
worksheet.setFormula('C6', '=AVERAGE(C4:C5)'); // Average price

// Logical formulas
worksheet.setFormula('E4', '=IF(D4>100000,"High","Low")');
worksheet.setFormula('F4', '=AND(B4>5,C4>10000)');
```

#### 8.2 Conditional Formatting

```typescript
// Set conditional formatting (basic support)
worksheet.setCell('A1', 'Conditional Format Test', {
    font: { bold: true },
    fill: { type: 'solid', color: 'FFFF00' }
});

// Set styles based on values
const salesData = [1000, 1200, 800, 900, 1500];
salesData.forEach((value, index) => {
    const row = index + 1;
    const cell = worksheet.setCell(`B${row}`, value);
    
    // Set colors based on sales amount
    if (value > 1200) {
        cell.style = { fill: { type: 'solid', color: '00FF00' } }; // Green
    } else if (value > 1000) {
        cell.style = { fill: { type: 'solid', color: 'FFFF00' } }; // Yellow
    } else {
        cell.style = { fill: { type: 'solid', color: 'FF0000' } }; // Red
    }
});
```

#### 8.3 Performance Optimization

```typescript
// Large data processing
const largeData = [];
for (let i = 0; i < 10000; i++) {
    largeData.push({
        id: i + 1,
        name: `Item${i + 1}`,
        value: Math.random() * 1000
    });
}

// Batch processing
const batchSize = 1000;
for (let i = 0; i < largeData.length; i += batchSize) {
    const batch = largeData.slice(i, i + batchSize);
    batch.forEach((item, index) => {
        const row = i + index + 1;
        worksheet.setCell(`A${row}`, item.id);
        worksheet.setCell(`B${row}`, item.name);
        worksheet.setCell(`C${row}`, item.value);
    });
}
```

## 🔧 API Reference

### Workbook

| Method | Description | Status |
|--------|-------------|---------|
| `new Workbook()` | Create new workbook | ✅ Stable |
| `getWorksheet(name)` | Get worksheet | ✅ Stable |
| `writeBuffer()` | Output as Buffer | ✅ Stable |
| `writeFile(path)` | Save file directly | ✅ Stable |
| `writeFileWithPivotTables(path, options)` | Save file with pivot tables | ✅ Stable |
| `createManualPivotTable(data, options)` | Create manual pivot table | ✅ Stable |

### Worksheet

| Method | Description | Status |
|--------|-------------|---------|
| `setCell(address, value, options)` | Set cell | ✅ Stable |
| `getCell(address)` | Get cell | ✅ Stable |
| `mergeCells(range)` | Merge cells | ✅ Stable |
| `setColumnWidth(col, width)` | Set column width | ✅ Stable |
| `setRowHeight(row, height)` | Set row height | ✅ Stable |
| `freezePanes(row?, col?)` | Freeze panes | ✅ Stable |
| `protect(password, options)` | Protect worksheet | ✅ Stable |
| `addChart(chart)` | Add chart | 🔶 Experimental |

### Reading

| Method | Description | Status |
|--------|-------------|---------|
| `Workbook.readFile(path, options)` | Read from file | ✅ Stable |
| `Workbook.readBuffer(buffer, options)` | Read from Buffer | ✅ Stable |
| `worksheet.toArray()` | Convert to array | ✅ Stable |
| `worksheet.toJSON(options)` | Convert to JSON | ✅ Stable |

## 📖 Examples

### Complete Example: Sales Report System

```typescript
import { Workbook } from 'xml-xlsx-lite';

async function createSalesReport() {
    const workbook = new Workbook();
    
    // 1. Create data worksheet
    const dataSheet = workbook.getWorksheet('Sales Data');
    
    // Set headers
    dataSheet.setCell('A1', 'Date', { font: { bold: true } });
    dataSheet.setCell('B1', 'Product', { font: { bold: true } });
    dataSheet.setCell('C1', 'Quantity', { font: { bold: true } });
    dataSheet.setCell('D1', 'Unit Price', { font: { bold: true } });
    dataSheet.setCell('E1', 'Total Amount', { font: { bold: true } });
    
    // Fill in data
    const salesData = [
        ['2024-01-01', 'Laptop', 2, 35000, 70000],
        ['2024-01-01', 'Mouse', 10, 500, 5000],
        ['2024-01-02', 'Keyboard', 5, 800, 4000],
        ['2024-01-02', 'Monitor', 3, 8000, 24000],
        ['2024-01-03', 'Headphones', 8, 1200, 9600]
    ];
    
    salesData.forEach((row, index) => {
        const rowNum = index + 2;
        row.forEach((value, colIndex) => {
            const col = String.fromCharCode(65 + colIndex);
            dataSheet.setCell(`${col}${rowNum}`, value);
        });
        
        // Set formulas
        const rowNum2 = index + 2;
        dataSheet.setFormula(`E${rowNum2}`, `=C${rowNum2}*D${rowNum2}`);
    });
    
    // 2. Create pivot table
    const pivotSheet = workbook.getWorksheet('Pivot Analysis');
    workbook.createManualPivotTable(salesData.map(row => ({
        date: row[0],
        product: row[1],
        quantity: row[2],
        price: row[3],
        total: row[4]
    })), {
        rowField: 'product',
        columnField: 'date',
        valueField: 'total',
        aggregation: 'sum',
        numberFormat: '#,##0'
    });
    
    // 3. Create charts
    const chartSheet = workbook.getWorksheet('Charts');
    chartSheet.setCell('A1', 'Product Sales Chart', { font: { bold: true, size: 16 } });
    
    // 4. Save file
    await workbook.writeFileWithPivotTables('Sales Report.xlsx');
    
    console.log('Sales report created successfully!');
}

createSalesReport();
```

## 🚨 Important Notes

### ⚠️ Important Reminders

- **Do NOT use `writeFile()` method**: This method is not fully implemented, please use `writeBuffer()` + `fs.writeFileSync()` or the new `writeFileWithPivotTables()` method
- **Pivot Table Limitations**: Dynamic pivot tables need to be manually refreshed once in Excel
- **Browser Compatibility**: Some features (such as file reading) only support Node.js environment

### 🔧 Correct File Saving Methods

```typescript
// ❌ Wrong way
await workbook.writeFile('file.xlsx');

// ✅ Correct way 1: Use Buffer
const buffer = await workbook.writeBuffer();
const fs = require('fs');
fs.writeFileSync('file.xlsx', new Uint8Array(buffer));

// ✅ Correct way 2: Use new convenient method
await workbook.writeFileWithPivotTables('file.xlsx', pivotOptions);
```

## 📊 Feature Matrix

| Feature | Status | Description | Alternatives |
|---------|--------|-------------|--------------|
| **Basic Features** |
| Create Workbook | ✅ Stable | Fully supported | - |
| Cell Operations | ✅ Stable | Fully supported | - |
| Style Setting | ✅ Stable | Fully supported | - |
| Formula Support | ✅ Stable | Basic formulas | - |
| **Advanced Features** |
| Pivot Tables | 🔶 Experimental | Dynamic insertion | Manual creation |
| Chart Support | 🔶 Experimental | Basic support | Manual creation |
| File Reading | ✅ Stable | Fully supported | - |
| **Performance Optimization** |
| Large Data | ✅ Stable | Batch processing | Streaming processing |
| Memory Optimization | ✅ Stable | Auto-optimization | Manual control |

## 🌐 Browser Support

- ✅ **Node.js**: Fully supported
- 🔶 **Modern Browsers**: Basic feature support (some features limited)
- ❌ **Legacy Browsers**: Not supported

### Browser Usage Example

```html
<!DOCTYPE html>
<html>
<head>
    <title>xml-xlsx-lite Browser Test</title>
</head>
<body>
    <h1>Excel Generation Test</h1>
    <button onclick="generateExcel()">Generate Excel</button>
    
    <script type="module">
        import { Workbook } from './node_modules/xml-xlsx-lite/dist/index.esm.js';
        
        async function generateExcel() {
            const workbook = new Workbook();
            const worksheet = workbook.getWorksheet('Sheet1');
            
            worksheet.setCell('A1', 'Hello from Browser!');
            worksheet.setCell('B1', new Date());
            
            const buffer = await workbook.writeBuffer();
            
            // Download file
            const blob = new Blob([buffer], { 
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
            });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'browser-test.xlsx';
            a.click();
            URL.revokeObjectURL(url);
        }
    </script>
</body>
</html>
```

## 🤝 Contributing

We welcome contributions! Please see our [Contributing Guide](CONTRIBUTING.md) for details.

### Development Environment Setup

```bash
git clone https://github.com/mikemikex1/xml-xlsx-lite.git
cd xml-xlsx-lite
npm install
npm run dev
```

### Testing

```bash
npm run test:all        # Run all tests
npm run verify          # Verify functionality
npm run build           # Build project
```

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

---

## 🌟 Feature Showcase

### 🚀 Quick Start

```bash
# Install
npm install xml-xlsx-lite

# Basic usage
node -e "
const { Workbook } = require('xml-xlsx-lite');
const wb = new Workbook();
const ws = wb.getWorksheet('Sheet1');
ws.setCell('A1', 'Hello Excel!');
wb.writeBuffer().then(buf => require('fs').writeFileSync('test.xlsx', new Uint8Array(buf)));
"
```

### 📊 Pivot Table Example

```typescript
// Create complete report with pivot tables
const workbook = new Workbook();
const dataSheet = workbook.getWorksheet('Data');

// Fill in sales data
const salesData = [
    ['Department', 'Month', 'Product', 'Quantity', 'Amount'],
    ['IT', 'Jan', 'Laptop', 5, 175000],
    ['IT', 'Feb', 'Laptop', 3, 105000],
    ['HR', 'Jan', 'Office Supplies', 20, 4000],
    ['HR', 'Feb', 'Office Supplies', 15, 3000]
];

salesData.forEach((row, i) => {
    row.forEach((value, j) => {
        const address = String.fromCharCode(65 + j) + (i + 1);
        dataSheet.setCell(address, value);
    });
});

// Create manual pivot table
workbook.createManualPivotTable(salesData.slice(1).map(row => ({
    Department: row[0],
    Month: row[1],
    Product: row[2],
    Quantity: row[3],
    Amount: row[4]
})), {
    rowField: 'Department',
    columnField: 'Month',
    valueField: 'Amount',
    aggregation: 'sum'
});

// Save file
await workbook.writeFileWithPivotTables('Sales Pivot Report.xlsx');
```

---

**🎯 Goal**: Provide the most complete and easy-to-use Excel generation solution!

**💡 Features**: From basic operations to advanced features, complete guide from zero to hero!

**🚀 Vision**: Let every developer easily create professional Excel reports!
