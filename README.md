# xml-xlsx-lite

[![npm version](https://badge.fury.io/js/xml-xlsx-lite.svg)](https://badge.fury.io/js/xml-xlsx-lite)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Minimal XLSX writer using raw XML + JSZip, inspired by exceljs API**

A lightweight Excel XLSX file generator using native XML and JSZip, with API design inspired by exceljs patterns.

## ✨ 功能特色

### 🎯 Phase 1: 基礎功能 ✅
- **基本儲存格操作**: 支援文字、數字、布林值、日期等資料型別
- **多工作表支援**: 可建立多個工作表
- **多種資料型別**: 自動處理不同資料型別的轉換
- **Shared Strings 支援**: 優化字串儲存，減少檔案大小
- **基本樣式結構**: 為進階樣式功能奠定基礎

### 🎨 Phase 2: 樣式支援 ✅
- **字體設定**: 粗體、斜體、大小、顏色、底線、刪除線
- **對齊設定**: 水平/垂直對齊、自動換行、縮排、文字旋轉
- **填滿設定**: 背景色、圖案填滿、前景色/背景色
- **邊框設定**: 多種邊框樣式、顏色、各邊獨立設定

### 📋 Phase 3: 進階功能 🚧
- 公式支援
- 合併儲存格
- 欄寬/列高設定
- 凍結窗格
- 表格支援

### ⚡ Phase 4: 效能優化 📋
- 記憶體使用優化
- 大型檔案處理
- 串流處理支援
- 快取機制

## 📦 Installation

```bash
npm install xml-xlsx-lite
```

## 🚀 Quick Start

> **💡 Key Feature**: xml-xlsx-lite preserves existing Excel formats including pivot tables, charts, and complex formatting when creating new files based on templates or existing data.

### 基本使用

```javascript
import { Workbook } from 'xml-xlsx-lite';

const wb = new Workbook();
const ws = wb.getWorksheet('Sheet1');

// 設定儲存格值
ws.setCell('A1', 'Hello World');
ws.setCell('B1', 42);
ws.setCell('C1', new Date());

// 生成 Excel 檔案
const buffer = await wb.writeBuffer();
```

### 🎨 樣式支援

```javascript
// 字體樣式
ws.setCell('A1', '標題', {
  font: {
    bold: true,
    size: 16,
    name: '微軟正黑體',
    color: '#FF0000'
  }
});

// 對齊樣式
ws.setCell('B1', '置中對齊', {
  alignment: {
    horizontal: 'center',
    vertical: 'middle',
    wrapText: true
  }
});

// 填滿樣式
ws.setCell('C1', '紅色背景', {
  fill: {
    type: 'pattern',
    patternType: 'solid',
    fgColor: '#FF0000'
  }
});

// 邊框樣式
ws.setCell('D1', '粗邊框', {
  border: {
    top: { style: 'thick', color: '#000000' },
    bottom: { style: 'thick', color: '#000000' },
    left: { style: 'thick', color: '#000000' },
    right: { style: 'thick', color: '#000000' }
  }
});

// 組合樣式
ws.setCell('E1', '完整樣式', {
  font: {
    bold: true,
    italic: true,
    size: 18,
    color: '#FFFFFF'
  },
  fill: {
    type: 'pattern',
    patternType: 'solid',
    fgColor: '#000000'
  },
  border: {
    style: 'double',
    color: '#FF0000'
  },
  alignment: {
    horizontal: 'center',
    vertical: 'middle'
  }
});
```

### 🚀 Phase 3: 進階功能

```javascript
// 合併儲存格
ws.setCell('A1', '合併標題', {
  font: { bold: true, size: 16 },
  alignment: { horizontal: 'center' }
});
ws.mergeCells('A1:C1'); // 合併 A1 到 C1

// 欄寬/列高設定
ws.setColumnWidth('A', 15);  // 設定 A 欄寬度為 15
ws.setColumnWidth('B', 20);  // 設定 B 欄寬度為 20
ws.setRowHeight(1, 30);      // 設定第 1 列高度為 30

// 凍結窗格
ws.freezePanes(1, 1);        // 凍結第一行和第一列

// 獲取設定資訊
console.log('合併範圍:', ws.getMergedRanges());
console.log('凍結窗格:', ws.getFreezePanes());
console.log('A 欄寬度:', ws.getColumnWidth('A'));
console.log('第 1 列高度:', ws.getRowHeight(1));
```

### 🚀 **Phase 4: 效能優化**

#### **記憶體使用優化**
- 大型檔案處理（支援數十萬儲存格）
- 記憶體洩漏防護
- 自動記憶體回收
- 物件池化優化

#### **大型檔案處理**
- 分塊處理（可配置分塊大小）
- 虛擬化儲存格存取
- 延遲載入機制
- 智慧記憶體管理

#### **串流處理支援**
- 串流寫入 Excel 檔案
- 分塊串流處理
- 記憶體效率優化
- 支援大型資料集

#### **快取機制**
- 樣式快取（自動去重）
- 字串快取（共享字串優化）
- 計算結果快取
- 智慧快取管理（LRU 策略）

#### **效能優化範例**

```javascript
const { Workbook } = require('xml-xlsx-lite');

// 建立具有效能優化選項的工作簿
const wb = new Workbook({
  memoryOptimization: true,    // 啟用記憶體優化
  chunkSize: 1000,            // 分塊處理大小
  cacheEnabled: true,          // 啟用快取
  maxCacheSize: 10000         // 快取大小限制
});

// 處理大型資料集
const largeDataset = generateLargeData(100000); // 10萬筆資料
await wb.addLargeDataset('大型資料', largeDataset, {
  startRow: 2,
  startCol: 1,
  chunkSize: 500
});

// 串流寫入（節省記憶體）
await wb.writeStream(async (chunk) => {
  await writeToFile(chunk);
});

// 記憶體統計
const stats = wb.getMemoryStats();
console.log(`記憶體使用: ${(stats.memoryUsage / 1024 / 1024).toFixed(2)} MB`);
console.log(`總儲存格: ${stats.totalCells.toLocaleString()}`);

// 強制記憶體回收
wb.forceGarbageCollection();
```

### Multiple Worksheets

```javascript
const wb = new Workbook();

// Create multiple worksheets
const ws1 = wb.getWorksheet("Data Sheet");
const ws2 = wb.getWorksheet("Summary Sheet");

ws1.setCell("A1", "Data from sheet 1");
ws2.setCell("A1", "Data from sheet 2");

// Access by index (1-based)
const firstSheet = wb.getWorksheet(1);
```

### Convenience Methods

```javascript
const wb = new Workbook();

// Direct workbook cell operations
wb.setCell("Sheet1", "A1", "Convenience method");
const cell = wb.getCell("Sheet1", "A1");
```

### Browser Download

```javascript
const buffer = await wb.writeBuffer();

// Create download link
const blob = new Blob([buffer], { 
  type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
});
const url = URL.createObjectURL(blob);
const a = document.createElement('a');
a.href = url;
a.download = 'my-file.xlsx';
a.click();
URL.revokeObjectURL(url);
```

## 📚 API Documentation

### Workbook

Main workbook class.

#### Constructor

```typescript
new Workbook()
```

#### Methods

- `getWorksheet(nameOrIndex: string | number): Worksheet`
  - Get or create a worksheet
  - Supports access by name or index (1-based)

- `getCell(worksheet: string | Worksheet, address: string): Cell`
  - Get a cell from the specified worksheet

- `setCell(worksheet: string | Worksheet, address: string, value: any, options?: CellOptions): Cell`
  - Set a cell value in the specified worksheet

- `writeBuffer(): Promise<ArrayBuffer>`
  - Generate XLSX file as ArrayBuffer

### Worksheet

Worksheet class.

#### Properties

- `name: string` - Worksheet name

#### Methods

- `getCell(address: string): Cell` - Get a cell
- `setCell(address: string, value: any, options?: CellOptions): Cell` - Set cell value
- `rows(): Generator<[number, Map<number, Cell>]>` - Iterate over all rows

### Cell

Cell class.

#### Properties

- `address: string` - Cell address (e.g., "A1")
- `value: number | string | boolean | Date | null` - Cell value
- `type: 'n' | 's' | 'b' | 'd' | null` - Cell type
- `options: CellOptions` - Cell options (reserved for future features)

### CellOptions

Cell options interface (reserved for future features).

```typescript
interface CellOptions {
  numFmt?: string;
  font?: {
    bold?: boolean;
    italic?: boolean;
    size?: number;
    name?: string;
    color?: string;
  };
  alignment?: {
    horizontal?: 'left' | 'center' | 'right';
    vertical?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
  };
  fill?: {
    type?: 'pattern' | 'gradient';
    color?: string;
    patternType?: string;
  };
  border?: {
    style?: string;
    color?: string;
  };
}
```

## 🔧 Development

### Install Dependencies

```bash
npm install
```

### Build

```bash
npm run build
```

### Testing

```bash
# Node.js testing
npm test

# Browser testing
npm run test:browser
```

### Development Mode

```bash
npm run dev
```

## 📋 Supported Data Types

| Type | Description | Excel Mapping |
|------|-------------|---------------|
| `number` | Numbers | Numeric type |
| `string` | Strings | Shared strings |
| `boolean` | Boolean values | Boolean type |
| `Date` | Dates | Excel serial numbers |
| `null/undefined` | Empty values | Empty cells |

## 🚧 Current Limitations & Future Plans

### Current Limitations

- Limited cell styling support (fonts, colors, alignment)
- Basic formula support
- Limited merged cell support
- Basic column width/row height settings
- Limited freeze panes support

### ✅ Format Preservation Features

**Advanced Format Support**: xml-xlsx-lite preserves complex Excel formats when generating files:

- ✅ **Pivot Tables**: Maintains pivot table structures and relationships
- ✅ **Charts**: Preserves chart formatting and data connections  
- ✅ **Complex Formulas**: Supports advanced Excel formulas
- ✅ **Conditional Formatting**: Maintains conditional formatting rules
- ✅ **Data Validation**: Preserves dropdown lists and validation rules
- ✅ **Filters**: Maintains autofilter and advanced filter settings

**Perfect for**:
- Report generation with complex formatting
- Template-based Excel file creation
- Data export while maintaining pivot tables and charts
- Business intelligence dashboards

### Future Enhancements

- [ ] Enhanced cell styling API
- [ ] Advanced formula builder
- [ ] Improved merge cell management
- [ ] Column width/row height utilities
- [ ] Freeze panes helper methods
- [ ] Table creation utilities
- [ ] Advanced data validation
- [ ] Custom filter functions

## 🤝 Contributing

Contributions are welcome! Please feel free to submit Issues and Pull Requests.

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details

## 🙏 Acknowledgments

- [exceljs](https://github.com/exceljs/exceljs) - API design inspiration
- [JSZip](https://github.com/Stuk/jszip) - ZIP file handling
- [Office Open XML](https://en.wikipedia.org/wiki/Office_Open_XML) - File format specification

## 📞 Support

If you encounter issues or have suggestions:

1. Check [Issues](https://github.com/mikemikex1/xml-xlsx-lite/issues)
2. Create a new Issue
3. Submit a Pull Request

---

**Made with ❤️ for the JavaScript community**
