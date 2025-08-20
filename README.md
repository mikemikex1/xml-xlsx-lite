# xml-xlsx-lite

[![npm version](https://badge.fury.io/js/xml-xlsx-lite.svg)](https://badge.fury.io/js/xml-xlsx-lite)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Minimal XLSX writer using raw XML + JSZip, inspired by exceljs API**

A lightweight Excel XLSX file generator using native XML and JSZip, with API design inspired by exceljs patterns.

## ✨ Features

- 🚀 **Lightweight**: Core functionality only, minimal dependencies
- 📝 **exceljs Compatible**: API design inspired by exceljs, low learning curve
- 🔧 **TypeScript Support**: Complete type definitions
- 🌐 **Cross-platform**: Supports Node.js and browser environments
- 📊 **Multiple Data Types**: Supports numbers, strings, booleans, dates
- 📋 **Multiple Worksheets**: Create and manage multiple worksheets
- 💾 **Shared Strings**: Automatic string deduplication for smaller file sizes
- 📊 **Format Preservation**: Maintains pivot tables, charts, and complex formatting when writing files

## 📦 Installation

```bash
npm install xml-xlsx-lite
```

## 🚀 Quick Start

> **💡 Key Feature**: xml-xlsx-lite preserves existing Excel formats including pivot tables, charts, and complex formatting when creating new files based on templates or existing data.

### Basic Usage

```javascript
import { Workbook } from 'xml-xlsx-lite';

// Create a new workbook
const wb = new Workbook();

// Get or create a worksheet
const ws = wb.getWorksheet("Sheet1");

// Set cell values
ws.setCell("A1", 123);
ws.setCell("B2", "Hello World");
ws.setCell("C3", true);
ws.setCell("D4", new Date());

// Generate XLSX file
const buffer = await wb.writeBuffer(); // ArrayBuffer
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
