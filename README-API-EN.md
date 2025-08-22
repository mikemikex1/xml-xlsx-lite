# 📚 xml-xlsx-lite API Reference

## 📖 Overview

This document provides comprehensive API reference for `xml-xlsx-lite`, a lightweight Excel XLSX file generator with complete Excel functionality support.

## 🚀 Core Classes

### Workbook

The main class for creating and managing Excel workbooks.

```typescript
class Workbook {
  constructor();
  
  // Worksheet Management
  getWorksheet(name: string): Worksheet;
  getWorksheets(): Worksheet[];
  
  // File Operations
  writeBuffer(): Promise<ArrayBuffer>;
  
  // Pivot Table Support
  createPivotTable(config: PivotTableConfig): PivotTable;
  getAllPivotTables(): PivotTable[];
  getPivotTable(name: string): PivotTable | undefined;
}
```

### Worksheet

Represents a single worksheet within a workbook.

```typescript
class Worksheet {
  // Cell Operations
  setCell(address: string, value: CellValue, options?: CellOptions): void;
  getCell(address: string): CellValue | undefined;
  
  // Column and Row Management
  setColumnWidth(column: string, width: number): void;
  setRowHeight(row: number, height: number): void;
  
  // Protection
  protect(options: ProtectionOptions): void;
  
  // Pivot Table Export
  exportToWorksheet(name: string): Worksheet;
}
```

## 📊 Pivot Table System

### PivotTable Interface

```typescript
interface PivotTable {
  name: string;
  config: PivotTableConfig;
  
  // Data Management
  setSourceData(data: any[][]): void;
  refresh(): void;
  getData(): any[][];
  
  // Field Management
  addField(field: PivotField): void;
  removeField(fieldName: string): void;
  reorderFields(fieldOrder: string[]): void;
  
  // Filtering
  applyFilter(fieldName: string, filterValues: any[]): void;
  clearFilters(): void;
  
  // Export
  exportToWorksheet(worksheetName: string): Worksheet;
}
```

### PivotTableConfig

```typescript
interface PivotTableConfig {
  name: string;
  sourceRange: string;
  targetRange: string;
  fields: PivotField[];
  
  // Display Options
  showRowHeaders: boolean;
  showColumnHeaders: boolean;
  showRowSubtotals: boolean;
  showColumnSubtotals: boolean;
  showGrandTotals: boolean;
  
  // Styling
  autoFormat: boolean;
  compactRows: boolean;
  outlineData: boolean;
  mergeLabels: boolean;
}
```

### PivotField

```typescript
interface PivotField {
  name: string;
  sourceColumn: string;
  type: 'row' | 'column' | 'value' | 'filter';
  
  // Aggregation (for value fields)
  function?: 'sum' | 'count' | 'average' | 'max' | 'min';
  customName?: string;
  
  // Display Options
  showSubtotal: boolean;
  showGrandTotal: boolean;
}
```

## 🔧 Dynamic Pivot Table Builder

### addPivotToWorkbookBuffer

Dynamically insert native Excel pivot tables into existing workbooks.

```typescript
async function addPivotToWorkbookBuffer(
  workbookBuf: Buffer, 
  opt: CreatePivotOptions
): Promise<Buffer>;
```

### CreatePivotOptions

```typescript
interface CreatePivotOptions {
  sourceSheet: string;     // Source worksheet name
  sourceRange: string;     // Data range (A1:D100)
  targetSheet: string;     // Target worksheet name
  anchorCell: string;      // Pivot table anchor (A3)
  layout: PivotLayout;     // Field configuration
  refreshOnLoad?: boolean; // Auto-refresh on open (default: true)
  styleName?: string;      // Style name (default: PivotStyleMedium9)
}
```

### PivotLayout

```typescript
interface PivotLayout {
  rows?: PivotFieldSpec[];    // Row fields (optional)
  cols?: PivotFieldSpec[];    // Column fields (optional)
  values: PivotValueSpec[];   // Value fields (required)
}
```

### PivotFieldSpec

```typescript
interface PivotFieldSpec {
  name: string;  // Field name (must match source data headers)
}
```

### PivotValueSpec

```typescript
interface PivotValueSpec {
  name: string;            // Field name (must be numeric)
  agg?: PivotAgg;         // Aggregation method (default: sum)
  displayName?: string;    // Display name (default: field name)
  numFmtId?: number;       // Number format ID (default: 0)
}
```

### PivotAgg

```typescript
type PivotAgg = "sum" | "count" | "average" | "max" | "min" | "product";
```

## 🎨 Cell Styling

### CellOptions

```typescript
interface CellOptions {
  // Font Styling
  font?: {
    bold?: boolean;
    italic?: boolean;
    size?: number;
    color?: string;
    name?: string;
  };
  
  // Alignment
  alignment?: {
    horizontal?: 'left' | 'center' | 'right';
    vertical?: 'top' | 'middle' | 'bottom';
    wrapText?: boolean;
  };
  
  // Background
  fill?: {
    type: 'pattern' | 'gradient';
    color: string;
  };
  
  // Borders
  border?: {
    top?: BorderStyle;
    bottom?: BorderStyle;
    left?: BorderStyle;
    right?: BorderStyle;
  };
  
  // Number Format
  numFmt?: string;
}
```

### BorderStyle

```typescript
interface BorderStyle {
  style: 'thin' | 'medium' | 'thick';
  color: string;
}
```

## 🔒 Protection System

### ProtectionOptions

```typescript
interface ProtectionOptions {
  password?: string;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  formatCells?: boolean;
  insertRows?: boolean;
  deleteRows?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
}
```

## 📈 Chart System

### ChartImpl

```typescript
class ChartImpl implements Chart {
  name: string;
  type: ChartType;
  data: ChartData;
  
  // Configuration
  setData(data: ChartData): void;
  setOptions(options: ChartOptions): void;
  
  // Export
  exportToWorksheet(worksheetName: string): Worksheet;
}
```

### ChartType

```typescript
type ChartType = 
  | 'bar' 
  | 'line' 
  | 'pie' 
  | 'doughnut' 
  | 'area' 
  | 'scatter' 
  | 'bubble' 
  | 'radar';
```

### ChartData

```typescript
interface ChartData {
  categories: string[];
  series: ChartSeries[];
}

interface ChartSeries {
  name: string;
  data: number[];
  color?: string;
}
```

## 🚀 Performance Optimization

### PerformanceOptimizer

```typescript
class PerformanceOptimizer {
  // Analysis
  analyzeWorksheet(worksheet: any): PerformanceStats;
  
  // Optimization Decisions
  shouldUseSharedStrings(): boolean;
  shouldUseStreaming(): boolean;
  shouldOptimizeMemory(): boolean;
  
  // Configuration
  getConfig(): PerformanceConfig;
  updateConfig(newConfig: Partial<PerformanceConfig>): void;
}
```

### PerformanceConfig

```typescript
interface PerformanceConfig {
  sharedStringsThreshold: number;      // Threshold for sharedStrings
  repetitionRateThreshold: number;     // Repetition rate threshold (%)
  largeFileThreshold: number;          // Large file threshold
  streamingThreshold: number;          // Streaming threshold (MB)
  cacheSizeLimit: number;              // Cache size limit (MB)
  memoryOptimization: boolean;         // Memory optimization flag
}
```

### StreamingProcessor

```typescript
class StreamingProcessor {
  // Chunk Processing
  async processInChunks<T>(
    data: T[], 
    processor: (chunk: T[]) => Promise<void>
  ): Promise<void>;
  
  // Progress Tracking
  setProgressCallback(callback: (progress: number) => void): void;
}
```

### CacheManager

```typescript
class CacheManager {
  // Cache Operations
  get(key: string): any | undefined;
  set(key: string, value: any): void;
  clear(): void;
  
  // Statistics
  getStats(): { size: number; maxSize: number; hitRate: number };
}
```

## 📖 Reading System

### WorkbookReader

```typescript
interface WorkbookReader {
  readFile(path: string, options?: ReadOptions): Promise<Workbook>;
  readBuffer(buf: ArrayBuffer, options?: ReadOptions): Promise<Workbook>;
  validateFile(path: string): Promise<ValidationResult>;
}
```

### WorksheetReader

```typescript
interface WorksheetReader {
  toArray(): CellValue[][];
  toJSON(opts?: { headerRow?: number }): Record<string, CellValue>[];
  getRange(range: string): CellValue[][];
  getRow(row: number): CellValue[];
  getColumn(col: string | number): CellValue[];
}
```

### ReadOptions

```typescript
interface ReadOptions {
  includeHiddenSheets?: boolean;
  includeFormulas?: boolean;
  includeStyles?: boolean;
  maxRows?: number;
  maxColumns?: number;
}
```

## 🚨 Error Handling

### Error Classes

```typescript
// Address and Range Errors
export class InvalidAddressError extends Error
export class InvalidRangeError extends Error

// Type and Value Errors
export class UnsupportedTypeError extends Error
export class InvalidValueError extends Error

// File and Format Errors
export class CorruptedFileError extends Error
export class UnsupportedFormatError extends Error

// Feature and Operation Errors
export class UnsupportedFeatureError extends Error
export class OperationNotAllowedError extends Error

// Validation and Performance Warnings
export class ValidationError extends Error
export class PerformanceWarning extends Error
```

### Error Creation Helpers

```typescript
// Create standardized errors with context
createInvalidAddressError(address: string, context?: string): InvalidAddressError;
createUnsupportedTypeError(value: any, expectedTypes: string[]): UnsupportedTypeError;
createCorruptedFileError(filePath: string, details: string): CorruptedFileError;
createValidationError(field: string, value: any, rule: string): ValidationError;
createPerformanceWarning(message: string, threshold: number): PerformanceWarning;
```

## 🔧 Utility Functions

### Address Conversion

```typescript
// Convert between A1 notation and row/column numbers
colToA1(col: number): string;
a1ToRC(address: string): { row: number; col: number };
addr(row: number, col: number): string;

// Range operations
parseRange(range: string): RangeInfo;
expandRange(range: string): string[];
```

### Type Utilities

```typescript
// Cell type detection
getCellType(value: any): 'n' | 's' | 'b' | 'd' | 'inlineStr' | null;

// Value validation
isNumeric(value: any): boolean;
isDate(value: any): boolean;
isBoolean(value: any): boolean;
```

## 📋 Usage Examples

### Basic Workbook Creation

```typescript
import { Workbook } from 'xml-xlsx-lite';

const wb = new Workbook();
const ws = wb.getWorksheet('Data');

// Add data
ws.setCell('A1', 'Name', { font: { bold: true } });
ws.setCell('B1', 'Value', { font: { bold: true } });
ws.setCell('A2', 'Item 1');
ws.setCell('B2', 100, { numFmt: '#,##0' });

// Save
const buffer = await wb.writeBuffer();
```

### Dynamic Pivot Table

```typescript
import { addPivotToWorkbookBuffer } from 'xml-xlsx-lite';

const pivotOptions = {
  sourceSheet: "Data",
  sourceRange: "A1:D100",
  targetSheet: "Pivot",
  anchorCell: "A3",
  layout: {
    rows: [{ name: "Department" }],
    cols: [{ name: "Month" }],
    values: [{ name: "Sales", agg: "sum" }]
  }
};

const enhancedBuffer = await addPivotToWorkbookBuffer(baseBuffer, pivotOptions);
```

### Advanced Styling

```typescript
ws.setCell('A1', 'Styled Cell', {
  font: { 
    bold: true, 
    size: 16, 
    color: 'FF0000' 
  },
  alignment: { 
    horizontal: 'center', 
    vertical: 'middle' 
  },
  fill: { 
    type: 'pattern', 
    color: 'E0E0E0' 
  },
  border: {
    top: { style: 'thick', color: '000000' },
    bottom: { style: 'thick', color: '000000' }
  }
});
```

## 🌟 Best Practices

### Performance Optimization

1. **Use appropriate cell types**: Choose the right cell type for your data
2. **Batch operations**: Group cell operations when possible
3. **Memory management**: Use streaming for large files
4. **Cache wisely**: Implement appropriate caching strategies

### Error Handling

1. **Always validate inputs**: Check data types and ranges before processing
2. **Use specific error types**: Choose the most appropriate error class
3. **Provide context**: Include relevant information in error messages
4. **Handle gracefully**: Implement fallback behavior when possible

### Data Management

1. **Plan your structure**: Design worksheets with clear organization
2. **Use consistent formatting**: Apply styles consistently across similar data
3. **Validate data**: Ensure data integrity before processing
4. **Document your schema**: Keep track of data structure and relationships

## 📚 Related Documentation

- [Main README](./README.md)
- [Dynamic Pivot Tables](./DYNAMIC_PIVOT_USAGE.md)
- [Implementation Report](./IMPLEMENTATION_REPORT.md)
- [Usage Guide](./USAGE_GUIDE_FIXED.md)

---

**xml-xlsx-lite API Reference** - Your complete guide to Excel automation! 🚀
