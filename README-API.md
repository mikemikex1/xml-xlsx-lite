# xml-xlsx-lite API 規格文檔

## 📋 目錄
- [核心介面](#核心介面)
- [儲存格相關](#儲存格相關)
- [工作表相關](#工作表相關)
- [工作簿相關](#工作簿相關)
- [Pivot Table 相關](#pivot-table-相關)
- [保護功能相關](#保護功能相關)
- [圖表相關](#圖表相關)
- [樣式相關](#樣式相關)
- [效能優化相關](#效能優化相關)

---

## 🏗️ 核心介面

### `Workbook`
工作簿的主要介面，提供工作表的創建、管理和 Excel 檔案生成功能。

**屬性：**
- 無公開屬性

**方法：**
```typescript
// 工作表管理
getWorksheet(nameOrIndex: string | number): Worksheet
getWorksheets(): Worksheet[]

// 儲存格操作
getCell(worksheet: string | Worksheet, address: string): Cell
setCell(worksheet: string | Worksheet, address: string, value: number | string | boolean | Date | null, options?: CellOptions): Cell

// 檔案輸出
writeBuffer(): Promise<ArrayBuffer>
writeFile(filename: string): Promise<void>
writeStream(writeStream: (chunk: Uint8Array) => Promise<void>): Promise<void>

// 大型資料集處理
addLargeDataset(worksheetName: string, data: Array<Array<any>>, options?: {
  startRow?: number;
  startCol?: number;
  chunkSize?: number;
}): Promise<void>

// 記憶體優化設定
setMemoryOptimization(enabled: boolean): void
setChunkSize(size: number): void
setCacheEnabled(enabled: boolean): void
setMaxCacheSize(size: number): void

// 記憶體統計
getMemoryStats(): {
  sheets: number;
  totalCells: number;
  cacheSize: number;
  cacheHitRate: number;
  memoryUsage: number;
}
forceGarbageCollection(): void

// Pivot Table 管理
createPivotTable(config: PivotTableConfig): PivotTable
getPivotTable(name: string): PivotTable | undefined
getAllPivotTables(): PivotTable[]
removePivotTable(name: string): boolean
refreshAllPivotTables(): void

// 工作簿保護
protect(password?: string, options?: WorkbookProtectionOptions): void
unprotect(password?: string): void
isProtected(): boolean
getProtectionOptions(): WorkbookProtectionOptions | null
```

**使用範例：**
```typescript
const workbook = new Workbook();
const sheet = workbook.getWorksheet('Sheet1');
sheet.setCell('A1', 'Hello World');
const buffer = await workbook.writeBuffer();
```

---

## 📝 儲存格相關

### `Cell`
表示工作表中的單個儲存格。

**屬性：**
```typescript
address: string                                // 儲存格地址 (如 "A1")
value: number | string | boolean | Date | null // 儲存格的值
type: 'n' | 's' | 'b' | 'd' | null          // 儲存格類型 (n=數字, s=字串, b=布林, d=日期)
options: CellOptions                          // 儲存格選項
```

**使用範例：**
```typescript
const cell = sheet.getCell('A1');
console.log(cell.value);        // 獲取值
console.log(cell.address);      // 獲取地址
console.log(cell.type);         // 獲取類型
```

### `CellOptions`
定義儲存格的所有樣式和格式選項。

**屬性：**
```typescript
// 數字格式
numFmt?: string

// 字體設定
font?: {
  bold?: boolean           // 粗體
  italic?: boolean         // 斜體
  size?: number            // 字體大小
  name?: string            // 字體名稱
  color?: string           // 字體顏色
  underline?: boolean      // 底線
  strike?: boolean         // 刪除線
}

// 對齊設定
alignment?: {
  horizontal?: 'left' | 'center' | 'right' | 'justify' | 'distributed'  // 水平對齊
  vertical?: 'top' | 'middle' | 'bottom' | 'justify' | 'distributed'    // 垂直對齊
  wrapText?: boolean       // 自動換行
  indent?: number          // 縮排
  rotation?: number        // 旋轉角度
}

// 填滿設定
fill?: {
  type?: 'pattern' | 'gradient'           // 填滿類型
  color?: string                          // 主要顏色
  patternType?: 'none' | 'solid' | 'darkGray' | 'mediumGray' | 'lightGray' | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp' | 'darkGrid' | 'darkTrellis' | 'lightHorizontal' | 'lightVertical' | 'lightDown' | 'lightUp' | 'lightGrid' | 'lightTrellis' | 'gray125' | 'gray0625'
  fgColor?: string                        // 前景顏色
  bgColor?: string                        // 背景顏色
}

// 邊框設定
border?: {
  style?: 'none' | 'thin' | 'medium' | 'dashed' | 'dotted' | 'thick' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot'
  color?: string                          // 邊框顏色
  top?: { style?: string; color?: string }     // 上邊框
  left?: { style?: string; color?: string }    // 左邊框
  bottom?: { style?: string; color?: string }  // 下邊框
  right?: { style?: string; color?: string }   // 右邊框
}

// 合併儲存格
mergeRange?: string        // 標記儲存格是否為合併儲存格的主儲存格
mergedInto?: string        // 標記儲存格是否被合併到某個範圍

// 公式支援
formula?: string           // Excel 公式，例如 "=SUM(A1:A10)"
formulaType?: 'array' | 'shared' | 'dataTable'  // 公式類型

// Pivot Table 支援
pivotTable?: {
  isPivotField?: boolean           // 是否為樞紐欄位
  pivotFieldType?: 'row' | 'column' | 'value' | 'filter'  // 樞紐欄位類型
  pivotFieldIndex?: number         // 樞紐欄位索引
  pivotItemIndex?: number          // 樞紐項目索引
  isSubtotal?: boolean             // 是否為小計
  isGrandTotal?: boolean           // 是否為總計
}
```

**使用範例：**
```typescript
const cellOptions: CellOptions = {
  font: { bold: true, size: 14, color: '#FF0000' },
  alignment: { horizontal: 'center', vertical: 'middle' },
  fill: { type: 'pattern', patternType: 'solid', fgColor: '#FFFF00' },
  border: { style: 'thin', color: '#000000' }
};

sheet.setCell('A1', 'Hello World', cellOptions);
```

---

## 📊 工作表相關

### `Worksheet`
表示工作簿中的單個工作表。

**屬性：**
```typescript
name: string  // 工作表名稱
```

**方法：**
```typescript
// 儲存格操作
getCell(address: string): Cell
setCell(address: string, value: number | string | boolean | Date | null, options?: CellOptions): Cell

// 行/列遍歷
rows(): Generator<[number, Map<number, Cell>]>

// 合併儲存格
mergeCells(range: string): void
unmergeCells(range: string): void
getMergedRanges(): string[]

// 欄寬/列高設定
setColumnWidth(column: string | number, width: number): void
getColumnWidth(column: string | number): number
setRowHeight(row: number, height: number): void
getRowHeight(row: number): number

// 凍結窗格
freezePanes(row?: number, column?: number): void
unfreezePanes(): void
getFreezePanes(): { row?: number; column?: number }

// 公式支援
setFormula(address: string, formula: string, options?: CellOptions): Cell
getFormula(address: string): string | null
validateFormula(formula: string): boolean
getFormulaDependencies(address: string): string[]

// 工作表保護
protect(password?: string, options?: WorksheetProtectionOptions): void
unprotect(password?: string): void
isProtected(): boolean
getProtectionOptions(): WorksheetProtectionOptions | null

// 圖表支援
addChart(chart: Chart): void
removeChart(chartName: string): void
getCharts(): Chart[]
getChart(chartName: string): Chart | undefined
```

**使用範例：**
```typescript
const sheet = workbook.getWorksheet('Sheet1');

// 設定儲存格
sheet.setCell('A1', 'Hello World');

// 合併儲存格
sheet.mergeCells('A1:B2');

// 設定欄寬
sheet.setColumnWidth('A', 15);

// 凍結窗格
sheet.freezePanes(2, 1);

// 設定公式
sheet.setFormula('B1', '=SUM(A1:A10)');
```

---

## 🔄 Pivot Table 相關

### `PivotField`
定義樞紐分析表中的欄位設定。

**屬性：**
```typescript
name: string                                    // 欄位名稱
sourceColumn: string                            // 來源欄位名稱
type: 'row' | 'column' | 'value' | 'filter'    // 欄位類型
function?: 'sum' | 'count' | 'average' | 'max' | 'min' | 'countNums' | 'stdDev' | 'stdDevP' | 'var' | 'varP'  // 彙總函數
numberFormat?: string                           // 數字格式
showSubtotal?: boolean                          // 是否顯示小計
showGrandTotal?: boolean                        // 是否顯示總計
sortOrder?: 'asc' | 'desc'                     // 排序順序
filterValues?: string[]                         // 篩選值
customName?: string                             // 自訂名稱
```

**使用範例：**
```typescript
const rowField: PivotField = {
  name: 'Month',
  sourceColumn: 'Month',
  type: 'row',
  showSubtotal: true,
  sortOrder: 'asc'
};

const valueField: PivotField = {
  name: 'Saving Amount',
  sourceColumn: 'Saving_Amount',
  type: 'value',
  function: 'sum',
  numberFormat: '#,##0.00',
  customName: 'Total Savings'
};
```

### `PivotTableConfig`
定義樞紐分析表的配置。

**屬性：**
```typescript
name: string                    // 樞紐分析表名稱
sourceRange: string             // 資料來源範圍，例如 "A1:D1000"
targetRange: string             // 目標範圍，例如 "F1:J20"
fields: PivotField[]            // 欄位設定陣列
showRowHeaders?: boolean        // 是否顯示列標題
showColumnHeaders?: boolean     // 是否顯示欄標題
showRowSubtotals?: boolean      // 是否顯示列小計
showColumnSubtotals?: boolean   // 是否顯示欄小計
showGrandTotals?: boolean       // 是否顯示總計
autoFormat?: boolean            // 是否自動格式化
compactRows?: boolean           // 是否壓縮列
outlineData?: boolean           // 是否顯示大綱資料
mergeLabels?: boolean           // 是否合併標籤
pageBreakBetweenGroups?: boolean // 群組間是否分頁
repeatRowLabels?: boolean       // 是否重複列標籤
rowGrandTotals?: boolean        // 是否顯示列總計
columnGrandTotals?: boolean     // 是否顯示欄總計
```

**使用範例：**
```typescript
const pivotConfig: PivotTableConfig = {
  name: 'Savings Summary',
  sourceRange: 'A1:C7',
  targetRange: 'E1:H10',
  fields: [rowField, valueField],
  showGrandTotals: true,
  autoFormat: true
};
```

### `PivotTable`
樞紐分析表的實例介面。

**屬性：**
```typescript
name: string                    // 樞紐分析表名稱
config: PivotTableConfig        // 配置設定
```

**方法：**
```typescript
refresh(): void                 // 重新整理資料
updateSourceData(sourceRange: string): void  // 更新資料來源
getField(fieldName: string): PivotField | undefined  // 獲取欄位
addField(field: PivotField): void            // 添加欄位
removeField(fieldName: string): void         // 移除欄位
reorderFields(fieldOrder: string[]): void    // 重新排序欄位
applyFilter(fieldName: string, filterValues: string[]): void  // 套用篩選
clearFilters(): void            // 清除篩選
getData(): any[][]              // 獲取樞紐資料
exportToWorksheet(worksheetName: string): Worksheet  // 匯出到工作表
```

**使用範例：**
```typescript
const pivotTable = workbook.createPivotTable(pivotConfig);

// 重新整理
pivotTable.refresh();

// 套用篩選
pivotTable.applyFilter('Month', ['January', 'February']);

// 匯出到新工作表
const newSheet = pivotTable.exportToWorksheet('Pivot Results');
```

---

## 🔒 保護功能相關

### `WorksheetProtectionOptions`
工作表保護選項。

**屬性：**
```typescript
selectLockedCells?: boolean     // 是否允許選取鎖定的儲存格
selectUnlockedCells?: boolean   // 是否允許選取未鎖定的儲存格
formatCells?: boolean           // 是否允許格式化儲存格
formatColumns?: boolean         // 是否允許格式化欄
formatRows?: boolean            // 是否允許格式化列
insertColumns?: boolean         // 是否允許插入欄
insertRows?: boolean            // 是否允許插入列
insertHyperlinks?: boolean      // 是否允許插入超連結
deleteColumns?: boolean         // 是否允許刪除欄
deleteRows?: boolean            // 是否允許刪除列
sort?: boolean                  // 是否允許排序
autoFilter?: boolean            // 是否允許自動篩選
pivotTables?: boolean           // 是否允許樞紐分析表
objects?: boolean               // 是否允許物件操作
scenarios?: boolean             // 是否允許情節
```

**使用範例：**
```typescript
const protectionOptions: WorksheetProtectionOptions = {
  selectLockedCells: false,
  formatCells: false,
  insertRows: false,
  deleteRows: false
};

sheet.protect('password123', protectionOptions);
```

### `WorkbookProtectionOptions`
工作簿保護選項。

**屬性：**
```typescript
structure?: boolean              // 是否保護工作簿結構
windows?: boolean               // 是否保護工作簿視窗
password?: string               // 保護密碼
```

**使用範例：**
```typescript
const workbookProtection: WorkbookProtectionOptions = {
  structure: true,
  windows: false,
  password: 'workbook123'
};

workbook.protect('workbook123', workbookProtection);
```

---

## 📈 圖表相關

### `ChartType`
圖表類型。

**類型：**
```typescript
type ChartType = 'column' | 'line' | 'pie' | 'bar' | 'area' | 'scatter' | 'doughnut' | 'radar'
```

### `ChartData`
圖表資料設定。

**屬性：**
```typescript
series: string                  // 系列名稱
categories: string              // 類別範圍，例如 "A2:A10"
values: string                  // 數值範圍，例如 "B2:B10"
color?: string                  // 系列顏色
```

### `ChartOptions`
圖表選項。

**屬性：**
```typescript
title?: string                  // 圖表標題
xAxisTitle?: string            // X 軸標題
yAxisTitle?: string            // Y 軸標題
width?: number                 // 圖表寬度
height?: number                // 圖表高度
showLegend?: boolean           // 是否顯示圖例
showDataLabels?: boolean       // 是否顯示資料標籤
showGridlines?: boolean        // 是否顯示格線
theme?: 'light' | 'dark'      // 主題
```

### `Chart`
圖表介面。

**屬性：**
```typescript
name: string                    // 圖表名稱
type: ChartType                // 圖表類型
data: ChartData[]              // 圖表資料陣列
options: ChartOptions          // 圖表選項
position: {                    // 圖表位置
  row: number;
  col: number;
}
```

**使用範例：**
```typescript
const chartData: ChartData[] = [
  {
    series: 'Sales',
    categories: 'A2:A10',
    values: 'B2:B10',
    color: '#FF0000'
  }
];

const chartOptions: ChartOptions = {
  title: 'Monthly Sales',
  xAxisTitle: 'Month',
  yAxisTitle: 'Sales Amount',
  showLegend: true,
  showGridlines: true
};

const chart: Chart = {
  name: 'Sales Chart',
  type: 'column',
  data: chartData,
  options: chartOptions,
  position: { row: 1, col: 1 }
};

sheet.addChart(chart);
```

---

## ⚡ 效能優化相關

### 記憶體統計
```typescript
interface MemoryStats {
  sheets: number;              // 工作表數量
  totalCells: number;          // 總儲存格數量
  cacheSize: number;           // 快取大小
  cacheHitRate: number;        // 快取命中率
  memoryUsage: number;         // 記憶體使用量
}
```

### 大型資料集選項
```typescript
interface LargeDatasetOptions {
  startRow?: number;           // 起始列
  startCol?: number;           // 起始欄
  chunkSize?: number;          // 分塊大小
}
```

**使用範例：**
```typescript
// 啟用記憶體優化
workbook.setMemoryOptimization(true);
workbook.setChunkSize(1000);
workbook.setCacheEnabled(true);
workbook.setMaxCacheSize(1000000);

// 添加大型資料集
await workbook.addLargeDataset('Sheet1', largeDataArray, {
  startRow: 2,
  startCol: 1,
  chunkSize: 500
});

// 獲取記憶體統計
const stats = workbook.getMemoryStats();
console.log(`工作表數量: ${stats.sheets}`);
console.log(`總儲存格: ${stats.totalCells}`);
console.log(`記憶體使用: ${stats.memoryUsage} bytes`);

// 強制垃圾回收
workbook.forceGarbageCollection();
```

---

## 🚀 完整使用範例

### 基本工作簿操作
```typescript
import { Workbook } from './src/index';

async function createBasicWorkbook() {
  const workbook = new Workbook();
  const sheet = workbook.getWorksheet('Sheet1');
  
  // 設定標題
  sheet.setCell('A1', 'Monthly Savings Report', {
    font: { bold: true, size: 16 },
    alignment: { horizontal: 'center' }
  });
  
  // 設定欄標題
  const headers = ['Month', 'Account', 'Saving Amount (NTD)'];
  headers.forEach((header, index) => {
    sheet.setCell(`A${index + 2}`, header, {
      font: { bold: true },
      fill: { type: 'pattern', patternType: 'solid', fgColor: '#E0E0E0' }
    });
  });
  
  // 設定資料
  const data = [
    ['January', 'Account A', 5000],
    ['January', 'Account B', 3000],
    ['February', 'Account A', 6000],
    ['February', 'Account B', 4000]
  ];
  
  data.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      sheet.setCell(`${String.fromCharCode(65 + colIndex)}${rowIndex + 3}`, value);
    });
  });
  
  // 設定欄寬
  sheet.setColumnWidth('A', 15);
  sheet.setColumnWidth('B', 15);
  sheet.setColumnWidth('C', 20);
  
  // 儲存檔案
  await workbook.writeFile('monthly-savings.xlsx');
}

createBasicWorkbook();
```

### 樞紐分析表示範
```typescript
async function createPivotTableExample() {
  const workbook = new Workbook();
  const sheet = workbook.getWorksheet('Detail');
  
  // 設定資料
  const data = [
    ['Month', 'Account', 'Saving Amount (NTD)'],
    ['January', 'Account A', 5000],
    ['January', 'Account B', 3000],
    ['February', 'Account A', 6000],
    ['February', 'Account B', 4000],
    ['March', 'Account A', 7000],
    ['March', 'Account B', 5000]
  ];
  
  data.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
      sheet.setCell(`${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`, value);
    });
  });
  
  // 創建樞紐分析表
  const pivotConfig: PivotTableConfig = {
    name: 'Savings Summary',
    sourceRange: 'A1:C7',
    targetRange: 'E1:H10',
    fields: [
      {
        name: 'Month',
        sourceColumn: 'Month',
        type: 'row',
        showSubtotal: true
      },
      {
        name: 'Account',
        sourceColumn: 'Account',
        type: 'column',
        showSubtotal: true
      },
      {
        name: 'Saving Amount',
        sourceColumn: 'Saving Amount (NTD)',
        type: 'value',
        function: 'sum',
        numberFormat: '#,##0.00'
      }
    ],
    showGrandTotals: true,
    autoFormat: true
  };
  
  const pivotTable = workbook.createPivotTable(pivotConfig);
  pivotTable.refresh();
  
  // 儲存檔案
  await workbook.writeFile('pivot-example.xlsx');
}

createPivotTableExample();
```

---

## 📚 注意事項

1. **記憶體管理**: 處理大型檔案時，建議啟用記憶體優化功能
2. **公式驗證**: 使用公式前請先驗證語法正確性
3. **樞紐分析表**: 確保資料來源範圍包含標題列
4. **檔案保護**: 設定密碼保護後請妥善保管密碼
5. **效能考量**: 大量資料操作時建議使用分塊處理

---

## 🔗 相關連結

- [專案首頁](../README.md)
- [安裝說明](../INSTALL.md)
- [設定指南](../SETUP.md)
- [專案文件](../PROJECT.md)
