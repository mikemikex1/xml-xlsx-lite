# 🚀 xml-xlsx-lite

> **Lightweight Excel XLSX generator with full Excel features: dynamic pivot tables, charts, styles, and Chinese support. Fast, TypeScript-friendly Excel file creation library.**

> **輕量級 Excel XLSX 生成器，支援樞紐分析表、圖表、樣式，完整繁體中文支援。**

[![npm version](https://img.shields.io/npm/v/xml-xlsx-lite.svg)](https://www.npmjs.com/package/xml-xlsx-lite)
[![npm downloads](https://img.shields.io/npm/dm/xml-xlsx-lite.svg)](https://www.npmjs.com/package/xml-xlsx-lite)
[![License](https://img.shields.io/npm/l/xml-xlsx-lite.svg)](https://github.com/mikemikex1/xml-xlsx-lite/blob/main/LICENSE)
[![TypeScript](https://img.shields.io/badge/TypeScript-5.0+-blue.svg)](https://www.typescriptlang.org/)

## 📋 Table of Contents

- [🚀 Features](#-features)
- [📦 Installation](#-installation)
- [🎯 Quick Start](#-quick-start)
- [📚 Complete Guide](#-complete-guide)
  - [1. 創建 Excel 檔案](#1-創建-excel-檔案)
  - [2. 基本儲存格操作](#2-基本儲存格操作)
  - [3. 樣式和格式化](#3-樣式和格式化)
  - [4. 工作表管理](#4-工作表管理)
  - [5. 讀取 Excel 檔案](#5-讀取-excel-檔案)
  - [6. 樞紐分析表](#6-樞紐分析表)
  - [7. 圖表支援](#7-圖表支援)
  - [8. 進階功能](#8-進階功能)
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

// 創建工作簿
const workbook = new Workbook();

// 取得工作表
const worksheet = workbook.getWorksheet('Sheet1');

// 設定儲存格
worksheet.setCell('A1', 'Hello World');
worksheet.setCell('B1', 42);

// 儲存檔案
const buffer = await workbook.writeBuffer();
```

## 📚 Complete Guide

### 1. 創建 Excel 檔案

#### 1.1 基本工作簿創建

```typescript
import { Workbook } from 'xml-xlsx-lite';

// 創建新的工作簿
const workbook = new Workbook();

// 取得預設工作表
const worksheet = workbook.getWorksheet('Sheet1');

// 設定標題
worksheet.setCell('A1', '產品銷售報表');
worksheet.setCell('A2', '2024年度');

// 設定欄位標題
worksheet.setCell('A3', '產品名稱');
worksheet.setCell('B3', '銷售數量');
worksheet.setCell('C3', '單價');
worksheet.setCell('D3', '總金額');

// 設定資料
worksheet.setCell('A4', '筆記型電腦');
worksheet.setCell('B4', 10);
worksheet.setCell('C4', 35000);
worksheet.setCell('D4', 350000);

worksheet.setCell('A5', '滑鼠');
worksheet.setCell('B5', 50);
worksheet.setCell('C5', 500);
worksheet.setCell('D5', 25000);

// 儲存檔案
const buffer = await workbook.writeBuffer();
```

#### 1.2 多工作表工作簿

```typescript
const workbook = new Workbook();

// 創建多個工作表
const summarySheet = workbook.getWorksheet('總覽');
const detailSheet = workbook.getWorksheet('詳細資料');
const chartSheet = workbook.getWorksheet('圖表');

// 在不同工作表中設定資料
summarySheet.setCell('A1', '銷售總覽');
detailSheet.setCell('A1', '詳細銷售資料');
chartSheet.setCell('A1', '銷售圖表');
```

### 2. 基本儲存格操作

#### 2.1 儲存格值設定

```typescript
const worksheet = workbook.getWorksheet('Sheet1');

// 不同類型的資料
worksheet.setCell('A1', '文字');                    // 字串
worksheet.setCell('B1', 123);                       // 數字
worksheet.setCell('C1', true);                      // 布林值
worksheet.setCell('D1', new Date());                // 日期
worksheet.setCell('E1', null);                      // 空值
worksheet.setCell('F1', '');                        // 空字串

// 使用座標
worksheet.setCell('G1', '使用 A1 座標');
worksheet.setCell(1, 8, '使用行列座標');            // 第1行，第8列
```

#### 2.2 儲存格範圍操作

```typescript
// 設定範圍內的儲存格
for (let row = 1; row <= 10; row++) {
    for (let col = 1; col <= 5; col++) {
        const value = `R${row}C${col}`;
        worksheet.setCell(row, col, value);
    }
}

// 設定整行
for (let col = 1; col <= 5; col++) {
    worksheet.setCell(1, col, `標題${col}`);
}

// 設定整列
for (let row = 1; row <= 10; row++) {
    worksheet.setCell(row, 1, `項目${row}`);
}
```

### 3. 樣式和格式化

#### 3.1 基本樣式

```typescript
// 字體樣式
worksheet.setCell('A1', '粗體標題', {
    font: {
        bold: true,
        size: 16,
        color: 'FF0000'  // 紅色
    }
});

// 對齊樣式
worksheet.setCell('B1', '置中對齊', {
    alignment: {
        horizontal: 'center',
        vertical: 'middle'
    }
});

// 邊框樣式
worksheet.setCell('C1', '有邊框', {
    border: {
        top: { style: 'thin', color: '000000' },
        bottom: { style: 'double', color: '000000' },
        left: { style: 'thin', color: '000000' },
        right: { style: 'thin', color: '000000' }
    }
});

// 填滿樣式
worksheet.setCell('D1', '有背景色', {
    fill: {
        type: 'solid',
        color: 'FFFF00'  // 黃色
    }
});
```

#### 3.2 數字格式

```typescript
// 貨幣格式
worksheet.setCell('A1', 1234.56, {
    numFmt: '¥#,##0.00'
});

// 百分比格式
worksheet.setCell('B1', 0.1234, {
    numFmt: '0.00%'
});

// 日期格式
worksheet.setCell('C1', new Date(), {
    numFmt: 'yyyy-mm-dd'
});

// 自訂格式
worksheet.setCell('D1', 42, {
    numFmt: '0 "件"'
});
```

#### 3.3 合併儲存格

```typescript
// 合併儲存格
worksheet.mergeCells('A1:D1');
worksheet.setCell('A1', '合併的標題');

// 合併多行
worksheet.mergeCells('A2:A5');
worksheet.setCell('A2', '垂直合併');
```

### 4. 工作表管理

#### 4.1 欄寬和列高

```typescript
// 設定欄寬
worksheet.setColumnWidth('A', 20);      // 欄 A 寬度 20
worksheet.setColumnWidth(2, 15);        // 欄 B 寬度 15

// 設定列高
worksheet.setRowHeight(1, 30);          // 第1列高度 30
worksheet.setRowHeight(2, 25);          // 第2列高度 25
```

#### 4.2 凍結窗格

```typescript
// 凍結第一行和第一列
worksheet.freezePanes(2, 2);

// 只凍結第一行
worksheet.freezePanes(2);

// 只凍結第一列
worksheet.freezePanes(undefined, 2);

// 取消凍結
worksheet.unfreezePanes();
```

#### 4.3 工作表保護

```typescript
// 保護工作表
worksheet.protect('password123', {
    selectLockedCells: false,
    selectUnlockedCells: true,
    formatCells: false,
    formatColumns: false,
    formatRows: false
});

// 檢查保護狀態
const isProtected = worksheet.isProtected();
```

### 5. 讀取 Excel 檔案

#### 5.1 基本讀取

```typescript
import { Workbook } from 'xml-xlsx-lite';

// 從檔案讀取
const workbook = await Workbook.readFile('existing-file.xlsx');

// 從 Buffer 讀取
const fs = require('fs');
const buffer = fs.readFileSync('existing-file.xlsx');
const workbook = await Workbook.readBuffer(buffer);
```

#### 5.2 讀取工作表資料

```typescript
// 取得工作表
const worksheet = workbook.getWorksheet('Sheet1');

// 轉換為二維陣列
const arrayData = worksheet.toArray();
console.log('陣列資料:', arrayData);

// 轉換為 JSON 物件陣列
const jsonData = worksheet.toJSON({ headerRow: 1 });
console.log('JSON 資料:', jsonData);

// 取得特定範圍
const rangeData = worksheet.getRange('A1:D10');
console.log('範圍資料:', rangeData);
```

#### 5.3 讀取選項

```typescript
const workbook = await Workbook.readFile('file.xlsx', {
    enableSharedStrings: true,      // 啟用共享字串優化
    preserveStyles: true,           // 保留樣式資訊
    preserveFormulas: true,         // 保留公式
    preservePivotTables: true,      // 保留樞紐分析表
    preserveCharts: true            // 保留圖表
});
```

### 6. 樞紐分析表

#### 6.1 手動創建樞紐分析表

```typescript
// 創建手動樞紐分析表
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

#### 6.2 動態樞紐分析表

```typescript
// 創建基礎工作簿
const workbook = new Workbook();
const dataSheet = workbook.getWorksheet('Data');

// 填入資料
const data = [
    ['部門', '月份', '銷售額'],
    ['IT', '1月', 1000],
    ['IT', '2月', 1200],
    ['HR', '1月', 800],
    ['HR', '2月', 900]
];

data.forEach((row, rowIndex) => {
    row.forEach((value, colIndex) => {
        const address = String.fromCharCode(65 + colIndex) + (rowIndex + 1);
        dataSheet.setCell(address, value);
    });
});

// 儲存基礎檔案
const baseBuffer = await workbook.writeBuffer();

// 動態插入樞紐分析表
import { addPivotToWorkbookBuffer } from 'xml-xlsx-lite';

const enhancedBuffer = await addPivotToWorkbookBuffer(baseBuffer, {
    sourceSheet: 'Data',
    sourceRange: 'A1:C100',
    targetSheet: 'Pivot',
    anchorCell: 'A3',
    layout: {
        rows: [{ name: '部門' }],
        cols: [{ name: '月份' }],
        values: [{ 
            name: '銷售額', 
            agg: 'sum', 
            displayName: '總銷售額' 
        }]
    },
    refreshOnLoad: true,
    styleName: 'PivotStyleMedium9'
});
```

#### 6.3 樞紐分析表配置選項

```typescript
const pivotOptions = {
    sourceSheet: 'Data',           // 來源工作表
    sourceRange: 'A1:C100',        // 來源範圍
    targetSheet: 'Pivot',          // 目標工作表
    anchorCell: 'A3',              // 錨點儲存格
    
    layout: {
        rows: [                     // 行欄位
            { name: '部門' },
            { name: '產品' }        // 多層級行欄位
        ],
        cols: [                     // 列欄位
            { name: '月份' },
            { name: '年份' }
        ],
        values: [                   // 值欄位
            { 
                name: '銷售額', 
                agg: 'sum',         // 彙總方式：sum, avg, count, max, min
                displayName: '總銷售額',
                numberFormat: '#,##0'
            },
            { 
                name: '數量', 
                agg: 'count',
                displayName: '訂單數'
            }
        ]
    },
    
    refreshOnLoad: true,            // 開啟時自動重新整理
    styleName: 'PivotStyleMedium9', // 樞紐分析表樣式
    showGrandTotals: true,          // 顯示總計
    showSubTotals: true,            // 顯示小計
    enableDrilldown: true           // 啟用向下鑽研
};
```

### 7. 圖表支援

#### 7.1 基本圖表

```typescript
// 創建圖表工作表
const chartSheet = workbook.getWorksheet('圖表');

// 設定圖表資料
chartSheet.setCell('A1', '月份');
chartSheet.setCell('B1', '銷售額');
chartSheet.setCell('A2', '1月');
chartSheet.setCell('B2', 1000);
chartSheet.setCell('A3', '2月');
chartSheet.setCell('B3', 1200);
chartSheet.setCell('A4', '3月');
chartSheet.setCell('B4', 1100);

// 添加圖表（基本支援）
chartSheet.addChart({
    type: 'bar',
    title: '月度銷售圖表',
    dataRange: 'A1:B4',
    position: { x: 100, y: 100, width: 400, height: 300 }
});
```

#### 7.2 圖表類型

```typescript
// 支援的圖表類型
const chartTypes = [
    'bar',          // 長條圖
    'line',         // 折線圖
    'pie',          // 圓餅圖
    'column',       // 直條圖
    'area',         // 區域圖
    'scatter'       // 散佈圖
];

chartTypes.forEach((type, index) => {
    const row = index + 1;
    chartSheet.setCell(`A${row}`, `${type} 圖表`);
    chartSheet.addChart({
        type: type,
        title: `${type} 圖表示例`,
        dataRange: 'A1:B4',
        position: { x: 100, y: 100 + index * 100, width: 300, height: 200 }
    });
});
```

### 8. 進階功能

#### 8.1 公式支援

```typescript
// 設定公式
worksheet.setFormula('D4', '=B4*C4');           // 乘法
worksheet.setFormula('D5', '=B5*C5');           // 乘法
worksheet.setFormula('D6', '=SUM(D4:D5)');     // 總和
worksheet.setFormula('B6', '=SUM(B4:B5)');     // 數量總和
worksheet.setFormula('C6', '=AVERAGE(C4:C5)'); // 平均單價

// 邏輯公式
worksheet.setFormula('E4', '=IF(D4>100000,"高","低")');
worksheet.setFormula('F4', '=AND(B4>5,C4>10000)');
```

#### 8.2 條件格式

```typescript
// 設定條件格式（基本支援）
worksheet.setCell('A1', '條件格式測試', {
    font: { bold: true },
    fill: { type: 'solid', color: 'FFFF00' }
});

// 根據值設定樣式
const salesData = [1000, 1200, 800, 900, 1500];
salesData.forEach((value, index) => {
    const row = index + 1;
    const cell = worksheet.setCell(`B${row}`, value);
    
    // 根據銷售額設定顏色
    if (value > 1200) {
        cell.style = { fill: { type: 'solid', color: '00FF00' } }; // 綠色
    } else if (value > 1000) {
        cell.style = { fill: { type: 'solid', color: 'FFFF00' } }; // 黃色
    } else {
        cell.style = { fill: { type: 'solid', color: 'FF0000' } }; // 紅色
    }
});
```

#### 8.3 效能優化

```typescript
// 大量資料處理
const largeData = [];
for (let i = 0; i < 10000; i++) {
    largeData.push({
        id: i + 1,
        name: `項目${i + 1}`,
        value: Math.random() * 1000
    });
}

// 批次處理
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

| 方法 | 描述 | 狀態 |
|------|------|------|
| `new Workbook()` | 創建新工作簿 | ✅ Stable |
| `getWorksheet(name)` | 取得工作表 | ✅ Stable |
| `writeBuffer()` | 輸出為 Buffer | ✅ Stable |
| `writeFile(path)` | 直接儲存檔案 | ✅ Stable |
| `writeFileWithPivotTables(path, options)` | 儲存含樞紐分析表的檔案 | ✅ Stable |
| `createManualPivotTable(data, options)` | 創建手動樞紐分析表 | ✅ Stable |

### Worksheet

| 方法 | 描述 | 狀態 |
|------|------|------|
| `setCell(address, value, options)` | 設定儲存格 | ✅ Stable |
| `getCell(address)` | 取得儲存格 | ✅ Stable |
| `mergeCells(range)` | 合併儲存格 | ✅ Stable |
| `setColumnWidth(col, width)` | 設定欄寬 | ✅ Stable |
| `setRowHeight(row, height)` | 設定列高 | ✅ Stable |
| `freezePanes(row?, col?)` | 凍結窗格 | ✅ Stable |
| `protect(password, options)` | 保護工作表 | ✅ Stable |
| `addChart(chart)` | 添加圖表 | 🔶 Experimental |

### Reading

| 方法 | 描述 | 狀態 |
|------|------|------|
| `Workbook.readFile(path, options)` | 從檔案讀取 | ✅ Stable |
| `Workbook.readBuffer(buffer, options)` | 從 Buffer 讀取 | ✅ Stable |
| `worksheet.toArray()` | 轉換為陣列 | ✅ Stable |
| `worksheet.toJSON(options)` | 轉換為 JSON | ✅ Stable |

## 📖 Examples

### 完整範例：銷售報表系統

```typescript
import { Workbook } from 'xml-xlsx-lite';

async function createSalesReport() {
    const workbook = new Workbook();
    
    // 1. 創建資料工作表
    const dataSheet = workbook.getWorksheet('銷售資料');
    
    // 設定標題
    dataSheet.setCell('A1', '日期', { font: { bold: true } });
    dataSheet.setCell('B1', '產品', { font: { bold: true } });
    dataSheet.setCell('C1', '數量', { font: { bold: true } });
    dataSheet.setCell('D1', '單價', { font: { bold: true } });
    dataSheet.setCell('E1', '總額', { font: { bold: true } });
    
    // 填入資料
    const salesData = [
        ['2024-01-01', '筆記型電腦', 2, 35000, 70000],
        ['2024-01-01', '滑鼠', 10, 500, 5000],
        ['2024-01-02', '鍵盤', 5, 800, 4000],
        ['2024-01-02', '螢幕', 3, 8000, 24000],
        ['2024-01-03', '耳機', 8, 1200, 9600]
    ];
    
    salesData.forEach((row, index) => {
        const rowNum = index + 2;
        row.forEach((value, colIndex) => {
            const col = String.fromCharCode(65 + colIndex);
            dataSheet.setCell(`${col}${rowNum}`, value);
        });
        
        // 設定公式
        const rowNum2 = index + 2;
        dataSheet.setFormula(`E${rowNum2}`, `=C${rowNum2}*D${rowNum2}`);
    });
    
    // 2. 創建樞紐分析表
    const pivotSheet = workbook.getWorksheet('樞紐分析');
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
    
    // 3. 創建圖表
    const chartSheet = workbook.getWorksheet('圖表');
    chartSheet.setCell('A1', '產品銷售圖表', { font: { bold: true, size: 16 } });
    
    // 4. 儲存檔案
    await workbook.writeFileWithPivotTables('銷售報表.xlsx');
    
    console.log('銷售報表已創建完成！');
}

createSalesReport();
```

## 🚨 Important Notes

### ⚠️ 重要提醒

- **不要使用 `writeFile()` 方法**：此方法尚未完全實現，請使用 `writeBuffer()` + `fs.writeFileSync()` 或新的 `writeFileWithPivotTables()` 方法
- **樞紐分析表限制**：動態樞紐分析表需要先在 Excel 中手動重新整理一次
- **瀏覽器相容性**：某些功能（如檔案讀取）僅支援 Node.js 環境

### 🔧 正確的檔案儲存方式

```typescript
// ❌ 錯誤方式
await workbook.writeFile('file.xlsx');

// ✅ 正確方式 1：使用 Buffer
const buffer = await workbook.writeBuffer();
const fs = require('fs');
fs.writeFileSync('file.xlsx', new Uint8Array(buffer));

// ✅ 正確方式 2：使用新的便捷方法
await workbook.writeFileWithPivotTables('file.xlsx', pivotOptions);
```

## 📊 Feature Matrix

| 功能 | 狀態 | 說明 | 替代方案 |
|------|------|------|----------|
| **基本功能** |
| 創建工作簿 | ✅ Stable | 完全支援 | - |
| 儲存格操作 | ✅ Stable | 完全支援 | - |
| 樣式設定 | ✅ Stable | 完全支援 | - |
| 公式支援 | ✅ Stable | 基本公式 | - |
| **進階功能** |
| 樞紐分析表 | 🔶 Experimental | 動態插入 | 手動創建 |
| 圖表支援 | 🔶 Experimental | 基本支援 | 手動創建 |
| 檔案讀取 | ✅ Stable | 完全支援 | - |
| **效能優化** |
| 大量資料 | ✅ Stable | 批次處理 | 串流處理 |
| 記憶體優化 | ✅ Stable | 自動優化 | 手動控制 |

## 🌐 Browser Support

- ✅ **Node.js**: 完全支援
- 🔶 **現代瀏覽器**: 基本功能支援（部分功能受限）
- ❌ **舊版瀏覽器**: 不支援

### 瀏覽器使用範例

```html
<!DOCTYPE html>
<html>
<head>
    <title>xml-xlsx-lite 瀏覽器測試</title>
</head>
<body>
    <h1>Excel 生成測試</h1>
    <button onclick="generateExcel()">生成 Excel</button>
    
    <script type="module">
        import { Workbook } from './node_modules/xml-xlsx-lite/dist/index.esm.js';
        
        async function generateExcel() {
            const workbook = new Workbook();
            const worksheet = workbook.getWorksheet('Sheet1');
            
            worksheet.setCell('A1', 'Hello from Browser!');
            worksheet.setCell('B1', new Date());
            
            const buffer = await workbook.writeBuffer();
            
            // 下載檔案
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

歡迎貢獻！請查看我們的 [貢獻指南](CONTRIBUTING.md)。

### 開發環境設置

```bash
git clone https://github.com/mikemikex1/xml-xlsx-lite.git
cd xml-xlsx-lite
npm install
npm run dev
```

### 測試

```bash
npm run test:all        # 運行所有測試
npm run verify          # 驗證功能
npm run build           # 構建專案
```

## 📄 License

MIT License - 詳見 [LICENSE](LICENSE) 檔案。

---

## 🌟 特色功能展示

### 🚀 快速開始

```bash
# 安裝
npm install xml-xlsx-lite

# 基本使用
node -e "
const { Workbook } = require('xml-xlsx-lite');
const wb = new Workbook();
const ws = wb.getWorksheet('Sheet1');
ws.setCell('A1', 'Hello Excel!');
wb.writeBuffer().then(buf => require('fs').writeFileSync('test.xlsx', new Uint8Array(buf)));
"
```

### 📊 樞紐分析表示例

```typescript
// 創建包含樞紐分析表的完整報表
const workbook = new Workbook();
const dataSheet = workbook.getWorksheet('資料');

// 填入銷售資料
const salesData = [
    ['部門', '月份', '產品', '數量', '金額'],
    ['IT', '1月', '筆電', 5, 175000],
    ['IT', '2月', '筆電', 3, 105000],
    ['HR', '1月', '辦公用品', 20, 4000],
    ['HR', '2月', '辦公用品', 15, 3000]
];

salesData.forEach((row, i) => {
    row.forEach((value, j) => {
        const address = String.fromCharCode(65 + j) + (i + 1);
        dataSheet.setCell(address, value);
    });
});

// 創建手動樞紐分析表
workbook.createManualPivotTable(salesData.slice(1).map(row => ({
    部門: row[0],
    月份: row[1],
    產品: row[2],
    數量: row[3],
    金額: row[4]
})), {
    rowField: '部門',
    columnField: '月份',
    valueField: '金額',
    aggregation: 'sum'
});

// 儲存檔案
await workbook.writeFileWithPivotTables('銷售樞紐報表.xlsx');
```

---

**🎯 目標**: 提供最完整、最易用的 Excel 生成解決方案！

**💡 特色**: 從基本操作到進階功能，從零開始的完整指南！

**🚀 願景**: 讓每個開發者都能輕鬆創建專業的 Excel 報表！
