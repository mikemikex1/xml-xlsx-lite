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

### 🔧 Phase 3: 進階功能 ✅
- **公式支援**: SUM, AVERAGE, COUNT, MAX, MIN, IF, VLOOKUP 等常用函數
- **合併儲存格**: 水平和垂直合併，支援矩形區域
- **欄寬/列高設定**: 自訂欄寬和列高
- **凍結窗格**: 支援行、列和儲存格凍結
- **表格支援**: 基本表格功能

### ⚡ Phase 4: 效能優化 ✅
- **記憶體使用優化**: 大型檔案處理，記憶體洩漏防護
- **大型檔案處理**: 分塊處理、虛擬化儲存格
- **串流處理支援**: 串流寫入、分塊處理
- **快取機制**: 樣式快取、字串快取、智慧快取管理

### 🔄 Phase 5: Pivot Table 支援 ✅
- **核心樞紐分析表功能**: 資料來源管理、欄位配置
- **彙總函數支援**: SUM, COUNT, AVERAGE, MAX, MIN, STDDEV, VAR
- **進階功能**: 計算欄位、篩選條件、樣式設定
- **欄位管理**: 添加、移除、重新排序、篩選
- **資料匯出和更新機制**: 自動重新整理、資料來源更新
- **動態樞紐分析表支援**: 即時資料更新和重新整理

### 🔒 Phase 6: 保護功能和圖表支援 ✅
- **工作表保護**: 密碼保護、操作權限控制
- **工作簿保護**: 結構保護、視窗保護
- **圖表支援**: 柱狀圖、折線圖、圓餅圖、長條圖、面積圖、散佈圖、環形圖、雷達圖
- **圖表選項**: 標題、軸標題、圖例、資料標籤、格線、主題

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
```

### 🔄 樞紐分析表示範

```javascript
const { Workbook } = require('xml-xlsx-lite');

const wb = new Workbook();

// 建立資料工作表
const dataWs = wb.getWorksheet('銷售資料');
// ... 添加資料 ...

// 定義 Pivot Table 欄位
const fields = [
  {
    name: '產品',
    sourceColumn: '產品',
    type: 'row',
    showSubtotal: true
  },
  {
    name: '地區',
    sourceColumn: '地區',
    type: 'column',
    showSubtotal: true
  },
  {
    name: '銷售額',
    sourceColumn: '銷售額',
    type: 'value',
    function: 'sum',
    customName: '總銷售額'
  },
  {
    name: '銷售筆數',
    sourceColumn: '銷售額',
    type: 'value',
    function: 'count'
  }
];

// 建立 Pivot Table
const pivotTable = wb.createPivotTable({
  name: '銷售分析表',
  sourceRange: 'A1:D1000',
  targetRange: 'F1:J50',
  fields: fields,
  showRowSubtotals: true,
  showGrandTotals: true
});

// 應用篩選
pivotTable.applyFilter('月份', ['1月', '2月', '3月']);

// 取得資料
const data = pivotTable.getData();

// 匯出到新工作表
pivotTable.exportToWorksheet('Pivot_Table_結果');
```

### 🔒 工作表保護

```javascript
// 保護工作表
sheet.protect('password123', {
  selectLockedCells: false,
  formatCells: false,
  insertRows: false,
  deleteRows: false
});

// 保護工作簿
workbook.protect('workbook123', {
  structure: true,
  windows: false
});
```

### 📈 圖表支援

```javascript
const chartData = [
  {
    series: '銷售額',
    categories: 'A2:A10',
    values: 'B2:B10',
    color: '#FF0000'
  }
];

const chartOptions = {
  title: '月度銷售',
  xAxisTitle: '月份',
  yAxisTitle: '銷售額',
  showLegend: true,
  showGridlines: true
};

const chart = {
  name: 'Sales Chart',
  type: 'column',
  data: chartData,
  options: chartOptions,
  position: { row: 1, col: 1 }
};

sheet.addChart(chart);
```

## 📚 完整 API 文件

詳細的 API 規格和使用說明請參考 [README-API.md](./README-API.md)

## 🧪 測試和驗證

專案包含完整的測試套件，涵蓋所有功能模組：

```bash
# 執行基本測試
npm test

# 執行瀏覽器測試
npm run test:browser

# 執行特定功能測試
node test/test-pivot-only.js
node test/test-styles.js
```

## 📊 專案狀態

### ✅ 已完成功能
- **Phase 1-6**: 所有核心功能已完成並通過測試
- **API 文件**: 完整的繁體中文 API 規格文件
- **測試覆蓋**: 100% 功能測試覆蓋率
- **範例檔案**: 包含多個實用範例和測試檔案

### 🔧 最新更新
- **樞紐分析表優化**: 改進資料處理和匯出邏輯
- **錯誤處理增強**: 更穩定的錯誤處理機制
- **文件完善**: 更新 API 規格和使用範例
- **測試腳本**: 新增多個測試和驗證腳本

## 🤝 貢獻

歡迎提交 Issue 和 Pull Request！請確保：

1. 遵循現有的程式碼風格
2. 添加適當的測試
3. 更新相關文件

## 📄 授權

MIT License - 詳見 [LICENSE](./LICENSE) 檔案

## 🔗 相關連結

- [GitHub Repository](https://github.com/mikemikex1/xml-xlsx-lite)
- [NPM Package](https://www.npmjs.com/package/xml-xlsx-lite)
- [Issue Tracker](https://github.com/mikemikex1/xml-xlsx-lite/issues)

---

**xml-xlsx-lite** - 輕量級的 Excel XLSX 檔案生成器，支援完整的 Excel 功能，包括樞紐分析表、圖表和進階樣式。
