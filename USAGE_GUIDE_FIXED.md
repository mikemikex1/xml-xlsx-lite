# 🔧 xml-xlsx-lite 使用指南 - 問題解決版

## 📅 更新時間
**2024年12月21日**

## 🚨 已知問題與解決方案

### ❌ 問題 1: TypeScript 型別與匯入方式
**問題描述**: 套件的 TypeScript 型別與匯入方式不直觀，需用 require 並加上 @ts-ignore

**解決方案**:
```typescript
// ❌ 錯誤方式
// @ts-ignore
const { Workbook } = require('xml-xlsx-lite');

// ✅ 正確方式
import { Workbook } from 'xml-xlsx-lite';
// 或者
const { Workbook } = require('xml-xlsx-lite'); // 不需要 @ts-ignore
```

### ❌ 問題 2: writeFile 方法未實作
**問題描述**: 官方文件與 API 實作有落差，writeFile 其實未實作

**解決方案**:
```typescript
// ❌ 錯誤方式
await workbook.writeFile('output.xlsx'); // 會拋出錯誤

// ✅ 正確方式
const buffer = await workbook.writeBuffer();
fs.writeFileSync('output.xlsx', new Uint8Array(buffer));
```

### ❌ 問題 3: 樞紐分析表 API 問題
**問題描述**: 樞紐分析表 API 雖然有設計，但型別嚴格且功能有缺陷

**解決方案**: 使用手動創建樞紐分析表結果的方式

### ❌ 問題 4: 文件與實作不符
**問題描述**: 文件主要為中文，且範例多為 JavaScript，TypeScript 用戶需自行摸索

**解決方案**: 提供完整的 TypeScript 範例和型別定義

---

## ✅ 修正後的使用範例

### 🎯 JavaScript 版本

```javascript
const { Workbook } = require('xml-xlsx-lite');
const fs = require('fs');

async function main() {
  // 建立工作簿
  const wb = new Workbook();
  
  // 建立數據表
  const ws = wb.getWorksheet('數據');
  
  const data = [
    ['部門', '姓名', '月份', '銷售額'],
    ['A', '小明', '1月', 100],
    ['A', '小明', '2月', 120],
    ['A', '小華', '1月', 90],
    ['B', '小美', '1月', 200],
    ['B', '小美', '2月', 180],
    ['B', '小強', '1月', 150],
  ];
  
  // 寫入數據表
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const cellAddress = String.fromCharCode(65 + c) + (r + 1);
      const cellValue = data[r][c];
      
      // 為標題行添加樣式
      if (r === 0) {
        ws.setCell(cellAddress, cellValue, { 
          font: { bold: true },
          fill: { type: 'pattern', color: 'E0E0E0' }
        });
      } else {
        // 為數值欄位添加格式
        if (c === 3) { // 銷售額欄位
          ws.setCell(cellAddress, cellValue, { 
            numFmt: '#,##0',
            alignment: { horizontal: 'right' }
          });
        } else {
          ws.setCell(cellAddress, cellValue);
        }
      }
    }
  }
  
  // 設定欄寬
  ws.setColumnWidth('A', 12); // 部門
  ws.setColumnWidth('B', 12); // 姓名
  ws.setColumnWidth('C', 10); // 月份
  ws.setColumnWidth('D', 15); // 銷售額
  
  // 使用 writeBuffer 方法輸出 Excel 檔案
  const buffer = await wb.writeBuffer();
  fs.writeFileSync('output.xlsx', new Uint8Array(buffer));
  console.log('Excel 檔案 output.xlsx 已產生');
}

main();
```

### 🎯 TypeScript 版本

```typescript
import { Workbook } from 'xml-xlsx-lite';
import * as fs from 'fs';

interface SalesData {
  department: string;
  name: string;
  month: string;
  amount: number;
}

interface PivotResult {
  department: string;
  name: string;
  month1: number;
  month2: number;
  total: number;
}

async function main(): Promise<void> {
  // 建立工作簿
  const wb = new Workbook();
  
  // 建立數據表
  const ws = wb.getWorksheet('數據');
  
  // 測試數據 - 使用強型別
  const data: (string | number)[][] = [
    ['部門', '姓名', '月份', '銷售額'],
    ['A', '小明', '1月', 100],
    ['A', '小明', '2月', 120],
    ['A', '小華', '1月', 90],
    ['B', '小美', '1月', 200],
    ['B', '小美', '2月', 180],
    ['B', '小強', '1月', 150],
  ];
  
  // 寫入數據表 - 使用更安全的方式
  for (let r = 0; r < data.length; r++) {
    for (let c = 0; c < data[r].length; c++) {
      const cellAddress = String.fromCharCode(65 + c) + (r + 1);
      const cellValue = data[r][c];
      
      // 為標題行添加樣式
      if (r === 0) {
        ws.setCell(cellAddress, cellValue, { 
          font: { bold: true },
          fill: { type: 'pattern', color: 'E0E0E0' }
        });
      } else {
        // 為數值欄位添加格式
        if (c === 3) { // 銷售額欄位
          ws.setCell(cellAddress, cellValue, { 
            numFmt: '#,##0',
            alignment: { horizontal: 'right' }
          });
        } else {
          ws.setCell(cellAddress, cellValue);
        }
      }
    }
  }
  
  // 設定欄寬
  ws.setColumnWidth('A', 12);
  ws.setColumnWidth('B', 12);
  ws.setColumnWidth('C', 10);
  ws.setColumnWidth('D', 15);
  
  // 使用 writeBuffer 方法輸出 Excel 檔案
  const buffer = await wb.writeBuffer();
  fs.writeFileSync('output.xlsx', new Uint8Array(buffer));
  console.log('Excel 檔案 output.xlsx 已產生');
}

main();
```

---

## 🔧 樞紐分析表解決方案

### ❌ 避免使用自動樞紐分析表

```typescript
// ❌ 不要使用這個（有問題）
const pivotTable = workbook.createPivotTable(pivotConfig);
const resultSheet = pivotTable.exportToWorksheet('工作表5');
```

### ✅ 使用手動創建樞紐分析表結果

```typescript
// ✅ 推薦使用這個方式
const pivotSheet = workbook.getWorksheet('樞紐分析表');

// 設定標題
pivotSheet.setCell('A1', '銷售額樞紐分析表', {
  font: { bold: true, size: 16 },
  alignment: { horizontal: 'center' }
});

// 設定欄標題
pivotSheet.setCell('A3', '部門', { font: { bold: true } });
pivotSheet.setCell('B3', '姓名', { font: { bold: true } });
pivotSheet.setCell('C3', '1月', { font: { bold: true } });
pivotSheet.setCell('D3', '2月', { font: { bold: true } });
pivotSheet.setCell('E3', '總計', { font: { bold: true } });

// 手動計算並填入結果
const pivotData = [
  ['A', '小明', 100, 120, 220],
  ['A', '小華', 90, 0, 90],
  ['B', '小美', 200, 180, 380],
  ['B', '小強', 150, 0, 150]
];

pivotData.forEach((row, index) => {
  const rowNum = index + 4;
  pivotSheet.setCell(`A${rowNum}`, row[0]);
  pivotSheet.setCell(`B${rowNum}`, row[1]);
  pivotSheet.setCell(`C${rowNum}`, row[2], { 
    numFmt: '#,##0',
    alignment: { horizontal: 'right' }
  });
  pivotSheet.setCell(`D${rowNum}`, row[3], { 
    numFmt: '#,##0',
    alignment: { horizontal: 'right' }
  });
  pivotSheet.setCell(`E${rowNum}`, row[4], { 
    numFmt: '#,##0',
    font: { bold: true },
    alignment: { horizontal: 'right' }
  });
});
```

---

## 📋 常用功能範例

### 🎨 儲存格樣式設定

```typescript
// 字體樣式
ws.setCell('A1', '標題', {
  font: { 
    bold: true, 
    size: 16, 
    color: 'FF0000' 
  }
});

// 對齊方式
ws.setCell('B1', '置中', {
  alignment: { 
    horizontal: 'center', 
    vertical: 'middle' 
  }
});

// 填滿顏色
ws.setCell('C1', '背景色', {
  fill: { 
    type: 'pattern', 
    color: 'E0E0E0' 
  }
});

// 邊框樣式
ws.setCell('D1', '邊框', {
  border: {
    top: { style: 'thick', color: '000000' },
    bottom: { style: 'thick', color: '000000' }
  }
});

// 數字格式
ws.setCell('E1', 1234.56, {
  numFmt: '#,##0.00'
});
```

### 📏 欄寬和列高設定

```typescript
// 設定欄寬
ws.setColumnWidth('A', 15);
ws.setColumnWidth('B', 20);
ws.setColumnWidth('C', 12);

// 設定列高
ws.setRowHeight(1, 30);
ws.setRowHeight(2, 25);
```

### 🔒 工作表保護

```typescript
// 保護工作表
ws.protect({
  password: 'password123',
  selectLockedCells: false,
  selectUnlockedCells: true,
  formatCells: false,
  formatColumns: false,
  formatRows: false,
  insertColumns: false,
  insertRows: false,
  insertHyperlinks: false,
  deleteColumns: false,
  deleteRows: false,
  sort: false,
  autoFilter: false,
  pivotTables: false
});
```

---

## 🚀 最佳實踐

### ✅ 推薦做法

1. **使用 writeBuffer 方法**: 避免使用未實作的 writeFile
2. **手動創建樞紐分析表**: 避免自動樞紐分析表的問題
3. **強型別定義**: 為複雜資料結構定義介面
4. **錯誤處理**: 添加適當的錯誤處理機制
5. **樣式設定**: 使用樣式提升 Excel 檔案品質

### ❌ 避免做法

1. **使用 @ts-ignore**: 會隱藏型別錯誤
2. **依賴自動樞紐分析表**: 目前功能有缺陷
3. **直接使用 writeFile**: 會拋出錯誤
4. **忽略型別檢查**: 會導致執行時錯誤

---

## 🔍 常見錯誤與解決方案

### ❌ 錯誤 1: writeFile method needs to be implemented externally

**錯誤訊息**: `Error: writeFile method needs to be implemented externally. Use writeBuffer() and save manually.`

**解決方案**: 使用 writeBuffer 方法
```typescript
// 錯誤方式
await workbook.writeFile('output.xlsx');

// 正確方式
const buffer = await workbook.writeBuffer();
fs.writeFileSync('output.xlsx', new Uint8Array(buffer));
```

### ❌ 錯誤 2: TypeScript 型別錯誤

**錯誤訊息**: `Property 'setCell' does not exist on type 'Worksheet'`

**解決方案**: 檢查匯入方式
```typescript
// 錯誤方式
import { Workbook } from 'xml-xlsx-lite/dist/index.js';

// 正確方式
import { Workbook } from 'xml-xlsx-lite';
```

### ❌ 錯誤 3: 樞紐分析表資料異常

**錯誤訊息**: 樞紐分析表顯示不正確的資料

**解決方案**: 使用手動創建方式
```typescript
// 不要使用自動樞紐分析表
// const pivotTable = workbook.createPivotTable(config);

// 使用手動創建
const pivotSheet = workbook.getWorksheet('樞紐分析表');
// 手動填入資料...
```

---

## 📚 相關資源

### 🔗 官方資源
- **NPM 套件**: https://www.npmjs.com/package/xml-xlsx-lite
- **GitHub 倉庫**: https://github.com/mikemikex1/xml-xlsx-lite
- **API 文件**: [README-API.md](./README-API.md)

### 📖 測試檔案
- **JavaScript 範例**: `test/fixed-usage-example.js`
- **TypeScript 範例**: `test/fixed-usage-example.ts`
- **樞紐分析表測試**: `test/test-simple-pivot-result.js`

---

## 🎯 總結

通過使用修正後的使用方式，您可以：

1. **✅ 完全避免 TypeScript 型別問題**
2. **✅ 正確使用 writeBuffer 方法**
3. **✅ 創建準確的樞紐分析表結果**
4. **✅ 享受完整的樣式和格式功能**
5. **✅ 生成高品質的 Excel 檔案**

**xml-xlsx-lite** 雖然存在一些 API 實作問題，但通過正確的使用方式，仍然可以創建功能完整的 Excel 檔案。我們將持續改進，為使用者提供更好的體驗。
