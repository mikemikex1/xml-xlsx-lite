# 🚀 動態樞紐分析表使用指南

## 📖 概述

xml-xlsx-lite 現在支援**動態樞紐分析表**功能！這意味著你可以在既有的 Excel 檔案上動態插入可刷新的樞紐分析表，就像在 Excel 中使用「插入→樞紐分析表」一樣。

## ✨ 主要特性

- **動態插入**：在既有 Excel 檔案上插入樞紐分析表
- **原生支援**：生成標準的 Excel 樞紐分析表 XML 結構
- **可刷新**：支援資料更新後的重新整理
- **樣式自訂**：可選擇不同的樞紐分析表樣式
- **欄位配置**：靈活配置行、列、值欄位

## 🚀 快速開始

### 安裝

```bash
npm install xml-xlsx-lite
```

### 基本使用

```javascript
const { addPivotToWorkbookBuffer } = require('xml-xlsx-lite');
const fs = require('fs');

async function createDynamicPivot() {
  // 1. 讀取既有 Excel 檔案
  const baseBuffer = fs.readFileSync('base-workbook.xlsx');
  
  // 2. 配置樞紐分析表
  const pivotOptions = {
    sourceSheet: "數據",           // 來源工作表名稱
    sourceRange: "A1:D100",       // 資料範圍（含標題列）
    targetSheet: "Pivot",         // 目標工作表名稱
    anchorCell: "A3",             // 樞紐分析表錨點位置
    layout: {
      rows: [{ name: "部門" }],   // 行欄位
      cols: [{ name: "月份" }],   // 列欄位
      values: [{                  // 值欄位
        name: "銷售額",           // 欄位名稱
        agg: "sum",               // 聚合方式
        displayName: "銷售額合計" // 顯示名稱
      }]
    },
    refreshOnLoad: true,          // 開啟時自動重新整理
    styleName: "PivotStyleMedium9" // 樣式名稱
  };
  
  // 3. 動態插入樞紐分析表
  const enhancedBuffer = await addPivotToWorkbookBuffer(baseBuffer, pivotOptions);
  
  // 4. 儲存結果
  fs.writeFileSync('pivot-workbook.xlsx', enhancedBuffer);
  console.log('樞紐分析表插入完成！');
}

createDynamicPivot();
```

## 📋 API 參考

### CreatePivotOptions

```typescript
interface CreatePivotOptions {
  sourceSheet: string;     // 來源工作表名稱
  sourceRange: string;     // 資料範圍（A1:D100）
  targetSheet: string;     // 目標工作表名稱
  anchorCell: string;      // 樞紐分析表錨點（A3）
  layout: PivotLayout;     // 欄位配置
  refreshOnLoad?: boolean; // 開啟時自動重新整理（預設：true）
  styleName?: string;      // 樣式名稱（預設：PivotStyleMedium9）
}
```

### PivotLayout

```typescript
interface PivotLayout {
  rows?: PivotFieldSpec[];    // 行欄位（可選）
  cols?: PivotFieldSpec[];    // 列欄位（可選）
  values: PivotValueSpec[];   // 值欄位（必須）
}
```

### PivotFieldSpec

```typescript
interface PivotFieldSpec {
  name: string;  // 欄位名稱（必須與來源資料的標題列匹配）
}
```

### PivotValueSpec

```typescript
interface PivotValueSpec {
  name: string;            // 欄位名稱（必須是數值欄）
  agg?: PivotAgg;         // 聚合方式（預設：sum）
  displayName?: string;    // 顯示名稱（預設：欄位名稱）
  numFmtId?: number;       // 數字格式 ID（預設：0）
}
```

### PivotAgg

```typescript
type PivotAgg = "sum" | "count" | "average" | "max" | "min" | "product";
```

## 🎯 使用範例

### 範例 1：簡單的銷售分析

```javascript
const pivotOptions = {
  sourceSheet: "銷售資料",
  sourceRange: "A1:E1000",
  targetSheet: "分析報表",
  anchorCell: "A3",
  layout: {
    rows: [{ name: "產品類別" }],
    cols: [{ name: "月份" }],
    values: [{ name: "銷售額", agg: "sum" }]
  }
};
```

### 範例 2：多維度分析

```javascript
const pivotOptions = {
  sourceSheet: "訂單資料",
  sourceRange: "A1:F500",
  targetSheet: "訂單分析",
  anchorCell: "B5",
  layout: {
    rows: [
      { name: "客戶地區" },
      { name: "產品類別" }
    ],
    cols: [{ name: "季度" }],
    values: [
      { name: "訂單金額", agg: "sum", displayName: "總金額" },
      { name: "訂單數量", agg: "count", displayName: "訂單數" }
    ]
  },
  styleName: "PivotStyleLight16"
};
```

### 範例 3：財務報表

```javascript
const pivotOptions = {
  sourceSheet: "財務資料",
  sourceRange: "A1:G200",
  targetSheet: "財務分析",
  anchorCell: "C3",
  layout: {
    rows: [{ name: "部門" }],
    cols: [{ name: "會計年度" }],
    values: [
      { name: "收入", agg: "sum", numFmtId: 44 },      // 會計格式
      { name: "支出", agg: "sum", numFmtId: 44 },
      { name: "淨利", agg: "sum", numFmtId: 44 }
    ]
  },
  refreshOnLoad: true
};
```

## 🔧 進階配置

### 樣式選擇

可用的樞紐分析表樣式：
- `PivotStyleLight1` 到 `PivotStyleLight28`
- `PivotStyleMedium1` 到 `PivotStyleMedium28`
- `PivotStyleDark1` 到 `PivotStyleDark28`

### 數字格式

常用的數字格式 ID：
- `0`: 一般
- `44`: 會計格式
- `2`: 數字（小數點後 2 位）
- `4`: 百分比
- `9`: 日期

## 📊 資料準備要求

### 來源資料格式

1. **必須包含標題列**：第一行必須是欄位名稱
2. **資料範圍**：使用 A1 表示法（如 A1:D100）
3. **值欄位類型**：用於聚合的欄位必須是數值
4. **欄位名稱**：必須與 `PivotLayout` 中指定的名稱完全匹配

### 範例資料結構

```
| 部門 | 月份 | 產品 | 銷售額 |
|------|------|------|--------|
| IT   | 一月 | 軟體 | 50000  |
| IT   | 一月 | 硬體 | 30000  |
| HR   | 一月 | 培訓 | 20000  |
```

## 🚨 注意事項

### 重要提醒

1. **來源工作表必須存在**：確保指定的 `sourceSheet` 存在
2. **欄位名稱匹配**：`PivotLayout` 中的欄位名稱必須與來源資料的標題列完全匹配
3. **值欄位類型**：用於聚合的欄位必須是數值類型
4. **資料範圍**：`sourceRange` 必須包含標題列和至少一行資料

### 常見錯誤

```javascript
// ❌ 錯誤：欄位名稱不匹配
rows: [{ name: "部門名稱" }],  // 來源資料標題是 "部門"

// ❌ 錯誤：值欄位不是數值
values: [{ name: "產品名稱", agg: "sum" }],  // "產品名稱" 是文字

// ❌ 錯誤：資料範圍不包含標題
sourceRange: "A2:D100"  // 應該從 A1 開始
```

## 🧪 測試與驗證

### 測試腳本

使用內建的測試腳本：

```bash
npm run pivot:insert
```

### 驗證清單

1. ✅ 檔案成功生成
2. ✅ 檔案大小合理增加
3. ✅ 在 Excel 中開啟正常
4. ✅ 樞紐分析表顯示在指定位置
5. ✅ 欄位配置正確
6. ✅ 資料聚合正確
7. ✅ 支援重新整理

## 🔄 更新與維護

### 資料更新流程

1. 修改來源工作表的資料
2. 在樞紐分析表上按右鍵
3. 選擇「重新整理」
4. 樞紐分析表自動更新

### 樞紐分析表修改

如需修改樞紐分析表配置：
1. 重新執行 `addPivotToWorkbookBuffer`
2. 使用新的配置選項
3. 覆蓋原有檔案

## 🌟 最佳實踐

### 設計建議

1. **欄位選擇**：選擇有意義的行/列欄位組合
2. **值欄位**：優先使用數值欄位進行聚合
3. **樣式選擇**：根據報表用途選擇適當的樣式
4. **位置規劃**：預留足夠空間給樞紐分析表

### 效能優化

1. **資料範圍**：只包含必要的資料範圍
2. **欄位數量**：避免過多的行/列欄位
3. **值欄位**：限制值欄位數量，避免過度複雜

## 📚 相關資源

- [完整 API 文檔](./README-API.md)
- [技術實現報告](./IMPLEMENTATION_REPORT.md)
- [測試範例](./test/test-dynamic-pivot.js)

---

**xml-xlsx-lite 動態樞紐分析表** - 讓你的 Excel 自動化更強大！🚀
