# xml-xlsx-lite

[![npm version](https://badge.fury.io/js/xml-xlsx-lite.svg)](https://badge.fury.io/js/xml-xlsx-lite)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**Minimal XLSX writer using raw XML + JSZip, inspired by exceljs API**

一個輕量級的 Excel XLSX 檔案生成器，使用原生 XML 和 JSZip，API 設計參考 exceljs 的習慣用法。

## ✨ 特色

- 🚀 **輕量級**: 只包含核心功能，無多餘依賴
- 📝 **exceljs 相容**: API 設計參考 exceljs，學習成本低
- 🔧 **TypeScript 支援**: 完整的型別定義
- 🌐 **跨平台**: 支援 Node.js 和瀏覽器環境
- 📊 **多種資料型別**: 支援數字、字串、布林值、日期
- 📋 **多工作表**: 可建立和管理多個工作表
- 💾 **Shared Strings**: 自動處理字串重複，節省檔案大小
- ⚡ **寫入專用**: 專注於快速建立新的 Excel 檔案（不支援讀取或格式保留）

## 📦 安裝

```bash
npm install xml-xlsx-lite
```

## 🚀 快速開始

> **⚠️ 重要提醒**：xml-xlsx-lite 是「寫入專用」函式庫，用於建立新的 Excel 檔案。如果您需要修改現有檔案並保留樞紐表、圖表等格式，請使用 [exceljs](https://github.com/exceljs/exceljs) 或 [xlsx](https://github.com/SheetJS/sheetjs)。

### 基本使用

```javascript
import { Workbook } from 'xml-xlsx-lite';

// 建立工作簿
const wb = new Workbook();

// 取得工作表（如果不存在會自動建立）
const ws = wb.getWorksheet("Sheet1");

// 設定儲存格值
ws.setCell("A1", 123);
ws.setCell("B2", "Hello World");
ws.setCell("C3", true);
ws.setCell("D4", new Date());

// 生成 XLSX 檔案
const buffer = await wb.writeBuffer(); // ArrayBuffer
```

### 多工作表

```javascript
const wb = new Workbook();

// 建立多個工作表
const ws1 = wb.getWorksheet("工作表1");
const ws2 = wb.getWorksheet("工作表2");

ws1.setCell("A1", "工作表1的資料");
ws2.setCell("A1", "工作表2的資料");

// 也可以透過索引存取
const firstSheet = wb.getWorksheet(1);
```

### 便利方法

```javascript
const wb = new Workbook();

// 直接在工作簿上操作儲存格
wb.setCell("Sheet1", "A1", "便利方法");
const cell = wb.getCell("Sheet1", "A1");
```

### 瀏覽器下載

```javascript
const buffer = await wb.writeBuffer();

// 建立下載連結
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

## 📚 API 文件

### Workbook

主要的工作簿類別。

#### 建構函數

```typescript
new Workbook()
```

#### 方法

- `getWorksheet(nameOrIndex: string | number): Worksheet`
  - 取得工作表，如果不存在會自動建立
  - 支援名稱或索引（1-based）存取

- `getCell(worksheet: string | Worksheet, address: string): Cell`
  - 取得指定工作表的儲存格

- `setCell(worksheet: string | Worksheet, address: string, value: any, options?: CellOptions): Cell`
  - 設定指定工作表的儲存格值

- `writeBuffer(): Promise<ArrayBuffer>`
  - 生成 XLSX 檔案的 ArrayBuffer

### Worksheet

工作表類別。

#### 屬性

- `name: string` - 工作表名稱

#### 方法

- `getCell(address: string): Cell` - 取得儲存格
- `setCell(address: string, value: any, options?: CellOptions): Cell` - 設定儲存格值
- `rows(): Generator<[number, Map<number, Cell>]>` - 迭代所有行

### Cell

儲存格類別。

#### 屬性

- `address: string` - 儲存格位址（如 "A1"）
- `value: number | string | boolean | Date | null` - 儲存格值
- `type: 'n' | 's' | 'b' | 'd' | null` - 儲存格型別
- `options: CellOptions` - 儲存格選項（預留給未來功能）

### CellOptions

儲存格選項介面（預留給未來功能）。

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

## 🔧 開發

### 安裝依賴

```bash
npm install
```

### 建置

```bash
npm run build
```

### 測試

```bash
# Node.js 測試
npm test

# 瀏覽器測試
npm run test:browser
```

### 開發模式

```bash
npm run dev
```

## 📋 支援的資料型別

| 型別 | 說明 | Excel 對應 |
|------|------|------------|
| `number` | 數字 | 數值型別 |
| `string` | 字串 | 共享字串 |
| `boolean` | 布林值 | 布林型別 |
| `Date` | 日期 | Excel 序列號 |
| `null/undefined` | 空值 | 空儲存格 |

## 🚧 限制與未來規劃

### 目前限制

- 不支援儲存格樣式（字體、顏色、對齊等）
- 不支援公式
- 不支援合併儲存格
- 不支援欄寬/列高設定
- 不支援凍結窗格

### ⚠️ 重要注意事項

**檔案格式保留**：xml-xlsx-lite 是一個「寫入專用」的函式庫，專門用於從零開始建立新的 Excel 檔案。

- ✅ **適用場景**：產生報表、匯出資料、建立新的 Excel 檔案
- ❌ **不適用**：修改現有 Excel 檔案並保留格式

**如果您需要修改現有的 Excel 檔案並保留樞紐表、圖表、複雜格式等，請使用：**
- [exceljs](https://github.com/exceljs/exceljs) - 完整的 Excel 讀寫功能
- [xlsx](https://github.com/SheetJS/sheetjs) - 功能豐富的試算表處理函式庫

xml-xlsx-lite 的設計理念是「輕量、快速、簡單」，專注於高效率地產生新的 Excel 檔案。

### 未來規劃

- [ ] 儲存格樣式支援
- [ ] 公式支援
- [ ] 合併儲存格
- [ ] 欄寬/列高設定
- [ ] 凍結窗格
- [ ] 表格支援
- [ ] 資料驗證
- [ ] 篩選功能

## 🤝 貢獻

歡迎提交 Issue 和 Pull Request！

## 📄 授權

MIT License - 詳見 [LICENSE](LICENSE) 檔案

## 🙏 致謝

- [exceljs](https://github.com/exceljs/exceljs) - API 設計靈感
- [JSZip](https://github.com/Stuk/jszip) - ZIP 檔案處理
- [Office Open XML](https://en.wikipedia.org/wiki/Office_Open_XML) - 檔案格式規範

## 📞 支援

如果您遇到問題或有建議，請：

1. 查看 [Issues](https://github.com/mikemikex1/xml-xlsx-lite/issues)
2. 建立新的 Issue
3. 提交 Pull Request

---

**Made with ❤️ for the JavaScript community**
