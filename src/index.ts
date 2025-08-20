// 主入口文件 - 重構後的簡化版本

// 匯出所有類型定義
export * from './types';

// 匯出工具函數
export * from './utils';

// 匯出儲存格相關類別
export { CellModel } from './cell';

// 匯出工作表相關類別
export { WorksheetImpl } from './worksheet';

// 匯出保護相關類別
export { WorksheetProtection, WorkbookProtection } from './protection';

// 匯出圖表相關類別
export { ChartImpl, ChartFactory } from './charts';

// 匯出 Pivot Table 相關類別
export { PivotTableImpl } from './pivot-table';

// 匯出工作簿相關類別
export { WorkbookImpl } from './workbook';

// 匯出 XML 生成器
export * from './xml-builders';

// 主要匯出
import { WorkbookImpl } from './workbook';
export const Workbook = WorkbookImpl;

// 預設匯出
export default { Workbook: WorkbookImpl };
