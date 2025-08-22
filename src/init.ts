/**
 * 初始化文件
 * 綁定所有相容性方法和手動樞紐分析表方法
 */

import { bindCompatibilityMethods } from './api/compat';
import { bindManualPivotMethods } from './pivot/manual';

/**
 * 初始化所有相容性方法
 * 在應用程式啟動時調用
 */
export function initializeXmlXlsxLite(): void {
  // 綁定 API 相容性方法
  bindCompatibilityMethods();
  
  // 綁定手動樞紐分析表方法
  bindManualPivotMethods();
  
  console.log('xml-xlsx-lite 相容性方法已初始化');
}

// 自動初始化（當模組被載入時）
initializeXmlXlsxLite();
