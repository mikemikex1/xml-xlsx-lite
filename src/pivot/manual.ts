/**
 * 手動樞紐分析表 API
 * 提供純程式彙總的一條龍服務
 */

import { WorkbookImpl as Workbook } from '../workbook';
import { WorksheetImpl } from '../worksheet';
import type { Worksheet } from '../types';
import { createInvalidDataError } from '../errors';

export interface ManualPivotOptions {
  name?: string;
  rowField: string;
  columnField: string;
  valueField: string;
  aggregation?: 'sum' | 'avg' | 'count' | 'max' | 'min';
  numberFormat?: string;
  showRowTotals?: boolean;
  showColumnTotals?: boolean;
  showGrandTotal?: boolean;
  sortBy?: 'row' | 'column' | 'value';
  sortOrder?: 'asc' | 'desc';
}

export interface ManualPivotResult {
  worksheet: Worksheet;
  summary: {
    totalRows: number;
    totalColumns: number;
    totalValue: number;
    uniqueRowValues: number;
    uniqueColumnValues: number;
  };
}

/**
 * 在手動樞紐分析表中創建樞紐分析表
 * 支援 rowField / columnField / valueField / aggregation / numberFormat 等
 */
export function createManualPivotTable(
  this: Workbook,
  data: Array<Record<string, any>>,
  options: ManualPivotOptions
): ManualPivotResult {
  // 驗證輸入資料
  if (!Array.isArray(data) || data.length === 0) {
    throw createInvalidDataError('data', data, 'must be a non-empty array');
  }

  if (!options.rowField || !options.columnField || !options.valueField) {
    throw createInvalidDataError('options', options, 'rowField, columnField, and valueField are required');
  }

  // 檢查欄位是否存在於資料中
  const firstRow = data[0];
  if (!(options.rowField in firstRow) || !(options.columnField in firstRow) || !(options.valueField in firstRow)) {
    throw createInvalidDataError('options', options, `fields must exist in data: ${options.rowField}, ${options.columnField}, ${options.valueField}`);
  }

  // 設定預設值
  const opts: Required<ManualPivotOptions> = {
    name: options.name || 'Manual Pivot',
    rowField: options.rowField,
    columnField: options.columnField,
    valueField: options.valueField,
    aggregation: options.aggregation || 'sum',
    numberFormat: options.numberFormat || '#,##0',
    showRowTotals: options.showRowTotals ?? true,
    showColumnTotals: options.showColumnTotals ?? true,
    showGrandTotal: options.showGrandTotal ?? true,
    sortBy: options.sortBy || 'row',
    sortOrder: options.sortOrder || 'asc'
  };

  // 創建工作表
  const sheet = this.getWorksheet(opts.name);

  // 1) 分組和聚合資料
  const groupedData = groupAndAggregateData(data, opts);

  // 2) 排序資料
  const sortedData = sortGroupedData(groupedData, opts);

  // 3) 渲染樞紐分析表
  renderPivotTable(sheet, sortedData, opts);

  // 4) 計算統計資訊
  const summary = calculateSummary(sortedData, opts);

  return { worksheet: sheet, summary };
}

/**
 * 分組和聚合資料
 */
function groupAndAggregateData(
  data: Array<Record<string, any>>,
  options: Required<ManualPivotOptions>
): Map<string, Map<string, number[]>> {
  const grouped = new Map<string, Map<string, number[]>>();

  for (const row of data) {
    const rowValue = String(row[options.rowField]);
    const colValue = String(row[options.columnField]);
    const value = Number(row[options.valueField]);

    if (isNaN(value)) continue;

    // 初始化行
    if (!grouped.has(rowValue)) {
      grouped.set(rowValue, new Map());
    }

    const rowMap = grouped.get(rowValue)!;

    // 初始化列
    if (!rowMap.has(colValue)) {
      rowMap.set(colValue, []);
    }

    // 添加值
    rowMap.get(colValue)!.push(value);
  }

  return grouped;
}

/**
 * 聚合數值陣列
 */
function aggregateValues(values: number[], aggregation: string): number {
  switch (aggregation) {
    case 'sum':
      return values.reduce((sum, val) => sum + val, 0);
    case 'avg':
      return values.reduce((sum, val) => sum + val, 0) / values.length;
    case 'count':
      return values.length;
    case 'max':
      return Math.max(...values);
    case 'min':
      return Math.min(...values);
    default:
      return values.reduce((sum, val) => sum + val, 0);
  }
}

/**
 * 排序分組資料
 */
function sortGroupedData(
  groupedData: Map<string, Map<string, number[]>>,
  options: Required<ManualPivotOptions>
): Array<{ row: string; columns: Map<string, number>; rowTotal: number }> {
  const result: Array<{ row: string; columns: Map<string, number>; rowTotal: number }> = [];

  // 收集所有唯一的列值
  const allColumnValues = new Set<string>();
  for (const rowMap of groupedData.values()) {
    for (const colValue of rowMap.keys()) {
      allColumnValues.add(colValue);
    }
  }

  // 轉換為陣列格式
  for (const [rowValue, rowMap] of groupedData) {
    const columns = new Map<string, number>();
    let rowTotal = 0;

    // 填充所有列的值
    for (const colValue of allColumnValues) {
      const values = rowMap.get(colValue) || [];
      const aggregatedValue = aggregateValues(values, options.aggregation);
      columns.set(colValue, aggregatedValue);
      rowTotal += aggregatedValue;
    }

    result.push({ row: rowValue, columns, rowTotal });
  }

  // 排序
  result.sort((a, b) => {
    let comparison = 0;
    
    switch (options.sortBy) {
      case 'row':
        comparison = a.row.localeCompare(b.row);
        break;
      case 'column':
        comparison = a.rowTotal - b.rowTotal;
        break;
      case 'value':
        comparison = a.rowTotal - b.rowTotal;
        break;
    }

    return options.sortOrder === 'desc' ? -comparison : comparison;
  });

  return result;
}

/**
 * 渲染樞紐分析表
 */
function renderPivotTable(
  sheet: Worksheet,
  sortedData: Array<{ row: string; columns: Map<string, number>; rowTotal: number }>,
  options: Required<ManualPivotOptions>
): void {
  if (sortedData.length === 0) return;

  // 收集所有列值
  const allColumnValues = Array.from(sortedData[0].columns.keys()).sort();

  // 計算起始位置
  const startRow = 1;
  const startCol = 1;

  // 1. 標題行
  let colIndex = startCol;
  sheet.setCell(`${getColumnName(colIndex)}${startRow}`, options.rowField, { font: { bold: true } });
  colIndex++;

  for (const colValue of allColumnValues) {
    sheet.setCell(`${getColumnName(colIndex)}${startRow}`, colValue, { font: { bold: true } });
    colIndex++;
  }

  if (options.showRowTotals) {
    sheet.setCell(`${getColumnName(colIndex)}${startRow}`, 'Total', { font: { bold: true } });
  }

  // 2. 資料行
  for (let rowIndex = 0; rowIndex < sortedData.length; rowIndex++) {
    const rowData = sortedData[rowIndex];
    const excelRow = startRow + rowIndex + 1;

    // 行標籤
    colIndex = startCol;
    sheet.setCell(`${getColumnName(colIndex)}${excelRow}`, rowData.row);
    colIndex++;

    // 列值
    for (const colValue of allColumnValues) {
      const value = rowData.columns.get(colValue) || 0;
      sheet.setCell(`${getColumnName(colIndex)}${excelRow}`, value, { numFmt: options.numberFormat });
      colIndex++;
    }

    // 行總計
    if (options.showRowTotals) {
      sheet.setCell(`${getColumnName(colIndex)}${excelRow}`, rowData.rowTotal, { 
        numFmt: options.numberFormat,
        font: { bold: true }
      });
    }
  }

  // 3. 列總計行
  if (options.showColumnTotals) {
    const totalRow = startRow + sortedData.length + 1;
    let colIndex = startCol;
    
    sheet.setCell(`${getColumnName(colIndex)}${totalRow}`, 'Total', { font: { bold: true } });
    colIndex++;

    let grandTotal = 0;
    for (const colValue of allColumnValues) {
      const colTotal = sortedData.reduce((sum, row) => sum + (row.columns.get(colValue) || 0), 0);
      sheet.setCell(`${getColumnName(colIndex)}${totalRow}`, colTotal, { 
        numFmt: options.numberFormat,
        font: { bold: true }
      });
      grandTotal += colTotal;
      colIndex++;
    }

    if (options.showRowTotals) {
      sheet.setCell(`${getColumnName(colIndex)}${totalRow}`, grandTotal, { 
        numFmt: options.numberFormat,
        font: { bold: true }
      });
    }
  }

  // 4. 設定欄寬
  sheet.setColumnWidth(getColumnName(startCol), 15); // 行標籤欄
  for (let i = 0; i < allColumnValues.length; i++) {
    sheet.setColumnWidth(getColumnName(startCol + 1 + i), 12);
  }
  if (options.showRowTotals) {
    sheet.setColumnWidth(getColumnName(startCol + 1 + allColumnValues.length), 12);
  }
}

/**
 * 計算統計資訊
 */
function calculateSummary(
  sortedData: Array<{ row: string; columns: Map<string, number>; rowTotal: number }>,
  options: Required<ManualPivotOptions>
): { totalRows: number; totalColumns: number; totalValue: number; uniqueRowValues: number; uniqueColumnValues: number } {
  if (sortedData.length === 0) {
    return { totalRows: 0, totalColumns: 0, totalValue: 0, uniqueRowValues: 0, uniqueColumnValues: 0 };
  }

  const uniqueColumnValues = sortedData[0].columns.size;
  const totalValue = sortedData.reduce((sum, row) => sum + row.rowTotal, 0);

  return {
    totalRows: sortedData.length,
    totalColumns: uniqueColumnValues,
    totalValue,
    uniqueRowValues: sortedData.length,
    uniqueColumnValues
  };
}

/**
 * 將列索引轉換為 Excel 欄名
 */
function getColumnName(index: number): string {
  let result = '';
  while (index > 0) {
    index--;
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26);
  }
  return result || 'A';
}

/**
 * 將 createManualPivotTable 方法綁定到 Workbook 原型
 */
export function bindManualPivotMethods(): void {
  (Workbook as any).prototype.createManualPivotTable = createManualPivotTable;
}
