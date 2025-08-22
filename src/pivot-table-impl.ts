/**
 * xml-xlsx-lite 樞紐分析表實現
 */

import { PivotTable, PivotTableConfig, PivotField, Worksheet, Cell } from './types';
import { parseAddress, addrFromRC } from './utils';

/**
 * 樞紐分析表實現類別
 */
export class PivotTableImpl implements PivotTable {
  name: string;
  config: PivotTableConfig;
  private data: any[][];
  private pivotData: any[][];
  private rowFields: PivotField[];
  private columnFields: PivotField[];
  private valueFields: PivotField[];
  private filters: Map<string, (string | number | boolean)[]>;

  constructor(name: string, config: PivotTableConfig) {
    this.name = name;
    this.config = config;
    this.data = [];
    this.pivotData = [];
    this.rowFields = [];
    this.columnFields = [];
    this.valueFields = [];
    this.filters = new Map();
    
    this.initializeFields();
  }

  /**
   * 初始化欄位分類
   */
  private initializeFields(): void {
    for (const field of this.config.fields) {
      switch (field.type) {
        case 'row':
          this.rowFields.push(field);
          break;
        case 'column':
          this.columnFields.push(field);
          break;
        case 'value':
          this.valueFields.push(field);
          break;
        case 'filter':
          // 篩選欄位暫時不處理
          break;
      }
    }
  }

  /**
   * 設定來源資料
   */
  setSourceData(data: any[][]): void {
    this.data = data;
    this.refresh();
  }

  /**
   * 重新整理樞紐分析表
   */
  refresh(): void {
    if (this.data.length === 0) return;
    
    // 解析來源範圍
    const sourceData = this.parseSourceData();
    
    // 建立樞紐分析表資料
    this.buildPivotData(sourceData);
  }

  /**
   * 解析來源資料
   */
  private parseSourceData(): any[] {
    const sourceRange = this.config.sourceRange;
    const { startRow, startCol, endRow, endCol } = this.parseRange(sourceRange);
    
    const sourceData = [];
    for (let row = startRow; row <= endRow; row++) {
      const rowData = [];
      for (let col = startCol; col <= endCol; col++) {
        const address = addrFromRC(row, col);
        // 這裡需要從工作表獲取資料，暫時使用模擬資料
        rowData.push(this.data[row - startRow]?.[col - startCol] || '');
      }
      sourceData.push(rowData);
    }
    
    return sourceData;
  }

  /**
   * 解析範圍字串
   */
  private parseRange(range: string): { startRow: number; startCol: number; endRow: number; endCol: number } {
    const parts = range.split(':');
    const start = parseAddress(parts[0]);
    const end = parseAddress(parts[1]);
    
    return {
      startRow: start.row,
      startCol: start.col,
      endRow: end.row,
      endCol: end.col
    };
  }

  /**
   * 建立樞紐分析表資料
   */
  private buildPivotData(sourceData: any[]): void {
    // 取得標題行
    const headers = sourceData[0] || [];
    const dataRows = sourceData.slice(1);
    
    // 建立欄位索引映射
    const fieldIndexMap = new Map<string, number>();
    headers.forEach((header, index) => {
      fieldIndexMap.set(header, index);
    });
    
    // 收集唯一值
    const uniqueValues = this.collectUniqueValues(dataRows, fieldIndexMap);
    
    // 建立樞紐分析表結構
    this.pivotData = this.createPivotStructure(uniqueValues, dataRows, fieldIndexMap);
  }

  /**
   * 收集唯一值
   */
  private collectUniqueValues(dataRows: any[][], fieldIndexMap: Map<string, number>): Map<string, Set<any>> {
    const uniqueValues = new Map<string, Set<any>>();
    
    // 初始化每個欄位的唯一值集合
    for (const field of [...this.rowFields, ...this.columnFields]) {
      const index = fieldIndexMap.get(field.sourceColumn);
      if (index !== undefined) {
        uniqueValues.set(field.sourceColumn, new Set());
      }
    }
    
    // 收集唯一值
    for (const row of dataRows) {
      for (const field of [...this.rowFields, ...this.columnFields]) {
        const index = fieldIndexMap.get(field.sourceColumn);
        if (index !== undefined && row[index] !== undefined) {
          uniqueValues.get(field.sourceColumn)!.add(row[index]);
        }
      }
    }
    
    return uniqueValues;
  }

  /**
   * 建立樞紐分析表結構
   */
  private createPivotStructure(
    uniqueValues: Map<string, Set<any>>, 
    dataRows: any[][], 
    fieldIndexMap: Map<string, number>
  ): any[][] {
    const pivotData: any[][] = [];
    
    // 建立標題行
    const titleRow = this.createTitleRow(uniqueValues);
    pivotData.push(titleRow);
    
    // 建立資料行
    const dataRowsResult = this.createDataRows(uniqueValues, dataRows, fieldIndexMap);
    pivotData.push(...dataRowsResult);
    
    return pivotData;
  }

  /**
   * 建立標題行
   */
  private createTitleRow(uniqueValues: Map<string, Set<any>>): any[] {
    const titleRow: any[] = [];
    
    // 添加行欄位標題
    for (const field of this.rowFields) {
      const values = Array.from(uniqueValues.get(field.sourceColumn) || []);
      titleRow.push(field.customName || field.sourceColumn);
      titleRow.push(...values);
    }
    
    // 添加值欄位標題
    for (const field of this.columnFields) {
      const values = Array.from(uniqueValues.get(field.sourceColumn) || []);
      titleRow.push(field.customName || field.sourceColumn);
      titleRow.push(...values);
    }
    
    // 添加值欄位標題
    for (const field of this.valueFields) {
      titleRow.push(field.customName || field.sourceColumn);
    }
    
    return titleRow;
  }

  /**
   * 建立資料行
   */
  private createDataRows(
    uniqueValues: Map<string, Set<any>>, 
    dataRows: any[][], 
    fieldIndexMap: Map<string, number>
  ): any[][] {
    const result: any[][] = [];
    
    // 這裡簡化處理，實際應該根據行欄位組合來分組
    for (const row of dataRows) {
      const pivotRow: any[] = [];
      
      // 添加行欄位值
      for (const field of this.rowFields) {
        const index = fieldIndexMap.get(field.sourceColumn);
        if (index !== undefined) {
          pivotRow.push(row[index]);
        }
      }
      
      // 添加值欄位值
      for (const field of this.valueFields) {
        const index = fieldIndexMap.get(field.sourceColumn);
        if (index !== undefined) {
          pivotRow.push(row[index]);
        }
      }
      
      result.push(pivotRow);
    }
    
    return result;
  }

  /**
   * 更新來源資料
   */
  updateSourceData(sourceRange: string): void {
    this.config.sourceRange = sourceRange;
    this.refresh();
  }

  /**
   * 取得欄位
   */
  getField(fieldName: string): PivotField | undefined {
    return this.config.fields.find(field => field.name === fieldName);
  }

  /**
   * 添加欄位
   */
  addField(field: PivotField): void {
    this.config.fields.push(field);
    this.initializeFields();
    this.refresh();
  }

  /**
   * 移除欄位
   */
  removeField(fieldName: string): void {
    const index = this.config.fields.findIndex(field => field.name === fieldName);
    if (index !== -1) {
      this.config.fields.splice(index, 1);
      this.initializeFields();
      this.refresh();
    }
  }

  /**
   * 重新排序欄位
   */
  reorderFields(fieldOrder: string[]): void {
    const newFields: PivotField[] = [];
    for (const fieldName of fieldOrder) {
      const field = this.config.fields.find(f => f.name === fieldName);
      if (field) {
        newFields.push(field);
      }
    }
    this.config.fields = newFields;
    this.initializeFields();
    this.refresh();
  }

  /**
   * 應用篩選
   */
  applyFilter(fieldName: string, filterValues: (string | number | boolean)[]): void {
    this.filters.set(fieldName, filterValues);
    this.refresh();
  }

  /**
   * 清除篩選
   */
  clearFilters(): void {
    this.filters.clear();
    this.refresh();
  }

  /**
   * 取得資料
   */
  getData(): any[][] {
    return this.pivotData;
  }

  /**
   * 匯出到工作表
   */
  exportToWorksheet(worksheetName: string): Worksheet {
    // 這裡需要返回一個工作表實例
    // 暫時拋出錯誤，因為需要工作簿實例
    throw new Error('exportToWorksheet requires workbook instance');
  }
}
