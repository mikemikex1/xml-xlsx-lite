/**
 * xml-xlsx-lite 讀取功能介面
 */

import { Workbook, Worksheet, Cell, CellOptions } from './types';
import { parseAddress } from './utils';

/**
 * 儲存格值類型
 */
export type CellValue = string | number | boolean | Date | null;

/**
 * 讀取選項
 */
export interface ReadOptions {
  /**
   * 是否啟用共享字串優化
   */
  enableSharedStrings?: boolean;
  
  /**
   * 共享字串閾值（字串長度超過此值時使用 sharedStrings）
   */
  sharedStringsThreshold?: number;
  
  /**
   * 是否啟用串流模式（大檔案）
   */
  streamingMode?: boolean;
  
  /**
   * 串流處理的塊大小
   */
  chunkSize?: number;
  
  /**
   * 是否保留樣式資訊
   */
  preserveStyles?: boolean;
  
  /**
   * 是否保留公式
   */
  preserveFormulas?: boolean;
  
  /**
   * 是否保留樞紐分析表
   */
  preservePivotTables?: boolean;
  
  /**
   * 是否保留圖表
   */
  preserveCharts?: boolean;
}

/**
 * 工作表讀取器介面
 */
export interface WorksheetReader extends Worksheet {
  /**
   * 將工作表轉換為二維陣列
   */
  toArray(): CellValue[][];
  
  /**
   * 將工作表轉換為 JSON 物件陣列
   */
  toJSON(opts?: { 
    headerRow?: number;
    includeEmptyRows?: boolean;
    includeEmptyColumns?: boolean;
  }): Record<string, CellValue>[];
  
  /**
   * 取得指定範圍的資料
   */
  getRange(range: string): CellValue[][];
  
  /**
   * 取得指定列的資料
   */
  getRow(row: number): CellValue[];
  
  /**
   * 取得指定欄的資料
   */
  getColumn(col: string | number): CellValue[];
  
  /**
   * 檢查儲存格是否存在
   */
  hasCell(address: string): boolean;
  
  /**
   * 取得工作表的維度資訊
   */
  getDimensions(): {
    minRow: number;
    maxRow: number;
    minCol: number;
    maxCol: number;
    usedRange: string;
  };
}

/**
 * 工作簿讀取器介面
 */
export interface WorkbookReader {
  /**
   * 從檔案讀取工作簿
   */
  readFile(path: string, options?: ReadOptions): Promise<Workbook>;
  
  /**
   * 從 Buffer 讀取工作簿
   */
  readBuffer(buf: ArrayBuffer, options?: ReadOptions): Promise<Workbook>;
  
  /**
   * 從 Stream 讀取工作簿
   */
  readStream(stream: ReadableStream, options?: ReadOptions): Promise<Workbook>;
  
  /**
   * 驗證檔案格式
   */
  validateFile(path: string): Promise<{
    isValid: boolean;
    format?: string;
    version?: string;
    sheets?: string[];
    errors?: string[];
  }>;
  
  /**
   * 取得檔案資訊
   */
  getFileInfo(path: string): Promise<{
    size: number;
    lastModified: Date;
    sheets: string[];
    hasStyles: boolean;
    hasSharedStrings: boolean;
    hasPivotTables: boolean;
    hasCharts: boolean;
  }>;
}

/**
 * 讀取器實現類別
 */
export class WorkbookReaderImpl implements WorkbookReader {
  async readFile(path: string, options: ReadOptions = {}): Promise<Workbook> {
    // 在 Node.js 環境中讀取檔案
    if (typeof window === 'undefined') {
      const fs = await import('fs');
      const buffer = fs.readFileSync(path);
      return this.readBuffer(buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength), options);
    } else {
      throw new Error('readFile is not supported in browser environment. Use readBuffer instead.');
    }
  }
  
  async readBuffer(buf: ArrayBuffer, options: ReadOptions = {}): Promise<Workbook> {
    try {
      // 引入必要的模組
      const JSZip = await import('jszip');
      const { parseXML } = await import('./xml-parser');
      const { WorkbookImpl } = await import('./workbook');
      const { WorksheetImpl } = await import('./worksheet');
      
      // 解壓縮 XLSX 檔案
      const zip = new JSZip.default();
      const zipContent = await zip.loadAsync(buf);
      
      // 驗證檔案格式
      const contentTypes = zipContent.file('[Content_Types].xml');
      if (!contentTypes) {
        throw new Error('Invalid XLSX file: missing [Content_Types].xml');
      }
      
      // 創建新的工作簿
      const workbook = new WorkbookImpl();
      
      // 讀取工作簿資訊
      const workbookXml = zipContent.file('xl/workbook.xml');
      if (!workbookXml) {
        throw new Error('Invalid XLSX file: missing xl/workbook.xml');
      }
      
      const workbookContent = await workbookXml.async('string');
      const workbookDoc = parseXML(workbookContent);
      
      // 讀取共享字串（如果存在）
      const sharedStrings = new Map<number, string>();
      const sharedStringsFile = zipContent.file('xl/sharedStrings.xml');
      if (sharedStringsFile) {
        const sharedStringsContent = await sharedStringsFile.async('string');
        const sharedStringsDoc = parseXML(sharedStringsContent);
        
        const stringItems = sharedStringsDoc.findAllDeep('si');
        stringItems.forEach((si, index) => {
          const textNode = si.findChild('t');
          if (textNode) {
            sharedStrings.set(index, textNode.getText());
          }
        });
      }
      
      // 讀取工作表
      const sheetsNode = workbookDoc.findChild('sheets');
      if (sheetsNode) {
        const sheetNodes = sheetsNode.findChildren('sheet');
        
        for (const sheetNode of sheetNodes) {
          const sheetName = sheetNode.getAttribute('name') || 'Sheet1';
          const sheetId = sheetNode.getAttribute('sheetId') || '1';
          const rId = sheetNode.getAttribute('r:id');
          
          // 讀取工作表檔案
          const sheetFile = zipContent.file(`xl/worksheets/sheet${sheetId}.xml`);
          if (sheetFile) {
            const sheetContent = await sheetFile.async('string');
            const sheetDoc = parseXML(sheetContent);
            
            // 創建工作表並解析資料
            const worksheet = workbook.getWorksheet(sheetName);
            this.parseWorksheetData(worksheet, sheetDoc, sharedStrings);
          }
        }
      }
      
      return workbook;
      
    } catch (error) {
      throw new Error(`Failed to read Excel file: ${error.message}`);
    }
  }
  
  async readStream(stream: ReadableStream, options: ReadOptions = {}): Promise<Workbook> {
    // TODO: 實現串流讀取邏輯
    throw new Error('readStream not yet implemented');
  }
  
  async validateFile(path: string): Promise<{
    isValid: boolean;
    format?: string;
    version?: string;
    sheets?: string[];
    errors?: string[];
  }> {
    // TODO: 實現檔案驗證邏輯
    throw new Error('validateFile not yet implemented');
  }
  
  async getFileInfo(path: string): Promise<{
    size: number;
    lastModified: Date;
    sheets: string[];
    hasStyles: boolean;
    hasSharedStrings: boolean;
    hasPivotTables: boolean;
    hasCharts: boolean;
  }> {
    // TODO: 實現檔案資訊讀取邏輯
    throw new Error('getFileInfo not yet implemented');
  }

  /**
   * 解析工作表資料
   */
  private parseWorksheetData(worksheet: any, sheetDoc: any, sharedStrings: Map<number, string>): void {
    
    // 找到 sheetData 節點
    const sheetDataNode = sheetDoc.findChild('sheetData');
    if (!sheetDataNode) {
      return;
    }
    
    // 解析每一行
    const rowNodes = sheetDataNode.findChildren('row');
    for (const rowNode of rowNodes) {
      const rowNumber = parseInt(rowNode.getAttribute('r') || '1');
      
      // 解析每個儲存格
      const cellNodes = rowNode.findChildren('c');
      for (const cellNode of cellNodes) {
        const cellRef = cellNode.getAttribute('r');
        if (!cellRef) continue;
        
        const cellType = cellNode.getAttribute('t') || 'n';
        let cellValue: any = null;
        
        // 根據儲存格類型解析值
        if (cellType === 's') {
          // 共享字串
          const valueNode = cellNode.findChild('v');
          if (valueNode) {
            const stringIndex = parseInt(valueNode.getText());
            cellValue = sharedStrings.get(stringIndex) || '';
          }
        } else if (cellType === 'inlineStr') {
          // 內聯字串
          const isNode = cellNode.findChild('is');
          if (isNode) {
            const textNode = isNode.findChild('t');
            if (textNode) {
              cellValue = textNode.getText();
            }
          }
        } else if (cellType === 'b') {
          // 布林值
          const valueNode = cellNode.findChild('v');
          if (valueNode) {
            cellValue = valueNode.getText() === '1';
          }
        } else {
          // 數值（包括日期）
          const valueNode = cellNode.findChild('v');
          if (valueNode) {
            const numValue = parseFloat(valueNode.getText());
            cellValue = isNaN(numValue) ? valueNode.getText() : numValue;
          }
        }
        
        // 設定儲存格值
        if (cellValue !== null) {
          worksheet.setCell(cellRef, cellValue);
        }
      }
    }
  }
}

/**
 * 工作表讀取器實現類別
 */
export class WorksheetReaderImpl implements WorksheetReader {
  // 繼承自 WorksheetImpl 的所有方法
  // 這裡只實現讀取相關的新方法
  
  // 基本屬性（需要從 WorksheetImpl 繼承）
  name: string = '';
  
  // 基本方法（需要從 WorksheetImpl 繼承）
  getCell(address: string): any {
    throw new Error('getCell not yet implemented');
  }
  
  setCell(address: string, value: any, options?: any): any {
    throw new Error('setCell not yet implemented');
  }
  
  *rows(): Generator<[number, Map<number, any>]> {
    throw new Error('rows not yet implemented');
  }
  
  // 讀取相關的新方法
  toArray(): CellValue[][] {
    const result: CellValue[][] = [];
    
    // 收集所有儲存格
    const cellMap = new Map<string, any>();
    let maxRow = 0;
    let maxCol = 0;
    
    // 假設 this 有 _cells 屬性（需要從 WorksheetImpl 繼承）
    if ((this as any)._cells) {
      for (const [address, cell] of (this as any)._cells) {
        const { row, col } = parseAddress(address);
        cellMap.set(`${row},${col}`, cell.value);
        maxRow = Math.max(maxRow, row);
        maxCol = Math.max(maxCol, col);
      }
    }
    
    // 創建二維陣列
    for (let row = 1; row <= maxRow; row++) {
      const rowData: CellValue[] = [];
      for (let col = 1; col <= maxCol; col++) {
        const value = cellMap.get(`${row},${col}`);
        rowData.push(value !== undefined ? value : null);
      }
      result.push(rowData);
    }
    
    return result;
  }
  
  toJSON(opts: { 
    headerRow?: number;
    includeEmptyRows?: boolean;
    includeEmptyColumns?: boolean;
  } = {}): Record<string, CellValue>[] {
    const { headerRow = 1, includeEmptyRows = false } = opts;
    const arrayData = this.toArray();
    
    if (arrayData.length === 0) {
      return [];
    }
    
    // 取得標題行
    const headers = arrayData[headerRow - 1] || [];
    const headerNames = headers.map((header, index) => 
      header ? String(header) : `Column${index + 1}`
    );
    
    // 轉換資料
    const result: Record<string, CellValue>[] = [];
    
    for (let i = headerRow; i < arrayData.length; i++) {
      const row = arrayData[i];
      const rowObj: Record<string, CellValue> = {};
      let hasData = false;
      
      for (let j = 0; j < Math.max(row.length, headerNames.length); j++) {
        const headerName = headerNames[j] || `Column${j + 1}`;
        const cellValue = j < row.length ? row[j] : null;
        
        rowObj[headerName] = cellValue;
        
        if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
          hasData = true;
        }
      }
      
      // 根據選項決定是否包含空行
      if (includeEmptyRows || hasData) {
        result.push(rowObj);
      }
    }
    
    return result;
  }
  
  getRange(range: string): CellValue[][] {
    // TODO: 實現範圍讀取邏輯
    throw new Error('getRange not yet implemented');
  }
  
  getRow(row: number): CellValue[] {
    // TODO: 實現列讀取邏輯
    throw new Error('getRow not yet implemented');
  }
  
  getColumn(col: string | number): CellValue[] {
    // TODO: 實現欄讀取邏輯
    throw new Error('getColumn not yet implemented');
  }
  
  hasCell(address: string): boolean {
    // TODO: 實現儲存格存在檢查邏輯
    throw new Error('hasCell not yet implemented');
  }
  
  getDimensions(): {
    minRow: number;
    maxRow: number;
    minCol: number;
    maxCol: number;
    usedRange: string;
  } {
    // TODO: 實現維度資訊讀取邏輯
    throw new Error('getDimensions not yet implemented');
  }
  
  // 其他必需的方法（暫時拋出錯誤）
  mergeCells(range: string): void {
    throw new Error('mergeCells not yet implemented');
  }
  
  unmergeCells(range: string): void {
    throw new Error('unmergeCells not yet implemented');
  }
  
  getMergedRanges(): string[] {
    throw new Error('getMergedRanges not yet implemented');
  }
  
  setColumnWidth(column: string | number, width: number): void {
    throw new Error('setColumnWidth not yet implemented');
  }
  
  getColumnWidth(column: string | number): number {
    throw new Error('getColumnWidth not yet implemented');
  }
  
  setRowHeight(row: number, height: number): void {
    throw new Error('setRowHeight not yet implemented');
  }
  
  getRowHeight(row: number): number {
    throw new Error('getRowHeight not yet implemented');
  }
  
  freezePanes(row?: number, column?: number): void {
    throw new Error('freezePanes not yet implemented');
  }
  
  unfreezePanes(): void {
    throw new Error('unfreezePanes not yet implemented');
  }
  
  getFreezePanes(): { row?: number; column?: number } {
    throw new Error('getFreezePanes not yet implemented');
  }
  
  setFormula(address: string, formula: string, options?: any): any {
    throw new Error('setFormula not yet implemented');
  }
  
  getFormula(address: string): string | null {
    throw new Error('getFormula not yet implemented');
  }
  
  validateFormula(formula: string): boolean {
    throw new Error('validateFormula not yet implemented');
  }
  
  getFormulaDependencies(address: string): string[] {
    throw new Error('getFormulaDependencies not yet implemented');
  }
  
  protect(password?: string, options?: any): void {
    throw new Error('protect not yet implemented');
  }
  
  unprotect(password?: string): void {
    throw new Error('unprotect not yet implemented');
  }
  
  isProtected(): boolean {
    throw new Error('isProtected not yet implemented');
  }
  
  getProtectionOptions(): any {
    throw new Error('getProtectionOptions not yet implemented');
  }
  
  addChart(chart: any): void {
    throw new Error('addChart not yet implemented');
  }
  
  removeChart(chartName: string): void {
    throw new Error('removeChart not yet implemented');
  }
  
  getCharts(): any[] {
    throw new Error('getCharts not yet implemented');
  }
  
  getChart(chartName: string): any {
    throw new Error('getChart not yet implemented');
  }
}
