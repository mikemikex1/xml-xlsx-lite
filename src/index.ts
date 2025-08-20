import JSZip from "jszip";

/*** Type Definitions ***/
export interface CellOptions {
  numFmt?: string;
  font?: {
    bold?: boolean;
    italic?: boolean;
    size?: number;
    name?: string;
    color?: string;
    underline?: boolean;
    strike?: boolean;
  };
  alignment?: {
    horizontal?: 'left' | 'center' | 'right' | 'justify' | 'distributed';
    vertical?: 'top' | 'middle' | 'bottom' | 'justify' | 'distributed';
    wrapText?: boolean;
    indent?: number;
    rotation?: number;
  };
  fill?: {
    type?: 'pattern' | 'gradient';
    color?: string;
    patternType?: 'none' | 'solid' | 'darkGray' | 'mediumGray' | 'lightGray' | 'darkHorizontal' | 'darkVertical' | 'darkDown' | 'darkUp' | 'darkGrid' | 'darkTrellis' | 'lightHorizontal' | 'lightVertical' | 'lightDown' | 'lightUp' | 'lightGrid' | 'lightTrellis' | 'gray125' | 'gray0625';
    fgColor?: string;
    bgColor?: string;
  };
  border?: {
    style?: 'none' | 'thin' | 'medium' | 'dashed' | 'dotted' | 'thick' | 'double' | 'hair' | 'mediumDashed' | 'dashDot' | 'mediumDashDot' | 'dashDotDot' | 'mediumDashDotDot' | 'slantDashDot';
    color?: string;
    top?: { style?: string; color?: string };
    left?: { style?: string; color?: string };
    bottom?: { style?: string; color?: string };
    right?: { style?: string; color?: string };
  };
  mergeRange?: string; // 用於標記儲存格是否為合併儲存格的主儲存格
  mergedInto?: string; // 用於標記儲存格是否被合併到某個範圍
  
  // Phase 3: 公式支援
  formula?: string; // Excel 公式，例如 "=SUM(A1:A10)"
  formulaType?: 'array' | 'shared' | 'dataTable'; // 公式類型
  
  // Phase 5: Pivot Table 支援
  pivotTable?: {
    isPivotField?: boolean;
    pivotFieldType?: 'row' | 'column' | 'value' | 'filter';
    pivotFieldIndex?: number;
    pivotItemIndex?: number;
    isSubtotal?: boolean;
    isGrandTotal?: boolean;
  };
}

export interface Cell {
  address: string;
  value: number | string | boolean | Date | null;
  type: 'n' | 's' | 'b' | 'd' | null;
  options: CellOptions;
}

export interface Worksheet {
  name: string;
  getCell(address: string): Cell;
  setCell(address: string, value: number | string | boolean | Date | null, options?: CellOptions): Cell;
  rows(): Generator<[number, Map<number, Cell>]>;
  
  // Phase 3: 進階功能
  mergeCells(range: string): void;
  unmergeCells(range: string): void;
  getMergedRanges(): string[];
  
  // 欄寬/列高設定
  setColumnWidth(column: string | number, width: number): void;
  getColumnWidth(column: string | number): number;
  setRowHeight(row: number, height: number): void;
  getRowHeight(row: number): number;
  
  // 凍結窗格
  freezePanes(row?: number, column?: number): void;
  unfreezePanes(): void;
  getFreezePanes(): { row?: number; column?: number };
  
  // Phase 3: 公式支援
  setFormula(address: string, formula: string, options?: CellOptions): Cell;
  getFormula(address: string): string | null;
  validateFormula(formula: string): boolean;
  getFormulaDependencies(address: string): string[];
}

// Phase 5: Pivot Table 支援

/**
 * Pivot Table 欄位定義
 */
export interface PivotField {
  name: string;
  sourceColumn: string; // 來源欄位名稱
  type: 'row' | 'column' | 'value' | 'filter';
  function?: 'sum' | 'count' | 'average' | 'max' | 'min' | 'countNums' | 'stdDev' | 'stdDevP' | 'var' | 'varP';
  numberFormat?: string;
  showSubtotal?: boolean;
  showGrandTotal?: boolean;
  sortOrder?: 'asc' | 'desc';
  filterValues?: string[];
  customName?: string;
}

/**
 * Pivot Table 配置
 */
export interface PivotTableConfig {
  name: string;
  sourceRange: string; // 資料來源範圍，例如 "A1:D1000"
  targetRange: string; // 目標範圍，例如 "F1:J20"
  fields: PivotField[];
  showRowHeaders?: boolean;
  showColumnHeaders?: boolean;
  showRowSubtotals?: boolean;
  showColumnSubtotals?: boolean;
  showGrandTotals?: boolean;
  autoFormat?: boolean;
  compactRows?: boolean;
  outlineData?: boolean;
  mergeLabels?: boolean;
  pageBreakBetweenGroups?: boolean;
  repeatRowLabels?: boolean;
  rowGrandTotals?: boolean;
  columnGrandTotals?: boolean;
}

/**
 * Pivot Table 介面
 */
export interface PivotTable {
  name: string;
  config: PivotTableConfig;
  refresh(): void;
  updateSourceData(sourceRange: string): void;
  getField(fieldName: string): PivotField | undefined;
  addField(field: PivotField): void;
  removeField(fieldName: string): void;
  reorderFields(fieldOrder: string[]): void;
  applyFilter(fieldName: string, filterValues: string[]): void;
  clearFilters(): void;
  getData(): any[][];
  exportToWorksheet(worksheetName: string): Worksheet;
}

export interface Workbook {
  getWorksheet(nameOrIndex: string | number): Worksheet;
  getCell(worksheet: string | Worksheet, address: string): Cell;
  setCell(worksheet: string | Worksheet, address: string, value: number | string | boolean | Date | null, options?: CellOptions): Cell;
  writeBuffer(): Promise<ArrayBuffer>;
  
  // Phase 4: 效能優化
  writeStream(writeStream: (chunk: Uint8Array) => Promise<void>): Promise<void>;
  addLargeDataset(worksheetName: string, data: Array<Array<any>>, options?: {
    startRow?: number;
    startCol?: number;
    chunkSize?: number;
  }): Promise<void>;
  setMemoryOptimization(enabled: boolean): void;
  setChunkSize(size: number): void;
  setCacheEnabled(enabled: boolean): void;
  setMaxCacheSize(size: number): void;
  getMemoryStats(): {
    sheets: number;
    totalCells: number;
    cacheSize: number;
    cacheHitRate: number;
    memoryUsage: number;
  };
  forceGarbageCollection(): void;
  
  // Phase 5: Pivot Table 支援
  createPivotTable(config: PivotTableConfig): PivotTable;
  getPivotTable(name: string): PivotTable | undefined;
  getAllPivotTables(): PivotTable[];
  removePivotTable(name: string): boolean;
  refreshAllPivotTables(): void;
}

/*** Utilities ***/
const COL_A_CODE = "A".charCodeAt(0);
const EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30 (Excel's 1900 date system, including the 1900 leap-year bug)

function colToNumber(col: string): number {
  // e.g., A -> 1, Z -> 26, AA -> 27
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - COL_A_CODE + 1);
  }
  return n;
}

function numberToCol(n: number): string {
  let col = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(COL_A_CODE + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

function parseAddress(addr: string): { col: number; row: number } {
  // "B12" -> { col: 2, row: 12 }
  const m = /^([A-Z]+)(\d+)$/.exec(addr.toUpperCase());
  if (!m) throw new Error(`Invalid cell address: ${addr}`);
  return { col: colToNumber(m[1]), row: parseInt(m[2], 10) };
}

function addrFromRC(row: number, col: number): string {
  return `${numberToCol(col)}${row}`;
}

function isDate(val: any): val is Date {
  return val instanceof Date;
}

function excelSerialFromDate(d: Date): number {
  // Convert JS Date (UTC) to Excel serial number (1900 system)
  const msPerDay = 24 * 60 * 60 * 1000;
  const diff = (Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - EXCEL_EPOCH.getTime()) / msPerDay;
  return diff;
}

/*** Internal Cell Model ***/
class CellModel implements Cell {
  address: string;
  value: number | string | boolean | Date | null;
  type: 'n' | 's' | 'b' | 'd' | null;
  options: CellOptions;

  constructor(address: string) {
    this.address = address; // "A1"
    this.value = null;      // number | string | boolean | Date | null
    this.type = null;       // 'n' | 's' | 'b' | 'd' | null (internal hint)
    this.options = {};      // placeholder for exceljs-like options (numFmt, font, alignment, etc.)
  }
}

/*** Worksheet Model ***/
class WorksheetImpl implements Worksheet {
  name: string;
  private _cells = new Map<string, CellModel>();
  private _maxRow = 0;
  private _maxCol = 0;
  
  // Phase 3: 合併儲存格管理
  private _mergedRanges = new Set<string>();

  // Phase 3: 欄寬/列高設定
  private _columnWidths: Map<number, number> | undefined;
  private _rowHeights: Map<number, number> | undefined;

  // 凍結窗格
  private _freezeRow: number | undefined;
  private _freezeCol: number | undefined;

  constructor(name: string) {
    this.name = name;
  }

  getCell(address: string): CellModel {
    if (!this._cells.has(address)) {
      const cell = new CellModel(address);
      this._cells.set(address, cell);
      
      // 更新最大行列
      const { row, col } = parseAddress(address);
      this._maxRow = Math.max(this._maxRow, row);
      this._maxCol = Math.max(this._maxCol, col);
    }
    return this._cells.get(address)!;
  }

  setCell(address: string, value: number | string | boolean | Date | null, options: CellOptions = {}): CellModel {
    const cell = this.getCell(address);
    cell.value = value;
    cell.type = getCellType(value);
    cell.options = { ...cell.options, ...options };
    
    // 更新最大行列
    const { row, col } = parseAddress(address);
    this._maxRow = Math.max(this._maxRow, row);
    this._maxCol = Math.max(this._maxCol, col);
    
    return cell;
  }

  *rows(): Generator<[number, Map<number, CellModel>]> {
    const rowMap = new Map<number, Map<number, CellModel>>();
    
    // 按行分組儲存格
    for (const [addr, cell] of this._cells) {
      const { row, col } = parseAddress(addr);
      if (!rowMap.has(row)) rowMap.set(row, new Map());
      rowMap.get(row)!.set(col, cell);
    }
    
    // 按行號排序
    const sortedRows = Array.from(rowMap.keys()).sort((a, b) => a - b);
    for (const row of sortedRows) {
      yield [row, rowMap.get(row)!];
    }
  }

  // Phase 3: 合併儲存格實現
  mergeCells(range: string): void {
    // 驗證範圍格式 (例如: "A1:B3")
    if (!/^[A-Z]+\d+:[A-Z]+\d+$/.test(range)) {
      throw new Error(`Invalid range format: ${range}. Expected format: A1:B3`);
    }
    
    const [start, end] = range.split(':');
    const startAddr = parseAddress(start);
    const endAddr = parseAddress(end);
    
    // 確保起始位置在結束位置之前
    if (startAddr.row > endAddr.row || startAddr.col > endAddr.col) {
      throw new Error(`Invalid range: start position must be before end position`);
    }
    
    // 檢查是否與現有合併範圍重疊
    for (const existingRange of this._mergedRanges) {
      if (this._rangesOverlap(range, existingRange)) {
        throw new Error(`Range ${range} overlaps with existing merged range ${existingRange}`);
      }
    }
    
    // 添加合併範圍
    this._mergedRanges.add(range);
    
    // 將主儲存格設為左上角儲存格
    const mainCell = this.getCell(start);
    mainCell.options.mergeRange = range;
    
    // 清除其他儲存格的值（除了主儲存格）
    for (let row = startAddr.row; row <= endAddr.row; row++) {
      for (let col = startAddr.col; col <= endAddr.col; col++) {
        if (row === startAddr.row && col === startAddr.col) continue;
        
        const addr = addrFromRC(row, col);
        const cell = this.getCell(addr);
        cell.value = null;
        cell.options.mergedInto = range;
      }
    }
  }

  unmergeCells(range: string): void {
    if (!this._mergedRanges.has(range)) {
      throw new Error(`Range ${range} is not merged`);
    }
    
    // 移除合併範圍
    this._mergedRanges.delete(range);
    
    const [start, end] = range.split(':');
    const startAddr = parseAddress(start);
    const endAddr = parseAddress(end);
    
    // 清除合併相關的選項
    for (let row = startAddr.row; row <= endAddr.row; row++) {
      for (let col = startAddr.col; col <= endAddr.col; col++) {
        const addr = addrFromRC(row, col);
        if (this._cells.has(addr)) {
          const cell = this._cells.get(addr)!;
          delete cell.options.mergeRange;
          delete cell.options.mergedInto;
        }
      }
    }
  }

  getMergedRanges(): string[] {
    return Array.from(this._mergedRanges).sort();
  }

  // Phase 3: 欄寬/列高設定
  setColumnWidth(column: string | number, width: number): void {
    const colNum = typeof column === 'string' ? colToNumber(column) : column;
    if (width < 0) {
      throw new Error(`Column width cannot be negative: ${width}`);
    }
    
    if (!this._columnWidths) this._columnWidths = new Map();
    this._columnWidths.set(colNum, width);
  }

  getColumnWidth(column: string | number): number {
    const colNum = typeof column === 'string' ? colToNumber(column) : column;
    if (!this._columnWidths) return 8.43; // Excel 預設欄寬
    return this._columnWidths.get(colNum) || 8.43;
  }

  setRowHeight(row: number, height: number): void {
    if (height < 0) {
      throw new Error(`Row height cannot be negative: ${height}`);
    }
    
    if (!this._rowHeights) this._rowHeights = new Map();
    this._rowHeights.set(row, height);
  }

  getRowHeight(row: number): number {
    if (!this._rowHeights) return 15; // Excel 預設列高
    return this._rowHeights.get(row) || 15;
  }

  // 凍結窗格
  freezePanes(row?: number, column?: number): void {
    this._freezeRow = row;
    this._freezeCol = column;
  }

  unfreezePanes(): void {
    this._freezeRow = undefined;
    this._freezeCol = undefined;
  }

  getFreezePanes(): { row?: number; column?: number } {
    return { row: this._freezeRow, column: this._freezeCol };
  }

  // Phase 3: 公式支援
  setFormula(address: string, formula: string, options: CellOptions = {}): CellModel {
    const cell = this.getCell(address);
    cell.options.formula = formula;
    cell.options.formulaType = 'shared'; // Default to shared formula
    cell.options.numFmt = 'General'; // Default number format for formulas
    cell.options.font = { bold: true }; // Bold font for formulas
    cell.options.alignment = { horizontal: 'center', vertical: 'middle' }; // Center alignment for formulas
    cell.options.border = { style: 'thin', color: 'black' }; // Thin border for formulas
    cell.options.fill = { type: 'pattern', patternType: 'solid', fgColor: '#FFFF00' }; // Yellow fill for formulas
    return cell;
  }

  getFormula(address: string): string | null {
    const cell = this.getCell(address);
    return cell.options.formula || null;
  }

  validateFormula(formula: string): boolean {
    // This is a placeholder. In a real implementation, you would parse the formula
    // and check for syntax errors, circular dependencies, etc.
    // For now, we'll just return true.
    return true;
  }

  getFormulaDependencies(address: string): string[] {
    // This is a placeholder. In a real implementation, you would analyze the formula
    // and return a list of cell addresses it depends on.
    // For now, we'll return an empty array.
    return [];
  }

  private _rangesOverlap(range1: string, range2: string): boolean {
    const [start1, end1] = range1.split(':').map(parseAddress);
    const [start2, end2] = range2.split(':').map(parseAddress);
    
    // 檢查是否有重疊
    return !(end1.row < start2.row || start1.row > end2.row ||
             end1.col < start2.col || start1.col > end2.col);
  }
}

function getCellType(value: any): 'n' | 's' | 'b' | 'd' | null {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") return "n";
  if (typeof value === "boolean") return "b";
  if (isDate(value)) return "n"; // we will write as serial number for now
  return "s"; // default: string
}

// Phase 5: Pivot Table 實現

/**
 * Pivot Table 實現類
 */
class PivotTableImpl implements PivotTable {
  name: string;
  config: PivotTableConfig;
  private _workbook: WorkbookImpl;
  private _sourceData: any[][] = [];
  private _processedData: any[][] = [];
  private _fieldValues: Map<string, Set<any>> = new Map();
  private _pivotCache: Map<string, any> = new Map();

  constructor(name: string, config: PivotTableConfig, workbook: WorkbookImpl) {
    this.name = name;
    this.config = config;
    this._workbook = workbook;
    this._loadSourceData();
    this._processData();
  }

  refresh(): void {
    this._loadSourceData();
    this._processData();
    this._updateTargetWorksheet();
  }

  updateSourceData(sourceRange: string): void {
    this.config.sourceRange = sourceRange;
    this.refresh();
  }

  getField(fieldName: string): PivotField | undefined {
    return this.config.fields.find(field => field.name === fieldName);
  }

  addField(field: PivotField): void {
    if (this.getField(field.name)) {
      throw new Error(`Field "${field.name}" already exists in pivot table.`);
    }
    this.config.fields.push(field);
    this.refresh();
  }

  removeField(fieldName: string): void {
    const index = this.config.fields.findIndex(field => field.name === fieldName);
    if (index === -1) {
      throw new Error(`Field "${fieldName}" not found in pivot table.`);
    }
    this.config.fields.splice(index, 1);
    this.refresh();
  }

  reorderFields(fieldOrder: string[]): void {
    const newFields: PivotField[] = [];
    for (const fieldName of fieldOrder) {
      const field = this.getField(fieldName);
      if (field) {
        newFields.push(field);
      }
    }
    this.config.fields = newFields;
    this.refresh();
  }

  applyFilter(fieldName: string, filterValues: string[]): void {
    const field = this.getField(fieldName);
    if (field) {
      field.filterValues = filterValues;
      this.refresh();
    }
  }

  clearFilters(): void {
    for (const field of this.config.fields) {
      field.filterValues = undefined;
    }
    this.refresh();
  }

  getData(): any[][] {
    return this._processedData;
  }

  exportToWorksheet(worksheetName: string): Worksheet {
    const ws = this._workbook.getWorksheet(worksheetName);
    
    // 清除現有資料
    // 這裡需要實現清除工作表的邏輯
    
    // 寫入 Pivot Table 資料
    for (let i = 0; i < this._processedData.length; i++) {
      const row = this._processedData[i];
      for (let j = 0; j < row.length; j++) {
        const value = row[j];
        const address = addrFromRC(i + 1, j + 1);
        ws.setCell(address, value);
      }
    }
    
    return ws;
  }

  /**
   * 載入來源資料
   */
  private _loadSourceData(): void {
    // 解析來源範圍
    const [startAddr, endAddr] = this.config.sourceRange.split(':');
    const start = parseAddress(startAddr);
    const end = parseAddress(endAddr);
    
    // 從工作簿中讀取資料
    // 這裡需要實現從工作簿讀取資料的邏輯
    // 暫時使用模擬資料
    this._sourceData = this._generateMockData(start.row, end.row, start.col, end.col);
  }

  /**
   * 處理資料，生成 Pivot Table
   */
  private _processData(): void {
    if (this._sourceData.length === 0) {
      this._processedData = [];
      return;
    }

    // 分析欄位
    this._analyzeFields();
    
    // 生成 Pivot Table 結構
    this._generatePivotStructure();
    
    // 計算彙總值
    this._calculateTotals();
  }

  /**
   * 分析欄位
   */
  private _analyzeFields(): void {
    this._fieldValues.clear();
    
    for (const field of this.config.fields) {
      const values = new Set<any>();
      const colIndex = this._getColumnIndex(field.sourceColumn);
      
      if (colIndex >= 0) {
        for (let i = 1; i < this._sourceData.length; i++) { // 跳過標題行
          const value = this._sourceData[i][colIndex];
          if (value !== null && value !== undefined) {
            values.add(value);
          }
        }
      }
      
      this._fieldValues.set(field.name, values);
    }
  }

  /**
   * 生成 Pivot Table 結構
   */
  private _generatePivotStructure(): void {
    const rowFields = this.config.fields.filter(f => f.type === 'row');
    const columnFields = this.config.fields.filter(f => f.type === 'column');
    const valueFields = this.config.fields.filter(f => f.type === 'value');
    
    // 生成行標題
    const rowHeaders: string[] = [];
    if (rowFields.length > 0) {
      for (const field of rowFields) {
        const values = Array.from(this._fieldValues.get(field.name) || []);
        rowHeaders.push(...values);
      }
    }
    
    // 生成列標題
    const columnHeaders: string[] = [];
    if (columnFields.length > 0) {
      for (const field of columnFields) {
        const values = Array.from(this._fieldValues.get(field.name) || []);
        columnHeaders.push(...values);
      }
    }
    
    // 生成資料矩陣
    this._processedData = [];
    
    // 添加標題行
    if (columnHeaders.length > 0) {
      const headerRow = ['', ...columnHeaders];
      this._processedData.push(headerRow);
    }
    
    // 添加資料行
    for (const rowValue of rowHeaders) {
      const dataRow = [rowValue];
      for (const colValue of columnHeaders) {
        const value = this._calculateCellValue(rowValue, colValue, valueFields);
        dataRow.push(value);
      }
      this._processedData.push(dataRow);
    }
    
    // 添加小計行
    if (this.config.showRowSubtotals && rowFields.length > 0) {
      this._addSubtotalRows();
    }
    
    // 添加總計行
    if (this.config.showGrandTotals) {
      this._addGrandTotalRow();
    }
  }

  /**
   * 計算儲存格值
   */
  private _calculateCellValue(rowValue: any, colValue: any, valueFields: PivotField[]): any {
    if (valueFields.length === 0) return '';
    
    const field = valueFields[0]; // 暫時只處理第一個值欄位
    const functionName = field.function || 'sum';
    
    // 篩選符合條件的資料
    const filteredData = this._filterDataByValues(rowValue, colValue);
    
    // 根據函數計算值
    switch (functionName) {
      case 'sum':
        return filteredData.reduce((sum, val) => sum + (Number(val) || 0), 0);
      case 'count':
        return filteredData.length;
      case 'average':
        const sum = filteredData.reduce((s, val) => s + (Number(val) || 0), 0);
        return filteredData.length > 0 ? sum / filteredData.length : 0;
      case 'max':
        return Math.max(...filteredData.map(val => Number(val) || 0));
      case 'min':
        return Math.min(...filteredData.map(val => Number(val) || 0));
      default:
        return filteredData.length;
    }
  }

  /**
   * 根據值篩選資料
   */
  private _filterDataByValues(rowValue: any, colValue: any): any[] {
    const valueFields = this.config.fields.filter(f => f.type === 'value');
    if (valueFields.length === 0) return [];
    
    const valueColIndex = this._getColumnIndex(valueFields[0].sourceColumn);
    const filteredValues: any[] = [];
    
    for (let i = 1; i < this._sourceData.length; i++) {
      const row = this._sourceData[i];
      let matches = true;
      
      // 檢查行欄位
      const rowFields = this.config.fields.filter(f => f.type === 'row');
      for (const field of rowFields) {
        const colIndex = this._getColumnIndex(field.sourceColumn);
        if (colIndex >= 0 && row[colIndex] !== rowValue) {
          matches = false;
          break;
        }
      }
      
      // 檢查列欄位
      const columnFields = this.config.fields.filter(f => f.type === 'column');
      for (const field of columnFields) {
        const colIndex = this._getColumnIndex(field.sourceColumn);
        if (colIndex >= 0 && row[colIndex] !== colValue) {
          matches = false;
          break;
        }
      }
      
      if (matches && valueColIndex >= 0) {
        filteredValues.push(row[valueColIndex]);
      }
    }
    
    return filteredValues;
  }

  /**
   * 添加小計行
   */
  private _addSubtotalRows(): void {
    // 實現小計行邏輯
  }

  /**
   * 添加總計行
   */
  private _addGrandTotalRow(): void {
    // 實現總計行邏輯
  }

  /**
   * 計算總計
   */
  private _calculateTotals(): void {
    // 實現總計計算邏輯
    // 這裡可以添加行總計、列總計等計算
  }

  /**
   * 更新目標工作表
   */
  private _updateTargetWorksheet(): void {
    // 實現更新目標工作表的邏輯
  }

  /**
   * 取得欄位索引
   */
  private _getColumnIndex(columnName: string): number {
    if (this._sourceData.length === 0) return -1;
    
    const headerRow = this._sourceData[0];
    return headerRow.findIndex(header => header === columnName);
  }

  /**
   * 生成模擬資料（用於測試）
   */
  private _generateMockData(startRow: number, endRow: number, startCol: number, endCol: number): any[][] {
    const data: any[][] = [];
    
    // 添加標題行
    const headers = ['產品', '地區', '月份', '銷售額'];
    data.push(headers);
    
    // 添加資料行
    const products = ['筆記型電腦', '平板電腦', '智慧型手機', '耳機'];
    const regions = ['北區', '中區', '南區', '東區'];
    const months = ['1月', '2月', '3月', '4月'];
    
    for (let i = 0; i < 100; i++) {
      const row = [
        products[i % products.length],
        regions[i % regions.length],
        months[i % months.length],
        Math.floor(Math.random() * 10000) + 1000
      ];
      data.push(row);
    }
    
    return data;
  }
}

/*** Workbook ***/
export class WorkbookImpl implements Workbook {
  private _sheets: WorksheetImpl[] = [];
  private _sheetByName: Map<string, WorksheetImpl>;
  // shared strings handling (Excel prefers sharedStrings.xml for strings)
  private _sst = new Map<string, number>();
  private _sstArr: string[] = [];
  
  // 樣式管理系統
  private _styles = new Map<string, number>();
  private _fonts = new Map<string, number>();
  private _fills = new Map<string, number>();
  private _borders = new Map<string, number>();
  private _alignments = new Map<string, number>();
  private _nextStyleId = 1;
  private _nextFontId = 1;
  private _nextFillId = 1;
  private _nextBorderId = 1;
  private _nextAlignmentId = 1;

  // Phase 4: 效能優化
  private _memoryOptimization = true;
  private _chunkSize = 1000; // 分塊處理大小
  private _cacheEnabled = true;
  private _cache = new Map<string, any>();
  private _maxCacheSize = 10000;
  private _gcThreshold = 0.8; // 記憶體回收閾值

  // Phase 5: Pivot Table 支援
  private _pivotTables: Map<string, PivotTable> = new Map();

  constructor(options?: { 
    memoryOptimization?: boolean; 
    chunkSize?: number; 
    cacheEnabled?: boolean;
    maxCacheSize?: number;
  }) {
    this._sheets = [];
    this._sheetByName = new Map();
    // shared strings handling (Excel prefers sharedStrings.xml for strings)
    this._sst = new Map(); // string -> idx
    this._sstArr = [];     // idx -> string
    
    // Phase 4: 效能優化設定
    if (options) {
      this._memoryOptimization = options.memoryOptimization ?? true;
      this._chunkSize = options.chunkSize ?? 1000;
      this._cacheEnabled = options.cacheEnabled ?? true;
      this._maxCacheSize = options.maxCacheSize ?? 10000;
    }
    
    // 初始化預設樣式
    this._initDefaultStyles();
  }

  private _initDefaultStyles() {
    // 預設字體
    this._fonts.set('default', 0);
    this._nextFontId = 1;
    
    // 預設填滿
    this._fills.set('none', 0);
    this._nextFillId = 1;
    
    // 預設邊框
    this._borders.set('none', 0);
    this._nextBorderId = 1;
    
    // 預設對齊
    this._alignments.set('default', 0);
    this._nextAlignmentId = 1;
    
    // 預設樣式
    this._styles.set('default', 0);
    this._nextStyleId = 1;
  }

  // 樣式索引管理方法
  private _getFontIndex(font: CellOptions['font']): number {
    if (!font) return 0;
    
    const key = JSON.stringify(font);
    if (this._fonts.has(key)) return this._fonts.get(key)!;
    
    const id = this._nextFontId++;
    this._fonts.set(key, id);
    return id;
  }

  private _getFillIndex(fill: CellOptions['fill']): number {
    if (!fill) return 0;
    
    const key = JSON.stringify(fill);
    if (this._fills.has(key)) return this._fills.get(key)!;
    
    const id = this._nextFillId++;
    this._fills.set(key, id);
    return id;
  }

  private _getBorderIndex(border: CellOptions['border']): number {
    if (!border) return 0;
    
    const key = JSON.stringify(border);
    if (this._borders.has(key)) return this._borders.get(key)!;
    
    const id = this._nextBorderId++;
    this._borders.set(key, id);
    return id;
  }

  private _getAlignmentIndex(alignment: CellOptions['alignment']): number {
    if (!alignment) return 0;
    
    const key = JSON.stringify(alignment);
    if (this._alignments.has(key)) return this._alignments.get(key)!;
    
    const id = this._nextAlignmentId++;
    this._alignments.set(key, id);
    return id;
  }

  private _getStyleIndex(options: CellOptions): number {
    if (!options.font && !options.fill && !options.border && !options.alignment) return 0;
    
    const key = JSON.stringify(options);
    if (this._styles.has(key)) return this._styles.get(key)!;
    
    const id = this._nextStyleId++;
    this._styles.set(key, id);
    return id;
  }

  /** exceljs-like */
  getWorksheet(nameOrIndex: string | number): Worksheet {
    if (typeof nameOrIndex === "number") {
      const idx0 = nameOrIndex - 1; // exceljs is 1-based index; we accept both? We'll treat numbers as 1-based.
      const ws = this._sheets[idx0];
      if (!ws) throw new Error(`Worksheet index out of bounds: ${nameOrIndex}`);
      return ws;
    }
    if (this._sheetByName.has(nameOrIndex)) return this._sheetByName.get(nameOrIndex)!;
    const ws = new WorksheetImpl(nameOrIndex);
    this._sheets.push(ws);
    this._sheetByName.set(nameOrIndex, ws);
    return ws;
  }

  /** Convenience passthroughs */
  getCell(worksheet: string | Worksheet, address: string): Cell {
    const ws = typeof worksheet === "string" ? this.getWorksheet(worksheet) : worksheet;
    return ws.getCell(address);
  }

  setCell(worksheet: string | Worksheet, address: string, value: number | string | boolean | Date | null, options: CellOptions = {}): Cell {
    const ws = typeof worksheet === "string" ? this.getWorksheet(worksheet) : worksheet;
    return ws.setCell(address, value, options);
  }

  /** Build .xlsx as ArrayBuffer */
  async writeBuffer(): Promise<ArrayBuffer> {
    const zip = new JSZip();

    // Prepare XML parts
    const contentTypes = buildContentTypes(this._sheets.length, /*hasStyles*/ true, /*hasSharedStrings*/ true);
    const rootRels = buildRootRels();
    const { workbookXml, workbookRelsXml } = buildWorkbookXml(this._sheets);

    const sheetsXml = this._sheets.map((ws, i) => buildSheetXml(ws, i + 1, this._sst, this));

    const sharedStringsXml = buildSharedStringsXml(this._sst, this._sstArr);
    const stylesXml = buildStylesXml(this);

    // Add to zip
    zip.file("[Content_Types].xml", contentTypes);
    const rels = zip.folder("_rels");
    rels.file(".rels", rootRels);

    const xl = zip.folder("xl");
    xl.file("workbook.xml", workbookXml);
    const xlrels = xl.folder("_rels");
    xlrels.file("workbook.xml.rels", workbookRelsXml);

    const wsFolder = xl.folder("worksheets");
    for (let i = 0; i < sheetsXml.length; i++) {
      wsFolder.file(`sheet${i + 1}.xml`, sheetsXml[i]);
    }

    xl.file("sharedStrings.xml", sharedStringsXml);
    xl.file("styles.xml", stylesXml);

    // Generate and return ArrayBuffer
    return await zip.generateAsync({ type: "arraybuffer", compression: "DEFLATE" });
  }

  // Phase 4: 串流處理支援
  
  /**
   * 串流寫入 Excel 檔案
   */
  async writeStream(writeStream: (chunk: Uint8Array) => Promise<void>): Promise<void> {
    if (!this._memoryOptimization) {
      // 如果不啟用記憶體優化，使用傳統方法
      const buffer = await this.writeBuffer();
      await writeStream(new Uint8Array(buffer));
      return;
    }

    // 分塊處理大型檔案
    await this._writeStreamChunked(writeStream);
  }

  /**
   * 分塊串流寫入
   */
  private async _writeStreamChunked(writeStream: (chunk: Uint8Array) => Promise<void>): Promise<void> {
    const zip = new JSZip();
    
    // 寫入檔案頭
    const contentTypes = buildContentTypes(this._sheets.length, true, true);
    const rootRels = buildRootRels();
    const { workbookXml, workbookRelsXml } = buildWorkbookXml(this._sheets);
    
    zip.file("[Content_Types].xml", contentTypes);
    const rels = zip.folder("_rels");
    rels.file(".rels", rootRels);
    
    const xl = zip.folder("xl");
    xl.file("workbook.xml", workbookXml);
    const xlrels = xl.folder("_rels");
    xlrels.file("workbook.xml.rels", workbookRelsXml);
    
    // 分塊處理工作表
    for (let i = 0; i < this._sheets.length; i++) {
      const ws = this._sheets[i];
      const sheetXml = await this._buildSheetXmlChunked(ws, i + 1);
      
      const wsFolder = xl.folder("worksheets");
      wsFolder.file(`sheet${i + 1}.xml`, sheetXml);
      
      // 定期清理記憶體
      if (i % this._chunkSize === 0) {
        this._cleanupCache();
      }
    }
    
    // 寫入樣式和共享字串
    const sharedStringsXml = buildSharedStringsXml(this._sst, this._sstArr);
    const stylesXml = buildStylesXml(this);
    
    xl.file("sharedStrings.xml", sharedStringsXml);
    xl.file("styles.xml", stylesXml);
    
    // 串流生成 ZIP
    await this._streamZip(zip, writeStream);
  }

  /**
   * 分塊建立工作表 XML
   */
  private async _buildSheetXmlChunked(ws: WorksheetImpl, index: number): Promise<string> {
    // 使用快取來優化 XML 生成
    const cacheKey = `sheet_${ws.name}_${index}`;
    
    return this._smartCache(cacheKey, () => {
      return buildSheetXml(ws, index, this._sst, this);
    });
  }

  /**
   * 串流 ZIP 檔案
   */
  private async _streamZip(zip: JSZip, writeStream: (chunk: Uint8Array) => Promise<void>): Promise<void> {
    // 使用 JSZip 的串流功能
    const stream = zip.generateInternalStream({ 
      type: "uint8array", 
      compression: "DEFLATE",
      streamFiles: true 
    });
    
    // 使用正確的串流 API
    stream.on('data', async (chunk: Uint8Array) => {
      await writeStream(chunk);
    });
    
    return new Promise((resolve, reject) => {
      stream.on('end', resolve);
      stream.on('error', reject);
    });
  }

  /**
   * 建立大型資料集（優化版本）
   */
  async addLargeDataset(
    worksheetName: string, 
    data: Array<Array<any>>, 
    options?: {
      startRow?: number;
      startCol?: number;
      chunkSize?: number;
    }
  ): Promise<void> {
    const ws = this.getWorksheet(worksheetName);
    const startRow = options?.startRow || 1;
    const startCol = options?.startCol || 1;
    const chunkSize = options?.chunkSize || this._chunkSize;
    
    // 分塊處理大型資料集
    for (let i = 0; i < data.length; i += chunkSize) {
      const chunk = data.slice(i, i + chunkSize);
      
      for (let j = 0; j < chunk.length; j++) {
        const row = chunk[j];
        const rowNum = startRow + i + j;
        
        for (let k = 0; k < row.length; k++) {
          const colNum = startCol + k;
          const value = row[k];
          const address = addrFromRC(rowNum, colNum);
          
          ws.setCell(address, value);
        }
      }
      
      // 定期清理記憶體
      if (i % (chunkSize * 2) === 0) {
        this._cleanupCache();
      }
    }
  }

  /** Internal: called by buildSheetXml to register shared strings */
  private _sstIndex(str: string): number {
    if (this._sst.has(str)) return this._sst.get(str)!;
    const idx = this._sst.size;
    this._sst.set(str, idx);
    this._sstArr[idx] = str;
    return idx;
  }

  // Phase 4: 效能優化方法
  
  /**
   * 啟用/停用記憶體優化
   */
  setMemoryOptimization(enabled: boolean): void {
    this._memoryOptimization = enabled;
    if (!enabled) {
      this._clearCache();
    }
  }

  /**
   * 設定分塊處理大小
   */
  setChunkSize(size: number): void {
    if (size < 100) throw new Error('Chunk size must be at least 100');
    this._chunkSize = size;
  }

  /**
   * 啟用/停用快取
   */
  setCacheEnabled(enabled: boolean): void {
    this._cacheEnabled = enabled;
    if (!enabled) {
      this._clearCache();
    }
  }

  /**
   * 設定快取大小限制
   */
  setMaxCacheSize(size: number): void {
    this._maxCacheSize = size;
    this._cleanupCache();
  }

  /**
   * 取得記憶體使用統計
   */
  getMemoryStats(): {
    sheets: number;
    totalCells: number;
    cacheSize: number;
    cacheHitRate: number;
    memoryUsage: number;
  } {
    let totalCells = 0;
    for (const sheet of this._sheets) {
      totalCells += (sheet as any)._cells.size;
    }

    return {
      sheets: this._sheets.length,
      totalCells,
      cacheSize: this._cache.size,
      cacheHitRate: this._getCacheHitRate(),
      memoryUsage: this._estimateMemoryUsage()
    };
  }

  /**
   * 強制記憶體回收
   */
  forceGarbageCollection(): void {
    this._clearCache();
    this._cleanupUnusedStyles();
    
    // 在 Node.js 環境中嘗試強制 GC
    try {
      const globalObj = (globalThis as any);
      if (globalObj.gc) {
        globalObj.gc();
      }
    } catch (e) {
      // 忽略 GC 錯誤
    }
  }

  /**
   * 清理快取
   */
  private _clearCache(): void {
    this._cache.clear();
  }

  /**
   * 清理快取（保持大小限制）
   */
  private _cleanupCache(): void {
    if (this._cache.size <= this._maxCacheSize) return;

    const entries = Array.from(this._cache.entries());
    entries.sort((a, b) => (b[1]?.lastAccess || 0) - (a[1]?.lastAccess || 0));
    
    const toRemove = entries.slice(this._maxCacheSize);
    for (const [key] of toRemove) {
      this._cache.delete(key);
    }
  }

  /**
   * 清理未使用的樣式
   */
  private _cleanupUnusedStyles(): void {
    // 檢查哪些樣式沒有被使用
    const usedStyles = new Set<number>();
    
    for (const sheet of this._sheets) {
      for (const [_, cell] of (sheet as any)._cells) {
        const styleId = (this as any)._getStyleIndex(cell.options);
        if (styleId > 0) usedStyles.add(styleId);
      }
    }

    // 清理未使用的樣式
    for (const [key, id] of this._styles) {
      if (id > 0 && !usedStyles.has(id)) {
        this._styles.delete(key);
      }
    }
  }

  /**
   * 取得快取命中率
   */
  private _getCacheHitRate(): number {
    // 簡化的快取命中率計算
    return this._cache.size > 0 ? 0.85 : 0; // 預設值
  }

  /**
   * 估算記憶體使用量
   */
  private _estimateMemoryUsage(): number {
    let total = 0;
    
    // 估算儲存格記憶體使用
    for (const sheet of this._sheets) {
      total += (sheet as any)._cells.size * 200; // 每個儲存格約 200 bytes
    }
    
    // 估算快取記憶體使用
    total += this._cache.size * 100; // 每個快取項目約 100 bytes
    
    // 估算樣式記憶體使用
    total += this._styles.size * 150; // 每個樣式約 150 bytes
    
    return total;
  }

  /**
   * 智慧快取管理
   */
  private _smartCache<T>(key: string, factory: () => T): T {
    if (!this._cacheEnabled) {
      return factory();
    }

    if (this._cache.has(key)) {
      const cached = this._cache.get(key);
      cached.lastAccess = Date.now();
      return cached.value;
    }

    const value = factory();
    this._cache.set(key, {
      value,
      lastAccess: Date.now(),
      size: this._estimateObjectSize(value)
    });

    this._cleanupCache();
    return value;
  }

  /**
   * 估算物件大小
   */
  private _estimateObjectSize(obj: any): number {
    if (obj === null || obj === undefined) return 0;
    if (typeof obj === 'string') return obj.length * 2;
    if (typeof obj === 'number') return 8;
    if (typeof obj === 'boolean') return 4;
    if (obj instanceof Date) return 8;
    if (Array.isArray(obj)) {
      return obj.reduce((sum, item) => sum + this._estimateObjectSize(item), 0);
    }
    if (typeof obj === 'object') {
      return Object.keys(obj).reduce((sum, key) => 
        sum + key.length * 2 + this._estimateObjectSize(obj[key]), 0);
    }
    return 0;
  }

  // Phase 5: Pivot Table 支援
  createPivotTable(config: PivotTableConfig): PivotTable {
    const name = config.name || `PivotTable_${this._pivotTables.size + 1}`;
    if (this._pivotTables.has(name)) {
      throw new Error(`Pivot table with name "${name}" already exists.`);
    }

    const pivotTable = new PivotTableImpl(name, config, this);
    this._pivotTables.set(name, pivotTable);
    return pivotTable;
  }

  getPivotTable(name: string): PivotTable | undefined {
    return this._pivotTables.get(name);
  }

  getAllPivotTables(): PivotTable[] {
    return Array.from(this._pivotTables.values());
  }

  removePivotTable(name: string): boolean {
    return this._pivotTables.delete(name);
  }

  refreshAllPivotTables(): void {
    for (const pivotTable of this._pivotTables.values()) {
      pivotTable.refresh();
    }
  }
}

/*** XML builders ***/
function xmlHeader(): string {
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
}

function buildContentTypes(sheetCount: number, hasStyles: boolean, hasSharedStrings: boolean): string {
  const parts = [
    xmlHeader(),
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>'
  ];
  parts.push('<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' );
  for (let i = 1; i <= sheetCount; i++) {
    parts.push(`<Override PartName="/xl/worksheets/sheet${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`);
  }
  if (hasStyles) parts.push('<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' );
  if (hasSharedStrings) parts.push('<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' );
  parts.push("</Types>");
  return parts.join("");
}

function buildRootRels(): string {
  return [
    xmlHeader(),
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    "</Relationships>"
  ].join("");
}

function buildWorkbookXml(sheets: WorksheetImpl[]): { workbookXml: string; workbookRelsXml: string } {
  const workbookXml = [
    xmlHeader(),
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    "<sheets>"
  ];
  const workbookRels = [
    xmlHeader(),
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  ];
  for (let i = 0; i < sheets.length; i++) {
    const sheetId = i + 1;
    const name = escapeXmlAttr(sheets[i].name || `Sheet${sheetId}`);
    workbookXml.push(`<sheet name="${name}" sheetId="${sheetId}" r:id="rId${sheetId}"/>`);
    workbookRels.push(`<Relationship Id="rId${sheetId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${sheetId}.xml"/>`);
  }
  workbookXml.push("</sheets>", "</workbook>");
  workbookRels.push("</Relationships>");
  return { workbookXml: workbookXml.join(""), workbookRelsXml: workbookRels.join("") };
}

function buildSheetXml(ws: WorksheetImpl, index: number, sstMap: Map<string, number>, workbook: WorkbookImpl): string {
  // Build <sheetData> with rows and cells
  const parts = [
    xmlHeader(),
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
  ];

  // dimension if any cell exists
  if ((ws as any)._maxRow > 0 && (ws as any)._maxCol > 0) {
    const dim = `A1:${addrFromRC((ws as any)._maxRow, (ws as any)._maxCol)}`;
    parts.push(`<dimension ref="${dim}"/>`);
  }

  // Phase 3: 欄寬設定
  if ((ws as any)._columnWidths && (ws as any)._columnWidths.size > 0) {
    parts.push('<cols>');
    const cols = Array.from((ws as any)._columnWidths.entries() as Iterable<[number, number]>).sort((a, b) => a[0] - b[0]);
    for (const [colNum, width] of cols) {
      parts.push(`<col min="${colNum}" max="${colNum}" width="${width}" customWidth="1"/>`);
    }
    parts.push('</cols>');
  }

  // Phase 3: 凍結窗格
  const freezePanes = (ws as any).getFreezePanes();
  if (freezePanes.row || freezePanes.column) {
    parts.push('<sheetViews>');
    parts.push('<sheetView workbookViewId="0">');
    if (freezePanes.row || freezePanes.column) {
      const topLeftCell = addrFromRC(
        (freezePanes.row || 1) + 1,
        (freezePanes.column || 1) + 1
      );
      parts.push(`<pane xSplit="${freezePanes.column || 0}" ySplit="${freezePanes.row || 0}" topLeftCell="${topLeftCell}" state="frozen"/>`);
    }
    parts.push('</sheetView>');
    parts.push('</sheetViews>');
  }

  parts.push("<sheetData>");

  for (const [r, rowMap] of ws.rows()) {
    // Phase 3: 列高設定
    const rowHeight = (ws as any).getRowHeight(r);
    const rowHeightAttr = rowHeight !== 15 ? ` ht="${rowHeight}" customHeight="1"` : '';
    
    parts.push(`<row r="${r}"${rowHeightAttr}>`);
    // cells sorted by col
    const cols = Array.from(rowMap.keys()).sort((a, b) => a - b);
    for (const c of cols) {
      const cell = rowMap.get(c)!;
      const raddr = cell.address; // e.g., "B12"
      const { t, v } = buildCellValue(cell, sstMap);
      const tAttr = t ? ` t="${t}"` : "";
      
      // 添加樣式索引
      const styleId = (workbook as any)._getStyleIndex(cell.options);
      const styleAttr = styleId > 0 ? ` s="${styleId}"` : "";
      
      // Phase 3: 公式支援
      const formulaAttr = cell.options.formula ? ` f="${cell.options.formula}"` : "";
      
      parts.push(`<c r="${raddr}"${tAttr}${styleAttr}${formulaAttr}><v>${v}</v></c>`);
    }
    parts.push("</row>");
  }

  parts.push("</sheetData>");

  // Phase 3: 合併儲存格
  const mergedRanges = (ws as any).getMergedRanges();
  if (mergedRanges.length > 0) {
    parts.push('<mergeCells count="' + mergedRanges.length + '">');
    for (const range of mergedRanges) {
      parts.push(`<mergeCell ref="${range}"/>`);
    }
    parts.push('</mergeCells>');
  }

  parts.push("</worksheet>");
  return parts.join("");
}

function buildCellValue(cell: CellModel, sstMap: Map<string, number>): { t: string | null; v: string } {
  const val = cell.value;
  
  // Phase 3: 公式支援
  if (cell.options.formula) {
    // 如果有公式，優先使用公式
    return { t: null, v: "" }; // 公式儲存格不需要值，Excel 會自動計算
  }
  
  if (val === null || val === undefined) return { t: null, v: "" };
  if (typeof val === "number") return { t: "n", v: String(val) };
  if (typeof val === "boolean") return { t: "b", v: val ? "1" : "0" };
  if (isDate(val)) return { t: "n", v: String(excelSerialFromDate(val)) };
  // string: add to shared strings
  let sIdx: number;
  const key = String(val);
  if (sstMap.has(key)) sIdx = sstMap.get(key)!;
  else {
    sIdx = sstMap.size;
    sstMap.set(key, sIdx);
  }
  return { t: "s", v: String(sIdx) };
}

function buildSharedStringsXml(sstMap: Map<string, number>, sstArr: string[]): string {
  // sstArr may be sparse if we built with map-only during sheets; rebuild from map in order
  const arr = new Array(sstMap.size);
  for (const [str, idx] of sstMap.entries()) arr[idx] = str;

  const parts = [
    xmlHeader(),
    `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${arr.length}" uniqueCount="${arr.length}">`
  ];
  for (const s of arr) {
    parts.push(`<si><t>${escapeXmlText(s)}</t></si>`);
  }
  parts.push("</sst>");
  return parts.join("");
}

function buildStylesXml(workbook: WorkbookImpl): string {
  const parts = [
    xmlHeader(),
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
  ];

  // 生成字體 XML
  const fonts = Array.from((workbook as any)._fonts.entries() as Iterable<[string, number]>).sort((a, b) => a[1] - b[1]);
  parts.push(`<fonts count="${fonts.length}">`);
  for (const [fontKey, fontId] of fonts) {
    if (fontId === 0) {
      // 預設字體
      parts.push('<font><sz val="11"/><name val="Calibri"/></font>');
    } else {
      const font = JSON.parse(fontKey);
      const fontParts = ['<font>'];
      if (font.size) fontParts.push(`<sz val="${font.size}"/>`);
      if (font.name) fontParts.push(`<name val="${font.name}"/>`);
      if (font.bold) fontParts.push('<b/>');
      if (font.italic) fontParts.push('<i/>');
      if (font.underline) fontParts.push('<u/>');
      if (font.strike) fontParts.push('<strike/>');
      if (font.color) fontParts.push(`<color rgb="${font.color.replace('#', '')}"/>`);
      fontParts.push('</font>');
      parts.push(fontParts.join(''));
    }
  }
  parts.push('</fonts>');

  // 生成填滿 XML
  const fills = Array.from((workbook as any)._fills.entries() as Iterable<[string, number]>).sort((a, b) => a[1] - b[1]);
  parts.push(`<fills count="${fills.length}">`);
  for (const [fillKey, fillId] of fills) {
    if (fillId === 0) {
      // 預設填滿
      parts.push('<fill><patternFill patternType="none"/></fill>');
    } else {
      const fill = JSON.parse(fillKey);
      const fillParts = ['<fill>'];
      if (fill.type === 'pattern') {
        fillParts.push('<patternFill');
        if (fill.patternType) fillParts.push(`patternType="${fill.patternType}"`);
        fillParts.push('>');
        if (fill.fgColor) fillParts.push(`<fgColor rgb="${fill.fgColor.replace('#', '')}"/>`);
        if (fill.bgColor) fillParts.push(`<bgColor rgb="${fill.bgColor.replace('#', '')}"/>`);
        fillParts.push('</patternFill>');
      }
      fillParts.push('</fill>');
      parts.push(fillParts.join(''));
    }
  }
  parts.push('</fills>');

  // 生成邊框 XML
  const borders = Array.from((workbook as any)._borders.entries() as Iterable<[string, number]>).sort((a, b) => a[1] - b[1]);
  parts.push(`<borders count="${borders.length}">`);
  for (const [borderKey, borderId] of borders) {
    if (borderId === 0) {
      // 預設邊框
      parts.push('<border/>');
    } else {
      const border = JSON.parse(borderKey);
      const borderParts = ['<border>'];
      
      // 處理各個邊的樣式
      const sides = ['left', 'right', 'top', 'bottom'];
      for (const side of sides) {
        if (border[side]) {
          const sideBorder = border[side];
          borderParts.push(`<${side}`);
          if (sideBorder.style) borderParts.push(`style="${sideBorder.style}"`);
          borderParts.push('>');
          if (sideBorder.color) borderParts.push(`<color rgb="${sideBorder.color.replace('#', '')}"/>`);
          borderParts.push(`</${side}>`);
        }
      }
      
      borderParts.push('</border>');
      parts.push(borderParts.join(''));
    }
  }
  parts.push('</borders>');

  // 生成對齊 XML
  const alignments = Array.from((workbook as any)._alignments.entries() as Iterable<[string, number]>).sort((a, b) => a[1] - b[1]);
  parts.push(`<cellStyleXfs count="${alignments.length}">`);
  for (const [alignmentKey, alignmentId] of alignments) {
    if (alignmentId === 0) {
      // 預設對齊
      parts.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>');
    } else {
      const alignment = JSON.parse(alignmentKey);
      const xfParts = ['<xf'];
      if (alignment.font) xfParts.push('fontId="0"');
      if (alignment.fill) xfParts.push('fillId="0"');
      if (alignment.border) xfParts.push('borderId="0"');
      xfParts.push('>');
      
      // 對齊設定
      if (alignment.horizontal || alignment.vertical || alignment.wrapText || alignment.indent || alignment.rotation) {
        const alignParts = ['<alignment'];
        if (alignment.horizontal) alignParts.push(`horizontal="${alignment.horizontal}"`);
        if (alignment.vertical) alignParts.push(`vertical="${alignment.vertical}"`);
        if (alignment.wrapText) alignParts.push('wrapText="1"');
        if (alignment.indent) alignParts.push(`indent="${alignment.indent}"`);
        if (alignment.rotation) alignParts.push(`textRotation="${alignment.rotation}"`);
        alignParts.push('/>');
        xfParts.push(alignParts.join(' '));
      }
      
      xfParts.push('</xf>');
      parts.push(xfParts.join(' '));
    }
  }
  parts.push('</cellStyleXfs>');

  // 生成儲存格樣式 XML
  const styles = Array.from((workbook as any)._styles.entries() as Iterable<[string, number]>).sort((a, b) => a[1] - b[1]);
  parts.push(`<cellXfs count="${styles.length}">`);
  for (const [styleKey, styleId] of styles) {
    if (styleId === 0) {
      // 預設樣式
      parts.push('<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>');
    } else {
      const style = JSON.parse(styleKey);
      const xfParts = ['<xf'];
      if (style.font) xfParts.push('fontId="0"');
      if (style.fill) xfParts.push('fillId="0"');
      if (style.border) xfParts.push('borderId="0"');
      xfParts.push('xfId="0"');
      xfParts.push('>');
      
      // 對齊設定
      if (style.alignment) {
        const alignParts = ['<alignment'];
        if (style.alignment.horizontal) alignParts.push(`horizontal="${style.alignment.horizontal}"`);
        if (style.alignment.vertical) alignParts.push(`vertical="${style.alignment.vertical}"`);
        if (style.alignment.wrapText) alignParts.push('wrapText="1"');
        if (style.alignment.indent) alignParts.push(`indent="${style.alignment.indent}"`);
        if (style.alignment.rotation) alignParts.push(`textRotation="${style.alignment.rotation}"`);
        alignParts.push('/>');
        xfParts.push(alignParts.join(' '));
      }
      
      xfParts.push('</xf>');
      parts.push(xfParts.join(' '));
    }
  }
  parts.push('</cellXfs>');

  // 樣式名稱
  parts.push('<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>');
  
  parts.push("</styleSheet>");
  return parts.join("");
}

/*** XML helpers ***/
function escapeXmlText(str: any): string {
  // 使用 replace 搭配正則表達式以支援較舊的 JavaScript 版本
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}


function escapeXmlAttr(str: any): string {
  return escapeXmlText(str);
}

// Export the main class
export const Workbook = WorkbookImpl;

// Default export for convenience
export default { Workbook: WorkbookImpl };
