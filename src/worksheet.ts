import { CellModel } from './cell';
import { WorksheetProtection } from './protection';
import { ChartImpl } from './charts';
import { Cell, CellOptions, Worksheet, WorksheetProtectionOptions, Chart } from './types';
import { parseAddress, colToNumber, addrFromRC, getCellType } from './utils';

/**
 * 工作表實現類別
 */
export class WorksheetImpl implements Worksheet {
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

  // Phase 6: 工作表保護
  private _protection: WorksheetProtection = new WorksheetProtection();

  // Phase 6: 圖表支援
  private _charts: Map<string, Chart> = new Map();

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
    // 檢查工作表保護
    if (this._protection.isProtected() && !this._protection.isOperationAllowed('formatCells')) {
      throw new Error('Worksheet is protected. Cannot modify cells.');
    }

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
    if (this._protection.isProtected() && !this._protection.isOperationAllowed('formatCells')) {
      throw new Error('Worksheet is protected. Cannot merge cells.');
    }

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
    if (this._protection.isProtected() && !this._protection.isOperationAllowed('formatCells')) {
      throw new Error('Worksheet is protected. Cannot unmerge cells.');
    }

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
    if (this._protection.isProtected() && !this._protection.isOperationAllowed('formatColumns')) {
      throw new Error('Worksheet is protected. Cannot modify column width.');
    }

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
    if (this._protection.isProtected() && !this._protection.isOperationAllowed('formatRows')) {
      throw new Error('Worksheet is protected. Cannot modify row height.');
    }

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
    if (this._protection.isProtected() && !this._protection.isOperationAllowed('formatCells')) {
      throw new Error('Worksheet is protected. Cannot set formulas.');
    }

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

  // Phase 6: 工作表保護
  protect(password?: string, options?: any): void {
    this._protection.protect(password, options);
  }

  unprotect(password?: string): void {
    this._protection.unprotect(password);
  }

  isProtected(): boolean {
    return this._protection.isProtected();
  }

  getProtectionOptions(): any {
    return this._protection.getProtectionOptions();
  }

  // Phase 6: 圖表支援
  addChart(chart: Chart): void {
    if (this._protection.isProtected() && !this._protection.isOperationAllowed('objects')) {
      throw new Error('Worksheet is protected. Cannot add charts.');
    }

    this._charts.set(chart.name, chart);
  }

  removeChart(chartName: string): void {
    if (this._protection.isProtected() && !this._protection.isOperationAllowed('objects')) {
      throw new Error('Worksheet is protected. Cannot remove charts.');
    }

    this._charts.delete(chartName);
  }

  getCharts(): Chart[] {
    return Array.from(this._charts.values());
  }

  getChart(chartName: string): Chart | undefined {
    return this._charts.get(chartName);
  }

  // 內部方法
  private _rangesOverlap(range1: string, range2: string): boolean {
    const [start1, end1] = range1.split(':').map(parseAddress);
    const [start2, end2] = range2.split(':').map(parseAddress);
    
    // 檢查是否有重疊
    return !(end1.row < start2.row || start1.row > end2.row ||
             end1.col < start2.col || start1.col > end2.col);
  }

  // 取得內部屬性（用於 XML 生成）
  get _maxRowValue(): number { return this._maxRow; }
  get _maxColValue(): number { return this._maxCol; }
  get _columnWidthsValue(): Map<number, number> | undefined { return this._columnWidths; }
  get _rowHeightsValue(): Map<number, number> | undefined { return this._rowHeights; }
}
