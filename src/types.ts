// Phase 1: 基本功能
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

// Phase 3: 進階功能
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
  
  // Phase 6: 工作表保護
  protect(password?: string, options?: WorksheetProtectionOptions): void;
  unprotect(password?: string): void;
  isProtected(): boolean;
  getProtectionOptions(): WorksheetProtectionOptions | null;
  
  // Phase 6: 圖表支援
  addChart(chart: Chart): void;
  removeChart(chartName: string): void;
  getCharts(): Chart[];
  getChart(chartName: string): Chart | undefined;
}

// Phase 5: Pivot Table 支援
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

// Phase 6: 工作表保護
export interface WorksheetProtectionOptions {
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;
  objects?: boolean;
  scenarios?: boolean;
}

// Phase 6: 圖表支援
export type ChartType = 'column' | 'line' | 'pie' | 'bar' | 'area' | 'scatter' | 'doughnut' | 'radar';

export interface ChartData {
  series: string; // 系列名稱
  categories: string; // 類別範圍，例如 "A2:A10"
  values: string; // 數值範圍，例如 "B2:B10"
  color?: string; // 系列顏色
}

export interface ChartOptions {
  title?: string;
  xAxisTitle?: string;
  yAxisTitle?: string;
  width?: number;
  height?: number;
  showLegend?: boolean;
  showDataLabels?: boolean;
  showGridlines?: boolean;
  theme?: 'light' | 'dark';
}

export interface Chart {
  name: string;
  type: ChartType;
  data: ChartData[];
  options: ChartOptions;
  position: {
    row: number;
    col: number;
  };
}

// Phase 4: 效能優化
export interface Workbook {
  getWorksheet(nameOrIndex: string | number): Worksheet;
  getWorksheets(): Worksheet[];
  getCell(worksheet: string | Worksheet, address: string): Cell;
  setCell(worksheet: string | Worksheet, address: string, value: number | string | boolean | Date | null, options?: CellOptions): Cell;
  writeBuffer(): Promise<ArrayBuffer>;
  writeFile(filename: string): Promise<void>;
  
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
  
  // Phase 6: 工作簿保護
  protect(password?: string, options?: WorkbookProtectionOptions): void;
  unprotect(password?: string): void;
  isProtected(): boolean;
  getProtectionOptions(): WorkbookProtectionOptions | null;
}

// Phase 6: 工作簿保護
export interface WorkbookProtectionOptions {
  structure?: boolean; // 保護工作簿結構
  windows?: boolean; // 保護工作簿視窗
  password?: string; // 保護密碼
}
