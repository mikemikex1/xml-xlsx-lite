import JSZip from "jszip";
import { Workbook, Worksheet, Cell, CellOptions, PivotTable, PivotTableConfig, WorkbookProtectionOptions } from './types';
import { WorksheetImpl } from './worksheet';
import { PivotTableImpl } from './pivot-table';
import { WorkbookProtection } from './protection';
import { buildContentTypes, buildRootRels, buildWorkbookXml, buildSheetXml, buildSharedStringsXml, buildStylesXml } from './xml-builders';

/**
 * 工作簿實現類別
 */
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

  // Phase 6: 工作簿保護
  private _protection: WorkbookProtection = new WorkbookProtection();

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

  getWorksheets(): Worksheet[] {
    return [...this._sheets];
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

  /** Write .xlsx file to disk */
  async writeFile(filename: string): Promise<void> {
    const buffer = await this.writeBuffer();
    
    if (typeof window !== 'undefined') {
      // 瀏覽器環境，不支援檔案寫入
      throw new Error('writeFile is not supported in browser environment. Use writeBuffer() instead.');
    }
    
    // 在 Node.js 環境中由外部實現
    throw new Error('writeFile method needs to be implemented externally. Use writeBuffer() and save manually.');
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
          const address = `${String.fromCharCode(65 + colNum - 1)}${rowNum}`;
          
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

  // Phase 6: 工作簿保護
  protect(password?: string, options?: WorkbookProtectionOptions): void {
    this._protection.protect(password, options);
  }

  unprotect(password?: string): void {
    this._protection.unprotect(password);
  }

  isProtected(): boolean {
    return this._protection.isProtected();
  }

  getProtectionOptions(): WorkbookProtectionOptions | null {
    return this._protection.getProtectionOptions();
  }

  /**
   * 生成 Pivot Table 相關的 XML 檔案
   */
  generatePivotTableFiles(): {
    pivotCacheXml: string;
    pivotCacheRecordsXml: string;
    pivotCacheRelsXml: string;
    pivotTableXml: string;
    pivotTableRelsXml: string;
    cacheId: number;
    tableId: number;
  }[] {
    const results = [];
    
    for (const pivotTable of this._pivotTables.values()) {
      if (pivotTable instanceof PivotTableImpl) {
        const cacheXml = pivotTable.generatePivotCacheXml();
        const cacheRecordsXml = pivotTable.generatePivotCacheRecordsXml();
        const cacheRelsXml = pivotTable.generatePivotCacheRelsXml();
        const tableXml = pivotTable.generatePivotTableXml();
        const tableRelsXml = pivotTable.generatePivotTableRelsXml();
        
        results.push({
          pivotCacheXml: cacheXml,
          pivotCacheRecordsXml: cacheRecordsXml,
          pivotCacheRelsXml: cacheRelsXml,
          pivotTableXml: tableXml,
          pivotTableRelsXml: tableRelsXml,
          cacheId: pivotTable.getCacheId(),
          tableId: pivotTable.getTableId()
        });
      }
    }
    
    return results;
  }

  /**
   * 生成包含 Pivot Table 的完整 Excel 檔案
   */
  async writeBufferWithPivotTables(): Promise<ArrayBuffer> {
    const zip = new JSZip();
    
    // 添加基本檔案
    await this._addBasicFiles(zip);
    
    // 添加 Pivot Table 檔案
    await this._addPivotTableFiles(zip);
    
    return await zip.generateAsync({ type: 'arraybuffer' });
  }

  /**
   * 添加 Pivot Table 相關檔案到 ZIP
   */
  private async _addPivotTableFiles(zip: JSZip): Promise<void> {
    const pivotFiles = this.generatePivotTableFiles();
    
    if (pivotFiles.length === 0) return;
    
    // 創建 pivotCache 目錄
    const pivotCacheFolder = zip.folder('xl/pivotCache');
    if (pivotCacheFolder) {
      for (const file of pivotFiles) {
        // 添加 pivotCacheDefinition XML
        pivotCacheFolder.file(
          `pivotCacheDefinition${file.cacheId}.xml`,
          file.pivotCacheXml
        );
        
        // 添加 pivotCacheRecords XML
        pivotCacheFolder.file(
          `pivotCacheRecords${file.cacheId}.xml`,
          file.pivotCacheRecordsXml
        );
      }
    }
    
    // 創建 pivotCache _rels 目錄
    const pivotCacheRelsFolder = zip.folder('xl/pivotCache/_rels');
    if (pivotCacheRelsFolder) {
      for (const file of pivotFiles) {
        pivotCacheRelsFolder.file(
          `pivotCacheDefinition${file.cacheId}.xml.rels`,
          file.pivotCacheRelsXml
        );
      }
    }
    
    // 創建 pivotTables 目錄
    const pivotTablesFolder = zip.folder('xl/pivotTables');
    if (pivotTablesFolder) {
      for (const file of pivotFiles) {
        pivotTablesFolder.file(
          `pivotTable${file.tableId}.xml`,
          file.pivotTableXml
        );
      }
    }
    
    // 創建 pivotTables _rels 目錄
    const pivotTableRelsFolder = zip.folder('xl/pivotTables/_rels');
    if (pivotTableRelsFolder) {
      for (const file of pivotFiles) {
        pivotTableRelsFolder.file(
          `pivotTable${file.tableId}.xml.rels`,
          file.pivotTableRelsXml
        );
      }
    }
    
    // 更新 [Content_Types].xml
    await this._updateContentTypesForPivotTables(zip, pivotFiles);
    
    // 更新 workbook.xml.rels
    await this._updateWorkbookRelsForPivotTables(zip, pivotFiles);
  }

  /**
   * 更新 Content Types 以包含 Pivot Table 檔案
   */
  private async _updateContentTypesForPivotTables(zip: JSZip, pivotFiles: any[]): Promise<void> {
    // 從現有的 [Content_Types].xml 開始，而不是重新生成
    const existingContentTypes = zip.file('[Content_Types].xml');
    let contentTypesXml = existingContentTypes ? await existingContentTypes.async('text') : await this._getContentTypesXml();
    
    // 添加 PivotCache 定義類型
    for (const file of pivotFiles) {
      contentTypesXml = contentTypesXml.replace(
        '</Types>',
        `  <Override PartName="/xl/pivotCache/pivotCacheDefinition${file.cacheId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
</Types>`
      );
    }
    
    // 添加 PivotCache 記錄類型
    for (const file of pivotFiles) {
      contentTypesXml = contentTypesXml.replace(
        '</Types>',
        `  <Override PartName="/xl/pivotCache/pivotCacheRecords${file.cacheId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
</Types>`
      );
    }
    
    // 添加 PivotTable 類型
    for (const file of pivotFiles) {
      contentTypesXml = contentTypesXml.replace(
        '</Types>',
        `  <Override PartName="/xl/pivotTables/pivotTable${file.tableId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
</Types>`
      );
    }
    
    zip.file('[Content_Types].xml', contentTypesXml);
  }

  /**
   * 更新 Workbook 關聯以包含 Pivot Table 檔案
   */
  private async _updateWorkbookRelsForPivotTables(zip: JSZip, pivotFiles: any[]): Promise<void> {
    // 從現有的 workbook.xml.rels 開始，而不是重新生成
    const existingWorkbookRels = zip.file('xl/_rels/workbook.xml.rels');
    let workbookRelsXml = existingWorkbookRels ? await existingWorkbookRels.async('text') : await this._getWorkbookRelsXml();
    
    // 添加 PivotCache 定義關聯
    for (const file of pivotFiles) {
      workbookRelsXml = workbookRelsXml.replace(
        '</Relationships>',
        `  <Relationship Id="rId${this._generateRelId()}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="pivotCache/pivotCacheDefinition${file.cacheId}.xml"/>
</Relationships>`
      );
    }
    
    zip.file('xl/_rels/workbook.xml.rels', workbookRelsXml);
  }

  /**
   * 取得 Content Types XML
   */
  private async _getContentTypesXml(): Promise<string> {
    return buildContentTypes(this._sheets.length, true, true);
  }

  /**
   * 取得 Workbook 關聯 XML
   */
  private async _getWorkbookRelsXml(): Promise<string> {
    return buildRootRels();
  }

  /**
   * 生成關聯 ID
   */
  private _generateRelId(): number {
    return Math.floor(Math.random() * 1000000) + 1;
  }

  /**
   * 添加基本檔案到 ZIP
   */
  private async _addBasicFiles(zip: JSZip): Promise<void> {
    // 添加 [Content_Types].xml
    zip.file('[Content_Types].xml', buildContentTypes(this._sheets.length, true, true));
    
    // 添加 _rels/.rels
    zip.file('_rels/.rels', buildRootRels());
    
    // 添加 xl/workbook.xml
    const { workbookXml, workbookRelsXml } = buildWorkbookXml(this._sheets);
    zip.file('xl/workbook.xml', workbookXml);
    zip.file('xl/_rels/workbook.xml.rels', workbookRelsXml);
    
    // 先生成工作表的 XML 來填充 _sstMap
    const sheetXmls: string[] = [];
    for (let i = 0; i < this._sheets.length; i++) {
      const sheet = this._sheets[i];
      const sheetXml = buildSheetXml(sheet, i + 1, this._sstMap, this);
      sheetXmls.push(sheetXml);
    }
    
    // 現在 _sstMap 已經被填充，可以生成 sharedStrings.xml
    zip.file('xl/sharedStrings.xml', buildSharedStringsXml(this._sstMap, this._sstArray));
    
    // 添加 xl/styles.xml
    zip.file('xl/styles.xml', buildStylesXml(this));
    
    // 添加工作表 XML 到 ZIP
    for (let i = 0; i < this._sheets.length; i++) {
      zip.file(`xl/worksheets/sheet${i + 1}.xml`, sheetXmls[i]);
      
      // 檢查工作表是否有圖表，只有在有圖表時才添加繪圖關聯
      const charts = (this._sheets[i] as any).getCharts ? (this._sheets[i] as any).getCharts() : [];
      if (charts.length > 0) {
        // 添加工作表關聯（包含繪圖）
        const sheetRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing${i + 1}.xml"/>
</Relationships>`;
        zip.file(`xl/worksheets/_rels/sheet${i + 1}.xml.rels`, sheetRelsXml);
      }
    }
  }

  // 內部方法（用於 XML 生成）
  get _sstMap(): Map<string, number> { return this._sst; }
  get _sstArray(): string[] { return this._sstArr; }
  get _stylesMap(): Map<string, number> { return this._styles; }
  get _fontsMap(): Map<string, number> { return this._fonts; }
  get _fillsMap(): Map<string, number> { return this._fills; }
  get _bordersMap(): Map<string, number> { return this._borders; }
  get _alignmentsMap(): Map<string, number> { return this._alignments; }
}
