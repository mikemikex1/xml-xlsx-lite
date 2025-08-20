import { PivotTable, PivotTableConfig, PivotField, Worksheet } from './types';
import { WorkbookImpl } from './workbook';
import { parseAddress, addrFromRC, colToNumber } from './utils';

/**
 * 動態 Pivot Table 實現類
 * 生成真正的 Excel 樞紐分析表，支援互動式操作
 */
export class PivotTableImpl implements PivotTable {
  name: string;
  config: PivotTableConfig;
  private _workbook: WorkbookImpl;
  private _sourceData: any[][] = [];
  private _processedData: any[][] = [];
  private _fieldValues: Map<string, Set<any>> = new Map();
  private _pivotCache: Map<string, any> = new Map();
  private _cacheId: number;
  private _tableId: number;

  constructor(name: string, config: PivotTableConfig, workbook: WorkbookImpl) {
    this.name = name;
    this.config = config;
    this._workbook = workbook;
    this._cacheId = this._generateCacheId();
    this._tableId = this._generateTableId();
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
    this._clearTargetWorksheet(ws);
    
    // 寫入 Pivot Table 資料
    this._writePivotData(ws);
    
    // 應用 Pivot Table 樣式
    this._applyPivotStyles(ws);
    
    return ws;
  }

  /**
   * 生成 PivotCache XML
   */
  generatePivotCacheXml(): string {
    const rowFields = this.config.fields.filter(f => f.type === 'row');
    const columnFields = this.config.fields.filter(f => f.type === 'column');
    const valueFields = this.config.fields.filter(f => f.type === 'value');
    const filterFields = this.config.fields.filter(f => f.type === 'filter');

    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
                     xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                     id="${this._cacheId}" 
                     recordCount="${this._sourceData.length - 1}" 
                     refreshOnLoad="1">`;

    // 快取來源
    xml += `
  <cacheSource type="worksheet">
    <worksheetSource ref="${this.config.sourceRange}" sheet="${this._getSourceSheetName()}"/>
  </cacheSource>`;

    // 欄位定義
    xml += `
  <cacheFields count="${this.config.fields.length}">`;
    
    for (const field of this.config.fields) {
      xml += this._generateCacheFieldXml(field);
    }
    
    xml += `
  </cacheFields>`;

    // 快取記錄
    xml += `
  <cacheRecords count="${this._sourceData.length - 1}">`;
    
    for (let i = 1; i < this._sourceData.length; i++) {
      xml += this._generateCacheRecordXml(this._sourceData[i], i);
    }
    
    xml += `
  </cacheRecords>`;

    xml += `
</pivotCacheDefinition>`;

    return xml;
  }

  /**
   * 生成 PivotTable XML
   */
  generatePivotTableXml(): string {
    const rowFields = this.config.fields.filter(f => f.type === 'row');
    const columnFields = this.config.fields.filter(f => f.type === 'column');
    const valueFields = this.config.fields.filter(f => f.type === 'value');
    const filterFields = this.config.fields.filter(f => f.type === 'filter');

    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
                     xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
                     mc:Ignorable="xr" 
                     xmlns:xr="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" 
                     name="${this.name}" 
                     cacheId="${this._cacheId}" 
                     dataCaption="Values" 
                     applyNumberFormatsInPivot="0" 
                     applyBorderFormatsInPivot="0" 
                     applyFontFormatsInPivot="0" 
                     applyPatternFormatsInPivot="0" 
                     applyAlignmentFormatsInPivot="0" 
                     applyWidthHeightFormatsInPivot="0" 
                     dataOnRows="0" 
                     dataPosition="0" 
                     grandTotalCaption="Grand Total" 
                     multipleFieldFilters="0" 
                     showDrill="1" 
                     showMemberPropertyTips="0" 
                     showMissing="0" 
                     showMultipleLabel="0" 
                     showPageMultipleLabel="0" 
                     showPageSubtotals="0" 
                     showRowGrandTotals="1" 
                     showRowSubtotals="1" 
                     showColGrandTotals="1" 
                     showColSubtotals="0" 
                     showItems="1" 
                     showDataDropDown="1" 
                     showError="0" 
                     showExpandCollapse="1" 
                     showOutline="1" 
                     showEmptyRow="0" 
                     showEmptyCol="0" 
                     showHeaders="1" 
                     compact="1" 
                     outline="1" 
                     outlineData="1" 
                     gridDropZones="0" 
                     indent="0" 
                     pageWrap="0" 
                     pageOverThenDown="0" 
                     pageDownThenOver="0" 
                     pageFieldsInReportFilter="0" 
                     pageWrapCount="0" 
                     pageBreakBetweenGroups="0" 
                     subtotalHiddenItems="0" 
                     printTitles="0" 
                     fieldPrintTitles="0" 
                     itemPrintTitles="0" 
                     mergeTitles="0" 
                     markAutoFormat="0" 
                     autoFormat="0" 
                     applyStyles="0" 
                     baseStyles="0" 
                     customListSort="1" 
                     applyDataValidation="0" 
                     enableDrill="1" 
                     fieldListSortAscending="0" 
                     mdxSubqueries="0" 
                     customSubtotals="0" 
                     visualTotals="1" 
                    showDataAs="0" 
                     calculatedMembers="0" 
                     visualTotalsFilters="0" 
                     showPageBreaks="0" 
                     useAutoFormat="0" 
                     pageGrandTotals="0" 
                     subtotalPageItems="0" 
                     rowGrandTotals="1" 
                     colGrandTotals="1" 
                     fieldSort="1" 
                     compactData="1" 
                     printDrill="0" 
                     itemDrill="0" 
                     drillThrough="0" 
                     fieldList="0" 
                     nonAutoSortDefault="0" 
                     showNew="0" 
                     autoShow="0" 
                     rankBy="0" 
                     defaultSubtotal="1" 
                     multipleItemSelectionMode="0" 
                     manualUpdate="0" 
                     showCalcMbrs="0" 
                     calculatedMembersInFilters="0" 
                     visualTotalsForSets="0" 
                     showASubtotalForPstTop="0" 
                     showAllDrill="0" 
                     showValue="1" 
                     expandMembersInDetail="0" 
                     dateFormatInPivot="0" 
                     pivotShowAs="0" 
                     enableWizard="0" 
                     enableDrill="1" 
                     enableFieldDialog="0" 
                     preserveFormatting="1" 
                     autoFormat="0" 
                     autoRepublish="0" 
                     showPageMultipleLabel="0" 
                     showPageSubtotals="0" 
                     showRowGrandTotals="1" 
                     showRowSubtotals="1" 
                     showColGrandTotals="1" 
                     showColSubtotals="0" 
                     showItems="1" 
                     showDataDropDown="1" 
                     showError="0" 
                     showExpandCollapse="1" 
                     showOutline="1" 
                     showEmptyRow="0" 
                     showEmptyCol="0" 
                     showHeaders="1" 
                     compact="1" 
                     outline="1" 
                     outlineData="1" 
                     gridDropZones="0" 
                     indent="0" 
                     pageWrap="0" 
                     pageOverThenDown="0" 
                     pageDownThenOver="0" 
                     pageFieldsInReportFilter="0" 
                     pageWrapCount="0" 
                     pageBreakBetweenGroups="0" 
                     subtotalHiddenItems="0" 
                     printTitles="0" 
                     fieldPrintTitles="0" 
                     itemPrintTitles="0" 
                     mergeTitles="0" 
                     markAutoFormat="0" 
                     autoFormat="0" 
                     applyStyles="0" 
                     baseStyles="0" 
                     customListSort="1" 
                     applyDataValidation="0" 
                     enableDrill="1" 
                     fieldListSortAscending="0" 
                     mdxSubqueries="0" 
                     customSubtotals="0" 
                     visualTotals="1" 
                     showDataAs="0" 
                     calculatedMembers="0" 
                     visualTotalsFilters="0" 
                     showPageBreaks="0" 
                     useAutoFormat="0" 
                     pageGrandTotals="0" 
                     subtotalPageItems="0" 
                     rowGrandTotals="1" 
                     colGrandTotals="1" 
                     fieldSort="1" 
                     compactData="1" 
                     printDrill="0" 
                     itemDrill="0" 
                     drillThrough="0" 
                     fieldList="0" 
                     nonAutoSortDefault="0" 
                     showNew="0" 
                     autoShow="0" 
                     rankBy="0" 
                     defaultSubtotal="1" 
                     multipleItemSelectionMode="0" 
                     manualUpdate="0" 
                     showCalcMbrs="0" 
                     calculatedMembersInFilters="0" 
                     visualTotalsForSets="0" 
                     showASubtotalForPstTop="0" 
                     showAllDrill="0" 
                     showValue="1" 
                     expandMembersInDetail="0" 
                     dateFormatInPivot="0" 
                     pivotShowAs="0" 
                     enableWizard="0" 
                     enableDrill="1" 
                     enableFieldDialog="0" 
                     preserveFormatting="1" 
                     autoFormat="0" 
                     autoRepublish="0">`;

    // 位置資訊
    xml += `
  <location firstDataCol="1" firstDataRow="1" firstHeaderRow="1" ref="${this.config.targetRange}"/>`;

    // 欄位配置
    xml += `
  <pivotFields count="${this.config.fields.length}">`;
    
    for (const field of this.config.fields) {
      xml += this._generatePivotFieldXml(field);
    }
    
    xml += `
  </pivotFields>`;

    // 行欄位
    if (rowFields.length > 0) {
      xml += `
  <rowFields count="${rowFields.length}">`;
      for (let i = 0; i < rowFields.length; i++) {
        xml += `
    <field x="${i}"/>`;
      }
      xml += `
  </rowFields>`;
    }

    // 列欄位
    if (columnFields.length > 0) {
      xml += `
  <colFields count="${columnFields.length}">`;
      for (let i = 0; i < columnFields.length; i++) {
        xml += `
    <field x="${i}"/>`;
      }
      xml += `
  </colFields>`;
    }

    // 值欄位
    if (valueFields.length > 0) {
      xml += `
  <dataFields count="${valueFields.length}">`;
      for (let i = 0; i < valueFields.length; i++) {
        const field = valueFields[i];
        xml += `
    <dataField name="${field.customName || field.name}" fld="${i}" baseField="0" baseItem="0" numFmtId="0" showDataAs="normal" subtotal="defaultFunction"/>`;
      }
      xml += `
  </dataFields>`;
    }

    // 篩選欄位
    if (filterFields.length > 0) {
      xml += `
  <pageFields count="${filterFields.length}">`;
      for (let i = 0; i < filterFields.length; i++) {
        xml += `
    <pageField fld="${i}" hier="-1"/>`;
      }
      xml += `
  </pageFields>`;
    }

    xml += `
</pivotTableDefinition>`;

    return xml;
  }

  /**
   * 生成 PivotTable 關聯 XML
   */
  generatePivotTableRelsXml(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${this._cacheId}.xml"/>
</Relationships>`;
  }

  /**
   * 取得快取 ID
   */
  getCacheId(): number {
    return this._cacheId;
  }

  /**
   * 取得表格 ID
   */
  getTableId(): number {
    return this._tableId;
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
    if (!this.config.showRowSubtotals) return;
    
    const rowFields = this.config.fields.filter(f => f.type === 'row');
    if (rowFields.length === 0) return;
    
    // 為每個行欄位添加小計行
    for (const field of rowFields) {
      if (field.showSubtotal) {
        this._addFieldSubtotal(field);
      }
    }
  }

  /**
   * 添加總計行
   */
  private _addGrandTotalRow(): void {
    if (!this.config.showGrandTotals) return;
    
    // 添加總計行到 Pivot Table 底部
    const totalRow: any[] = ['總計'];
    
    // 計算每列的總計
    const columnFields = this.config.fields.filter(f => f.type === 'column');
    for (const field of columnFields) {
      const values = this._getColumnValues(field.sourceColumn);
      const total = this._calculateFieldTotal(values, field.function);
      totalRow.push(total);
    }
    
    // 添加總計行到結果資料
    if (this._processedData.length > 0) {
      this._processedData.push(totalRow);
    }
  }

  /**
   * 計算總計
   */
  private _calculateTotals(): void {
    // 計算行總計
    this._calculateRowTotals();
    
    // 計算列總計
    this._calculateColumnTotals();
    
    // 計算總計
    this._calculateGrandTotal();
  }

  /**
   * 更新目標工作表
   */
  private _updateTargetWorksheet(): void {
    // 清除現有內容
    this._clearTargetWorksheet();
    
    // 寫入 Pivot Table 資料
    this._writePivotData();
    
    // 應用樣式
    this._applyPivotStyles();
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
   * 添加欄位小計
   */
  private _addFieldSubtotal(field: PivotField): void {
    // 實現欄位小計邏輯
    const fieldValues = this._fieldValues.get(field.sourceColumn);
    if (!fieldValues) return;
    
    // 為每個唯一值添加小計行
    for (const value of fieldValues) {
      const subtotalRow = [`${value} 小計`];
      // 這裡可以添加小計計算邏輯
      this._processedData.push(subtotalRow);
    }
  }

  /**
   * 取得欄位值
   */
  private _getColumnValues(columnName: string): any[] {
    const colIndex = this._getColumnIndex(columnName);
    if (colIndex === -1) return [];
    
    const values: any[] = [];
    for (let i = 1; i < this._sourceData.length; i++) {
      if (this._sourceData[i] && this._sourceData[i][colIndex] !== undefined) {
        values.push(this._sourceData[i][colIndex]);
      }
    }
    return values;
  }

  /**
   * 計算欄位總計
   */
  private _calculateFieldTotal(values: any[], functionType?: string): number {
    if (!values || values.length === 0) return 0;
    
    const numericValues = values.filter(v => typeof v === 'number');
    if (numericValues.length === 0) return 0;
    
    switch (functionType) {
      case 'sum':
        return numericValues.reduce((sum, val) => sum + val, 0);
      case 'count':
        return values.length;
      case 'average':
        return numericValues.reduce((sum, val) => sum + val, 0) / numericValues.length;
      case 'max':
        return Math.max(...numericValues);
      case 'min':
        return Math.min(...numericValues);
      default:
        return numericValues.reduce((sum, val) => sum + val, 0);
    }
  }

  /**
   * 計算行總計
   */
  private _calculateRowTotals(): void {
    // 實現行總計計算邏輯
    if (this._processedData.length === 0) return;
    
    // 為每行添加總計
    for (let i = 0; i < this._processedData.length; i++) {
      const row = this._processedData[i];
      if (row.length > 1) {
        const numericValues = row.slice(1).filter(v => typeof v === 'number');
        const rowTotal = numericValues.reduce((sum, val) => sum + val, 0);
        row.push(rowTotal);
      }
    }
  }

  /**
   * 計算列總計
   */
  private _calculateColumnTotals(): void {
    // 實現列總計計算邏輯
    if (this._processedData.length === 0) return;
    
    // 計算每列的總計
    const maxCols = Math.max(...this._processedData.map(row => row.length));
    for (let col = 0; col < maxCols; col++) {
      let colTotal = 0;
      for (let row = 0; row < this._processedData.length; row++) {
        const value = this._processedData[row][col];
        if (typeof value === 'number') {
          colTotal += value;
        }
      }
      // 將列總計添加到最後一行
      if (this._processedData.length > 0) {
        const lastRow = this._processedData[this._processedData.length - 1];
        lastRow[col] = colTotal;
      }
    }
  }

  /**
   * 計算總計
   */
  private _calculateGrandTotal(): void {
    // 實現總計計算邏輯
    if (this._processedData.length === 0) return;
    
    let grandTotal = 0;
    for (const row of this._processedData) {
      for (const value of row) {
        if (typeof value === 'number') {
          grandTotal += value;
        }
      }
    }
    
    // 將總計添加到最後一行的最後一列
    if (this._processedData.length > 0) {
      const lastRow = this._processedData[this._processedData.length - 1];
      lastRow.push(grandTotal);
    }
  }

  /**
   * 清除目標工作表
   */
  private _clearTargetWorksheet(ws?: Worksheet): void {
    // 實現清除目標工作表的邏輯
    // 這裡可以清除指定範圍的儲存格
  }

  /**
   * 寫入 Pivot Table 資料
   */
  private _writePivotData(ws?: Worksheet): void {
    // 實現寫入 Pivot Table 資料的邏輯
    // 這裡可以將處理後的資料寫入工作表
  }

  /**
   * 應用 Pivot Table 樣式
   */
  private _applyPivotStyles(ws?: Worksheet): void {
    // 實現應用 Pivot Table 樣式的邏輯
    // 這裡可以應用表格樣式、邊框等
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

  /**
   * 生成快取 ID
   */
  private _generateCacheId(): number {
    return Math.floor(Math.random() * 1000000) + 1;
  }

  /**
   * 生成表格 ID
   */
  private _generateTableId(): number {
    return Math.floor(Math.random() * 1000000) + 1;
  }

  /**
   * 取得來源工作表名稱
   */
  private _getSourceSheetName(): string {
    // 從來源範圍中提取工作表名稱
    // 暫時返回預設值
    return 'Sheet1';
  }

  /**
   * 生成快取欄位 XML
   */
  private _generateCacheFieldXml(field: PivotField): string {
    const values = Array.from(this._fieldValues.get(field.name) || []);
    
    let xml = `
    <cacheField name="${field.name}" numFmtId="0" formula="0" sqlType="0" hierarchy="0" level="0" databaseField="0" mappingCount="0" olap="0">`;
    
    // 欄位項目
    if (values.length > 0) {
      xml += `
      <sharedItems count="${values.length}" containsSemiMixedTypes="0" containsString="0" containsNumber="0" containsInteger="0" containsNonDate="0" containsDate="0" containsBlank="0" mixedTypes="0" minDate="1900-01-01T00:00:00" maxDate="9999-12-31T23:59:59" minNumber="0" maxNumber="0" minInteger="0" maxInteger="0" containsLocal="0" containsRemote="0" remote="0">`;
      
      for (const value of values) {
        if (typeof value === 'string') {
          xml += `
        <s v="${this._escapeXmlValue(value)}"/>`;
        } else if (typeof value === 'number') {
          xml += `
        <n v="${value}"/>`;
        }
      }
      
      xml += `
      </sharedItems>`;
    }
    
    xml += `
    </cacheField>`;
    
    return xml;
  }

  /**
   * 生成快取記錄 XML
   */
  private _generateCacheRecordXml(row: any[], index: number): string {
    let xml = `
      <r>`;
    
    for (const value of row) {
      if (typeof value === 'string') {
        xml += `
        <s v="${this._escapeXmlValue(value)}"/>`;
      } else if (typeof value === 'number') {
        xml += `
        <n v="${value}"/>`;
      } else if (typeof value === 'boolean') {
        xml += `
        <b v="${value ? '1' : '0'}"/>`;
      }
    }
    
    xml += `
      </r>`;
    
    return xml;
  }

  /**
   * 生成 Pivot 欄位 XML
   */
  private _generatePivotFieldXml(field: PivotField): string {
    let xml = `
    <pivotField axis="axisRow" showAll="0">`;
    
    // 欄位項目
    const values = Array.from(this._fieldValues.get(field.name) || []);
    if (values.length > 0) {
      xml += `
      <items count="${values.length + 1}">`;
      
      // 預設項目
      xml += `
        <item x="0" t="default"/>`;
      
      // 值項目
      for (let i = 0; i < values.length; i++) {
        xml += `
        <item x="${i + 1}"/>`;
      }
      
      xml += `
      </items>`;
    }
    
    xml += `
    </pivotField>`;
    
    return xml;
  }

  /**
   * 轉義 XML 值
   */
  private _escapeXmlValue(value: string): string {
    return value
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
