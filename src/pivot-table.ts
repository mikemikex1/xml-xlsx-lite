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
    // 注意：這裡不調用 _updateTargetWorksheet，因為它需要工作表參數
    // 實際的更新應該在 exportToWorksheet 中進行
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
   * 生成 PivotCache 定義 XML（不包含記錄）
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

    xml += `
</pivotCacheDefinition>`;

    return xml;
  }

  /**
   * 生成 PivotCache 記錄 XML（獨立檔案）
   */
  generatePivotCacheRecordsXml(): string {
    let xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" 
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
                  count="${this._sourceData.length - 1}">`;
    
    // 生成記錄
    for (let i = 1; i < this._sourceData.length; i++) {
      xml += this._generateCacheRecordXml(this._sourceData[i], i);
    }
    
    xml += `
</pivotCacheRecords>`;

    return xml;
  }

  /**
   * 生成 PivotCache 關聯 XML
   */
  generatePivotCacheRelsXml(): string {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords${this._cacheId}.xml"/>
</Relationships>`;
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
    try {
      // 嘗試從工作簿中讀取真實資料
      if (this.config.sourceRange) {
        const sourceData = this._loadDataFromWorkbook();
        if (sourceData && sourceData.length > 0) {
          this._sourceData = sourceData;
          return;
        }
      }
    } catch (e) {
      // 如果讀取失敗，記錄錯誤並使用模擬資料
      console.warn(`無法從工作簿讀取資料: ${e.message}，使用模擬資料`);
    }
    
    // 使用模擬資料作為備用方案
    const [startAddr, endAddr] = this.config.sourceRange ? this.config.sourceRange.split(':') : ['A1', 'D100'];
    const start = parseAddress(startAddr);
    const end = parseAddress(endAddr);
    this._sourceData = this._generateMockData(start.row, end.row, start.col, end.col);
  }

  /**
   * 從工作簿中讀取資料
   */
  private _loadDataFromWorkbook(): any[][] | null {
    if (!this.config.sourceRange) return null;
    
    try {
      // 解析來源範圍
      const [startAddr, endAddr] = this.config.sourceRange.split(':');
      const start = parseAddress(startAddr);
      const end = parseAddress(endAddr);
      
      // 嘗試從工作簿中讀取資料
      const worksheets = this._workbook.getWorksheets();
      if (worksheets.length === 0) return null;
      
      // 使用第一個工作表作為來源
      const sourceSheet = worksheets[0];
      const data: any[][] = [];
      
      // 讀取標題行
      const headers: string[] = [];
      for (let col = start.col; col <= end.col; col++) {
        const address = `${String.fromCharCode(65 + col - 1)}${start.row}`;
        try {
          const cell = sourceSheet.getCell(address);
          headers.push(cell ? String(cell.value || `欄位${col}`) : `欄位${col}`);
        } catch (e) {
          headers.push(`欄位${col}`);
        }
      }
      data.push(headers);
      
      // 讀取資料行
      for (let row = start.row + 1; row <= end.row; row++) {
        const rowData: any[] = [];
        for (let col = start.col; col <= end.col; col++) {
          const address = `${String.fromCharCode(65 + col - 1)}${row}`;
          try {
            const cell = sourceSheet.getCell(address);
            rowData.push(cell ? cell.value : null);
          } catch (e) {
            rowData.push(null);
          }
        }
        data.push(rowData);
      }
      
      return data;
    } catch (e) {
      console.warn(`讀取工作簿資料時發生錯誤: ${e.message}`);
      return null;
    }
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
    if (columnHeaders.length > 0 && valueFields.length > 0) {
      // 為每個值欄位生成標題
      const headerRow = ['']; // 第一列為空（行標題列）
      
      if (valueFields.length === 1) {
        // 單一值欄位：直接添加列標題
        headerRow.push(...columnHeaders);
      } else {
        // 多個值欄位：為每個值欄位添加列標題
        for (const valueField of valueFields) {
          const functionName = valueField.function || 'sum';
          const fieldName = valueField.customName || valueField.name;
          
          for (const colValue of columnHeaders) {
            headerRow.push(`${colValue} - ${fieldName} (${functionName})`);
          }
        }
      }
      
      this._processedData.push(headerRow);
    }
    
    // 添加資料行
    for (const rowValue of rowHeaders) {
      const dataRow = [rowValue];
      
      if (columnHeaders.length > 0 && valueFields.length > 0) {
        if (valueFields.length === 1) {
          // 單一值欄位
          for (const colValue of columnHeaders) {
            const value = this._calculateCellValue(rowValue, colValue, valueFields);
            dataRow.push(value);
          }
        } else {
          // 多個值欄位
          for (const valueField of valueFields) {
            for (const colValue of columnHeaders) {
              const value = this._calculateCellValue(rowValue, colValue, [valueField]);
              dataRow.push(value);
            }
          }
        }
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
    
    // 支援多個值欄位
    const results: any[] = [];
    
    for (const field of valueFields) {
      const functionName = field.function || 'sum';
      
      // 篩選符合條件的資料
      const filteredData = this._filterDataByValues(rowValue, colValue, field.sourceColumn);
      
      // 根據函數計算值
      const result = this._calculateValueByFunction(filteredData, functionName);
      results.push(result);
    }
    
    // 如果只有一個值欄位，返回單一值；否則返回陣列
    return valueFields.length === 1 ? results[0] : results;
  }

  /**
   * 根據函數計算值
   */
  private _calculateValueByFunction(data: any[], functionName: string): any {
    if (data.length === 0) return 0;
    
    const numericData = data.map(val => Number(val)).filter(val => !isNaN(val));
    
    switch (functionName) {
      case 'sum':
        return numericData.reduce((sum, val) => sum + val, 0);
      case 'count':
        return data.length;
      case 'countNums':
        return numericData.length;
      case 'average':
        return numericData.length > 0 ? numericData.reduce((sum, val) => sum + val, 0) / numericData.length : 0;
      case 'max':
        return numericData.length > 0 ? Math.max(...numericData) : 0;
      case 'min':
        return numericData.length > 0 ? Math.min(...numericData) : 0;
      case 'stdDev':
        return this._calculateStandardDeviation(numericData);
      case 'stdDevP':
        return this._calculateStandardDeviationP(numericData);
      case 'var':
        return this._calculateVariance(numericData);
      case 'varP':
        return this._calculateVarianceP(numericData);
      default:
        return numericData.reduce((sum, val) => sum + val, 0);
    }
  }

  /**
   * 根據值篩選資料（支援指定欄位）
   */
  private _filterDataByValues(rowValue: any, colValue: any, valueColumn?: string): any[] {
    const valueFields = this.config.fields.filter(f => f.type === 'value');
    if (valueFields.length === 0) return [];
    
    // 如果指定了值欄位，使用指定的；否則使用第一個值欄位
    const targetColumn = valueColumn || valueFields[0].sourceColumn;
    const valueColIndex = this._getColumnIndex(targetColumn);
    const filteredValues: any[] = [];
    
    for (let i = 1; i < this._sourceData.length; i++) {
      const row = this._sourceData[i];
      let matches = true;
      
      // 檢查行欄位
      const rowFields = this.config.fields.filter(f => f.type === 'row');
      for (const field of rowFields) {
        const colIndex = this._getColumnIndex(field.sourceColumn);
        if (colIndex >= 0) {
          if (rowValue !== null && row[colIndex] !== rowValue) {
            matches = false;
            break;
          }
        }
      }
      
      // 檢查列欄位
      const columnFields = this.config.fields.filter(f => f.type === 'column');
      for (const field of columnFields) {
        const colIndex = this._getColumnIndex(field.sourceColumn);
        if (colIndex >= 0) {
          if (colValue !== null && row[colIndex] !== colValue) {
            matches = false;
            break;
          }
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
    
    const columnFields = this.config.fields.filter(f => f.type === 'column');
    const valueFields = this.config.fields.filter(f => f.type === 'value');
    
    if (columnFields.length > 0 && valueFields.length > 0) {
      if (valueFields.length === 1) {
        // 單一值欄位：為每個列欄位計算總計
        for (const colField of columnFields) {
          const values = this._getColumnValues(colField.sourceColumn);
          const total = this._calculateFieldTotal(values, valueFields[0].function);
          totalRow.push(total);
        }
      } else {
        // 多個值欄位：為每個值欄位的每個列欄位計算總計
        for (const valueField of valueFields) {
          for (const colField of columnFields) {
            const values = this._getColumnValues(colField.sourceColumn);
            const total = this._calculateFieldTotal(values, valueField.function);
            totalRow.push(total);
          }
        }
      }
    } else if (valueFields.length > 0) {
      // 沒有列欄位：為每個值欄位計算總計
      for (const valueField of valueFields) {
        const values = this._getColumnValues(valueField.sourceColumn);
        const total = this._calculateFieldTotal(values, valueField.function);
        totalRow.push(total);
      }
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
  private _updateTargetWorksheet(ws?: Worksheet): void {
    // 清除現有內容
    this._clearTargetWorksheet(ws);
    
    // 寫入 Pivot Table 資料
    this._writePivotData(ws);
    
    // 應用樣式
    this._applyPivotStyles(ws);
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
    const fieldValues = this._fieldValues.get(field.sourceColumn);
    if (!fieldValues) return;
    
    const valueFields = this.config.fields.filter(f => f.type === 'value');
    const columnFields = this.config.fields.filter(f => f.type === 'column');
    
    // 為每個唯一值添加小計行
    for (const value of fieldValues) {
      const subtotalRow = [`${value} 小計`];
      
      if (valueFields.length === 1) {
        // 單一值欄位：為每個列欄位計算小計
        for (const colField of columnFields) {
          const filteredData = this._filterDataByValues(value, null, valueFields[0].sourceColumn);
          const subtotal = this._calculateValueByFunction(filteredData, valueFields[0].function);
          subtotalRow.push(subtotal);
        }
      } else if (valueFields.length > 1 && columnFields.length > 0) {
        // 多個值欄位：為每個值欄位的每個列欄位計算小計
        for (const valueField of valueFields) {
          for (const colField of columnFields) {
            const filteredData = this._filterDataByValues(value, null, valueField.sourceColumn);
            const subtotal = this._calculateValueByFunction(filteredData, valueField.function);
            subtotalRow.push(subtotal);
          }
        }
      } else if (valueFields.length > 1) {
        // 沒有列欄位：為每個值欄位計算小計
        for (const valueField of valueFields) {
          const filteredData = this._filterDataByValues(value, null, valueField.sourceColumn);
          const subtotal = this._calculateValueByFunction(filteredData, valueField.function);
          subtotalRow.push(subtotal);
        }
      }
      
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
    
    switch (functionType) {
      case 'sum':
        return numericValues.reduce((sum, val) => sum + val, 0);
      case 'count':
        return values.length;
      case 'countNums':
        return numericValues.length;
      case 'average':
        return numericValues.length > 0 ? numericValues.reduce((sum, val) => sum + val, 0) / numericValues.length : 0;
      case 'max':
        return numericValues.length > 0 ? Math.max(...numericValues) : 0;
      case 'min':
        return numericValues.length > 0 ? Math.min(...numericValues) : 0;
      case 'stdDev':
        return this._calculateStandardDeviation(numericValues);
      case 'stdDevP':
        return this._calculateStandardDeviationP(numericValues);
      case 'var':
        return this._calculateVariance(numericValues);
      case 'varP':
        return this._calculateVarianceP(numericValues);
      default:
        return numericValues.reduce((sum, val) => sum + val, 0);
    }
  }

  /**
   * 計算樣本標準差
   */
  private _calculateStandardDeviation(values: number[]): number {
    if (values.length < 2) return 0;
    const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
    const squaredDiffs = values.map(val => Math.pow(val - mean, 2));
    const variance = squaredDiffs.reduce((sum, val) => sum + val, 0) / (values.length - 1);
    return Math.sqrt(variance);
  }

  /**
   * 計算母體標準差
   */
  private _calculateStandardDeviationP(values: number[]): number {
    if (values.length === 0) return 0;
    const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
    const squaredDiffs = values.map(val => Math.pow(val - mean, 2));
    const variance = squaredDiffs.reduce((sum, val) => sum + val, 0) / values.length;
    return Math.sqrt(variance);
  }

  /**
   * 計算樣本變異數
   */
  private _calculateVariance(values: number[]): number {
    if (values.length < 2) return 0;
    const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
    const squaredDiffs = values.map(val => Math.pow(val - mean, 2));
    return squaredDiffs.reduce((sum, val) => sum + val, 0) / (values.length - 1);
  }

  /**
   * 計算母體變異數
   */
  private _calculateVarianceP(values: number[]): number {
    if (values.length === 0) return 0;
    const mean = values.reduce((sum, val) => sum + val, 0) / values.length;
    const squaredDiffs = values.map(val => Math.pow(val - mean, 2));
    return squaredDiffs.reduce((sum, val) => sum + val, 0) / values.length;
  }

  /**
   * 計算行總計
   */
  private _calculateRowTotals(): void {
    if (this._processedData.length === 0) return;
    
    const valueFields = this.config.fields.filter(f => f.type === 'value');
    const columnFields = this.config.fields.filter(f => f.type === 'column');
    
    // 為每行添加總計
    for (let i = 0; i < this._processedData.length; i++) {
      const row = this._processedData[i];
      if (row.length > 1) {
        let rowTotal = 0;
        
        if (valueFields.length === 1) {
          // 單一值欄位：計算所有數值欄位的總和
          const numericValues = row.slice(1).filter(v => typeof v === 'number');
          rowTotal = numericValues.reduce((sum, val) => sum + val, 0);
        } else if (valueFields.length > 1 && columnFields.length > 0) {
          // 多個值欄位：為每個值欄位計算總和
          const valueFieldCount = valueFields.length;
          const columnCount = columnFields.length;
          
          for (let vf = 0; vf < valueFieldCount; vf++) {
            let fieldTotal = 0;
            const startCol = 1 + vf * columnCount;
            const endCol = startCol + columnCount;
            
            for (let col = startCol; col < endCol && col < row.length; col++) {
              if (typeof row[col] === 'number') {
                fieldTotal += row[col];
              }
            }
            rowTotal += fieldTotal;
          }
        } else {
          // 沒有列欄位：計算所有數值欄位的總和
          const numericValues = row.slice(1).filter(v => typeof v === 'number');
          rowTotal = numericValues.reduce((sum, val) => sum + val, 0);
        }
        
        row.push(rowTotal);
      }
    }
  }

  /**
   * 計算列總計
   */
  private _calculateColumnTotals(): void {
    if (this._processedData.length === 0) return;
    
    const valueFields = this.config.fields.filter(f => f.type === 'value');
    const columnFields = this.config.fields.filter(f => f.type === 'column');
    
    // 計算每列的總計
    const maxCols = Math.max(...this._processedData.map(row => row.length));
    
    if (valueFields.length === 1) {
      // 單一值欄位：為每列計算總計
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
    } else if (valueFields.length > 1 && columnFields.length > 0) {
      // 多個值欄位：為每個值欄位的每列計算總計
      const columnCount = columnFields.length;
      
      for (let vf = 0; vf < valueFields.length; vf++) {
        for (let col = 0; col < columnCount; col++) {
          let colTotal = 0;
          const actualCol = 1 + vf * columnCount + col;
          
          for (let row = 0; row < this._processedData.length; row++) {
            const value = this._processedData[row][actualCol];
            if (typeof value === 'number') {
              colTotal += value;
            }
          }
          
          // 將列總計添加到最後一行
          if (this._processedData.length > 0) {
            const lastRow = this._processedData[this._processedData.length - 1];
            lastRow[actualCol] = colTotal;
          }
        }
      }
    }
  }

  /**
   * 計算總計
   */
  private _calculateGrandTotal(): void {
    if (this._processedData.length === 0) return;
    
    const valueFields = this.config.fields.filter(f => f.type === 'value');
    const columnFields = this.config.fields.filter(f => f.type === 'column');
    
    let grandTotal = 0;
    
    if (valueFields.length === 1) {
      // 單一值欄位：計算所有數值的總和
      for (const row of this._processedData) {
        for (const value of row) {
          if (typeof value === 'number') {
            grandTotal += value;
          }
        }
      }
    } else if (valueFields.length > 1 && columnFields.length > 0) {
      // 多個值欄位：為每個值欄位計算總和
      const columnCount = columnFields.length;
      
      for (let vf = 0; vf < valueFields.length; vf++) {
        let fieldTotal = 0;
        const startCol = 1 + vf * columnCount;
        const endCol = startCol + columnCount;
        
        for (const row of this._processedData) {
          for (let col = startCol; col < endCol && col < row.length; col++) {
            const value = row[col];
            if (typeof value === 'number') {
              fieldTotal += value;
            }
          }
        }
        grandTotal += fieldTotal;
      }
    } else {
      // 沒有列欄位：計算所有數值的總和
      for (const row of this._processedData) {
        for (const value of row) {
          if (typeof value === 'number') {
            grandTotal += value;
          }
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
    if (!ws) return;
    
    try {
      // 解析目標範圍
      if (this.config.targetRange) {
        const [startAddr, endAddr] = this.config.targetRange.split(':');
        const start = parseAddress(startAddr);
        const end = parseAddress(endAddr);
        
        // 清除指定範圍的儲存格
        for (let row = start.row; row <= end.row; row++) {
          for (let col = start.col; col <= end.col; col++) {
            const address = `${String.fromCharCode(65 + col - 1)}${row}`;
            try {
              ws.setCell(address, null);
            } catch (e) {
              // 忽略清除儲存格時的錯誤
            }
          }
        }
      } else {
        // 如果沒有指定目標範圍，清除前 100 行和前 26 列
        for (let row = 1; row <= 100; row++) {
          for (let col = 1; col <= 26; col++) {
            const address = `${String.fromCharCode(64 + col)}${row}`;
            try {
              ws.setCell(address, null);
            } catch (e) {
              // 忽略清除儲存格時的錯誤
            }
          }
        }
      }
    } catch (e) {
      console.warn(`清除目標工作表時發生錯誤: ${e.message}`);
    }
  }

  /**
   * 寫入 Pivot Table 資料到工作表
   */
  private _writePivotData(ws?: Worksheet): void {
    if (!ws) return;
    
    // 取得處理後的資料
    const data = this.getData();
    if (!data || data.length === 0) return;
    
    // 動態生成標題行
    const headers = this._generateDynamicHeaders();
    for (let col = 0; col < headers.length; col++) {
      const address = `${String.fromCharCode(65 + col)}1`;
      ws.setCell(address, headers[col], { font: { bold: true } });
    }
    
    // 寫入資料行
    for (let row = 0; row < data.length; row++) {
      const rowData = data[row];
      for (let col = 0; col < rowData.length; col++) {
        const address = `${String.fromCharCode(65 + col)}${row + 2}`;
        const value = rowData[col];
        
        // 根據資料類型設定樣式
        let options = {};
        if (typeof value === 'number') {
          options = { 
            font: { bold: row === 0 || col === rowData.length - 1 }, // 第一行和最後一列加粗
            alignment: { horizontal: 'right' } // 數字右對齊
          };
        } else if (typeof value === 'string') {
          options = { 
            font: { bold: row === 0 }, // 第一行加粗
            alignment: { horizontal: 'left' } // 文字左對齊
          };
        }
        
        ws.setCell(address, value, options);
      }
    }
    
    // 動態設定欄寬
    this._setDynamicColumnWidths(ws, headers.length);
  }

  /**
   * 動態生成標題行
   */
  private _generateDynamicHeaders(): string[] {
    const headers: string[] = [];
    
    // 根據欄位配置動態生成標題
    for (const field of this.config.fields) {
      if (field.type === 'row' || field.type === 'column') {
        headers.push(field.customName || field.name);
      }
    }
    
    // 添加值欄位標題
    for (const field of this.config.fields) {
      if (field.type === 'value') {
        const functionName = field.function || 'sum';
        const displayName = field.customName || `${field.name} (${functionName})`;
        headers.push(displayName);
      }
    }
    
    // 如果啟用小計，添加小計欄位
    if (this.config.showRowSubtotals) {
      headers.push('小計');
    }
    
    // 如果啟用總計，添加總計欄位
    if (this.config.showGrandTotals) {
      headers.push('總計');
    }
    
    return headers;
  }

  /**
   * 動態設定欄寬
   */
  private _setDynamicColumnWidths(ws: Worksheet, columnCount: number): void {
    for (let col = 0; col < columnCount; col++) {
      const columnLetter = String.fromCharCode(65 + col);
      let width = 12; // 預設寬度
      
      // 根據欄位類型調整寬度
      if (col < this.config.fields.length) {
        const field = this.config.fields[col];
        if (field.type === 'row' || field.type === 'column') {
          width = Math.max(10, (field.customName || field.name).length + 2);
        } else if (field.type === 'value') {
          width = 15; // 數值欄位稍寬
        }
      }
      
      ws.setColumnWidth(columnLetter, width);
    }
  }

  /**
   * 應用 Pivot Table 樣式
   */
  private _applyPivotStyles(ws?: Worksheet): void {
    if (!ws) return;
    
    // 取得資料範圍
    const data = this.getData();
    if (!data || data.length === 0) return;
    
    const startRow = 1;
    const endRow = data.length + 1;
    const startCol = 1;
    const endCol = this._generateDynamicHeaders().length; // 動態計算欄數
    
    // 為標題行添加邊框和背景色
    for (let col = startCol; col <= endCol; col++) {
      const address = `${String.fromCharCode(64 + col)}${startRow}`;
      const cell = ws.getCell(address);
      if (cell) {
        cell.options = {
          ...cell.options,
          fill: { type: 'pattern', patternType: 'solid', fgColor: 'E0E0E0' },
          border: {
            top: { style: 'thin' },
            bottom: { style: 'thin' },
            left: { style: 'thin' },
            right: { style: 'thin' }
          }
        };
      }
    }
    
    // 為資料行添加邊框
    for (let row = startRow + 1; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const address = `${String.fromCharCode(64 + col)}${row}`;
        const cell = ws.getCell(address);
        if (cell) {
          cell.options = {
            ...cell.options,
            border: {
              top: { style: 'thin' },
              bottom: { style: 'thin' },
              left: { style: 'thin' },
              right: { style: 'thin' }
            }
          };
        }
      }
    }
    
    // 為小計行添加特殊樣式
    if (this.config.showRowSubtotals) {
      for (let col = startCol; col <= endCol; col++) {
        const address = `${String.fromCharCode(64 + col)}${endRow}`;
        const cell = ws.getCell(address);
        if (cell) {
          cell.options = {
            ...cell.options,
            fill: { type: 'pattern', patternType: 'solid', fgColor: 'F0F0F0' },
            font: { bold: true }
          };
        }
      }
    }
  }

  /**
   * 生成模擬資料（用於測試）
   */
  private _generateMockData(startRow: number, endRow: number, startCol: number, endCol: number): any[][] {
    const data: any[][] = [];
    
    // 動態生成標題行
    const headers = this._generateSourceHeaders();
    data.push(headers);
    
    // 動態生成資料行
    const rowCount = Math.max(100, endRow - startRow);
    for (let i = 0; i < rowCount; i++) {
      const row = this._generateMockDataRow(headers, i);
      data.push(row);
    }
    
    return data;
  }

  /**
   * 動態生成來源資料標題
   */
  private _generateSourceHeaders(): string[] {
    const headers: string[] = [];
    
    // 根據欄位配置動態生成標題
    for (const field of this.config.fields) {
      if (field.sourceColumn && !headers.includes(field.sourceColumn)) {
        headers.push(field.sourceColumn);
      }
    }
    
    // 如果沒有找到任何來源欄位，使用預設欄位
    if (headers.length === 0) {
      headers.push('欄位1', '欄位2', '欄位3', '數值');
    }
    
    return headers;
  }

  /**
   * 動態生成模擬資料行
   */
  private _generateMockDataRow(headers: string[], rowIndex: number): any[] {
    const row: any[] = [];
    
    for (const header of headers) {
      // 根據欄位名稱生成適當的模擬資料
      if (header.toLowerCase().includes('產品') || header.toLowerCase().includes('product')) {
        const products = ['筆記型電腦', '平板電腦', '智慧型手機', '耳機', '鍵盤', '滑鼠'];
        row.push(products[rowIndex % products.length]);
      } else if (header.toLowerCase().includes('地區') || header.toLowerCase().includes('region')) {
        const regions = ['北區', '中區', '南區', '東區', '西區'];
        row.push(regions[rowIndex % regions.length]);
      } else if (header.toLowerCase().includes('月份') || header.toLowerCase().includes('month')) {
        const months = ['1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月', '11月', '12月'];
        row.push(months[rowIndex % months.length]);
      } else if (header.toLowerCase().includes('日期') || header.toLowerCase().includes('date')) {
        const date = new Date(2024, rowIndex % 12, (rowIndex % 28) + 1);
        row.push(date);
      } else if (header.toLowerCase().includes('時間') || header.toLowerCase().includes('time')) {
        const time = new Date(2024, 0, 1, rowIndex % 24, rowIndex % 60);
        row.push(time);
      } else if (header.toLowerCase().includes('狀態') || header.toLowerCase().includes('status')) {
        const statuses = ['啟用', '停用', '待審核', '已核准'];
        row.push(statuses[rowIndex % statuses.length]);
      } else if (header.toLowerCase().includes('類別') || header.toLowerCase().includes('category')) {
        const categories = ['A類', 'B類', 'C類', 'D類'];
        row.push(categories[rowIndex % categories.length]);
      } else if (header.toLowerCase().includes('是否') || header.toLowerCase().includes('is') || 
                 header.toLowerCase().includes('有') || header.toLowerCase().includes('has')) {
        // 布林值欄位
        row.push(rowIndex % 2 === 0);
      } else if (header.toLowerCase().includes('銷售') || header.toLowerCase().includes('sales') || 
                 header.toLowerCase().includes('數量') || header.toLowerCase().includes('quantity') ||
                 header.toLowerCase().includes('金額') || header.toLowerCase().includes('amount') ||
                 header.toLowerCase().includes('數值') || header.toLowerCase().includes('value')) {
        // 數值欄位
        row.push(Math.floor(Math.random() * 10000) + 1000);
      } else if (header.toLowerCase().includes('百分比') || header.toLowerCase().includes('percentage')) {
        // 百分比欄位
        row.push(Math.round((Math.random() * 100) * 100) / 100);
      } else if (header.toLowerCase().includes('評分') || header.toLowerCase().includes('rating')) {
        // 評分欄位
        row.push(Math.round((Math.random() * 5 + 1) * 10) / 10);
      } else {
        // 預設為字串欄位
        row.push(`值${rowIndex + 1}`);
      }
    }
    
    return row;
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
    // 嘗試從來源範圍中提取工作表名稱
    if (this.config.sourceRange && this.config.sourceRange.includes('!')) {
      const parts = this.config.sourceRange.split('!');
      if (parts.length > 0) {
        // 移除單引號（如果有的話）
        let sheetName = parts[0];
        if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
          sheetName = sheetName.slice(1, -1);
        }
        return sheetName;
      }
    }
    
    // 如果無法從來源範圍提取，嘗試從工作簿中獲取第一個工作表名稱
    try {
      const worksheets = this._workbook.getWorksheets();
      if (worksheets.length > 0) {
        return worksheets[0].name;
      }
    } catch (e) {
      // 忽略錯誤
    }
    
    // 預設值
    return 'Sheet1';
  }

  /**
   * 生成快取欄位 XML
   */
  private _generateCacheFieldXml(field: PivotField): string {
    const values = Array.from(this._fieldValues.get(field.name) || []);
    
    // 動態檢測資料類型
    const hasStrings = values.some(v => typeof v === 'string');
    const hasNumbers = values.some(v => typeof v === 'number');
    const hasDates = values.some(v => v instanceof Date);
    const hasBooleans = values.some(v => typeof v === 'boolean');
    
    // 設定 SQL 類型
    let sqlType = 0; // 預設為一般類型
    if (hasNumbers && !hasStrings && !hasDates && !hasBooleans) {
      sqlType = 2; // 數值類型
    } else if (hasDates && !hasStrings && !hasNumbers && !hasBooleans) {
      sqlType = 3; // 日期類型
    } else if (hasBooleans && !hasStrings && !hasNumbers && !hasDates) {
      sqlType = 4; // 布林類型
    } else if (hasStrings) {
      sqlType = 1; // 字串類型
    }
    
    let xml = `
    <cacheField name="${field.name}" numFmtId="0" formula="0" sqlType="${sqlType}" hierarchy="0" level="0" databaseField="0" mappingCount="0" olap="0">`;
    
    // 欄位項目
    if (values.length > 0) {
      // 動態設定 contains 屬性
      const containsSemiMixedTypes = hasStrings && hasNumbers ? 1 : 0;
      const containsString = hasStrings ? 1 : 0;
      const containsNumber = hasNumbers ? 1 : 0;
      const containsInteger = hasNumbers ? 1 : 0;
      const containsNonDate = hasStrings || hasNumbers || hasBooleans ? 1 : 0;
      const containsDate = hasDates ? 1 : 0;
      const containsBlank = 0; // 暫時設為 0
      const mixedTypes = (hasStrings ? 1 : 0) + (hasNumbers ? 1 : 0) + (hasDates ? 1 : 0) + (hasBooleans ? 1 : 0);
      
      // 計算數值範圍
      const numericValues = values.filter(v => typeof v === 'number');
      const minNumber = numericValues.length > 0 ? Math.min(...numericValues) : 0;
      const maxNumber = numericValues.length > 0 ? Math.max(...numericValues) : 0;
      const minInteger = numericValues.length > 0 ? Math.floor(Math.min(...numericValues)) : 0;
      const maxInteger = numericValues.length > 0 ? Math.floor(Math.max(...numericValues)) : 0;
      
      // 計算日期範圍
      const dateValues = values.filter(v => v instanceof Date);
      const minDate = dateValues.length > 0 ? dateValues[0].toISOString() : '1900-01-01T00:00:00';
      const maxDate = dateValues.length > 0 ? dateValues[dateValues.length - 1].toISOString() : '9999-12-31T23:59:59';
      
      xml += `
      <sharedItems count="${values.length}" containsSemiMixedTypes="${containsSemiMixedTypes}" containsString="${containsString}" containsNumber="${containsNumber}" containsInteger="${containsInteger}" containsNonDate="${containsNonDate}" containsDate="${containsDate}" containsBlank="${containsBlank}" mixedTypes="${mixedTypes}" minDate="${minDate}" maxDate="${maxDate}" minNumber="${minNumber}" maxNumber="${maxNumber}" minInteger="${minInteger}" maxInteger="${maxInteger}" containsLocal="0" containsRemote="0" remote="0">`;
      
      for (const value of values) {
        if (typeof value === 'string') {
          xml += `
        <s v="${this._escapeXmlValue(value)}"/>`;
        } else if (typeof value === 'number') {
          xml += `
        <n v="${value}"/>`;
        } else if (value instanceof Date) {
          xml += `
        <d v="${value.toISOString()}"/>`;
        } else if (typeof value === 'boolean') {
          xml += `
        <b v="${value ? '1' : '0'}"/>`;
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
      } else if (value instanceof Date) {
        xml += `
        <d v="${value.toISOString()}"/>`;
      } else if (value === null || value === undefined) {
        xml += `
        <m/>`; // 空值
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
    // 根據欄位類型設定軸屬性
    let axis = 'axisRow';
    if (field.type === 'column') {
      axis = 'axisCol';
    } else if (field.type === 'value') {
      axis = 'axisValues';
    } else if (field.type === 'filter') {
      axis = 'axisPage';
    }
    
    let xml = `
    <pivotField axis="${axis}" showAll="0"`;
    
    // 為值欄位添加特殊屬性
    if (field.type === 'value') {
      xml += ` dataField="1"`;
      if (field.function) {
        xml += ` subtotal="${field.function}"`;
      }
    }
    
    xml += `>`;
    
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
