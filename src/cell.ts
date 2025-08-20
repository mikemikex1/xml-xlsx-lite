import { Cell, CellOptions } from './types';

/**
 * 儲存格模型實現
 */
export class CellModel implements Cell {
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

  /**
   * 設定儲存格值
   */
  setValue(value: number | string | boolean | Date | null): void {
    this.value = value;
  }

  /**
   * 設定儲存格選項
   */
  setOptions(options: CellOptions): void {
    this.options = { ...this.options, ...options };
  }

  /**
   * 清除儲存格內容
   */
  clear(): void {
    this.value = null;
    this.type = null;
    this.options = {};
  }

  /**
   * 檢查儲存格是否為空
   */
  isEmpty(): boolean {
    return this.value === null || this.value === undefined || this.value === '';
  }

  /**
   * 取得儲存格顯示值
   */
  getDisplayValue(): string {
    if (this.value === null || this.value === undefined) return '';
    if (this.value instanceof Date) return this.value.toLocaleDateString();
    return String(this.value);
  }
}
