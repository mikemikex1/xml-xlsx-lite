import { WorksheetProtectionOptions, WorkbookProtectionOptions } from './types';
import { hashPassword, verifyPassword } from './utils';

/**
 * 工作表保護實現
 */
export class WorksheetProtection {
  private _isProtected = false;
  private _passwordHash: string | null = null;
  private _options: WorksheetProtectionOptions = {};

  constructor() {
    this._options = {
      selectLockedCells: false,
      selectUnlockedCells: true,
      formatCells: false,
      formatColumns: false,
      formatRows: false,
      insertColumns: false,
      insertRows: false,
      insertHyperlinks: false,
      deleteColumns: false,
      deleteRows: false,
      sort: false,
      autoFilter: false,
      pivotTables: false,
      objects: false,
      scenarios: false
    };
  }

  /**
   * 保護工作表
   */
  protect(password?: string, options?: Partial<WorksheetProtectionOptions>): void {
    this._isProtected = true;
    if (password) {
      this._passwordHash = hashPassword(password);
    }
    if (options) {
      this._options = { ...this._options, ...options };
    }
  }

  /**
   * 解除工作表保護
   */
  unprotect(password?: string): void {
    if (this._isProtected && this._passwordHash && password) {
      if (!verifyPassword(password, this._passwordHash)) {
        throw new Error('Incorrect password');
      }
    }
    this._isProtected = false;
    this._passwordHash = null;
  }

  /**
   * 檢查工作表是否受保護
   */
  isProtected(): boolean {
    return this._isProtected;
  }

  /**
   * 取得保護選項
   */
  getProtectionOptions(): WorksheetProtectionOptions | null {
    if (!this._isProtected) return null;
    return { ...this._options };
  }

  /**
   * 檢查操作是否被允許
   */
  isOperationAllowed(operation: keyof WorksheetProtectionOptions): boolean {
    if (!this._isProtected) return true;
    return this._options[operation] || false;
  }

  /**
   * 驗證密碼
   */
  validatePassword(password: string): boolean {
    if (!this._passwordHash) return true;
    return verifyPassword(password, this._passwordHash);
  }
}

/**
 * 工作簿保護實現
 */
export class WorkbookProtection {
  private _isProtected = false;
  private _passwordHash: string | null = null;
  private _options: WorkbookProtectionOptions = {};

  constructor() {
    this._options = {
      structure: false,
      windows: false
    };
  }

  /**
   * 保護工作簿
   */
  protect(password?: string, options?: Partial<WorkbookProtectionOptions>): void {
    this._isProtected = true;
    if (password) {
      this._passwordHash = hashPassword(password);
    }
    if (options) {
      this._options = { ...this._options, ...options };
    }
  }

  /**
   * 解除工作簿保護
   */
  unprotect(password?: string): void {
    if (this._isProtected && this._passwordHash && password) {
      if (!verifyPassword(password, this._passwordHash)) {
        throw new Error('Incorrect password');
      }
    }
    this._isProtected = false;
    this._passwordHash = null;
  }

  /**
   * 檢查工作簿是否受保護
   */
  isProtected(): boolean {
    return this._isProtected;
  }

  /**
   * 取得保護選項
   */
  getProtectionOptions(): WorkbookProtectionOptions | null {
    if (!this._isProtected) return null;
    return { ...this._options };
  }

  /**
   * 檢查操作是否被允許
   */
  isOperationAllowed(operation: keyof Omit<WorkbookProtectionOptions, 'password'>): boolean {
    if (!this._isProtected) return true;
    return this._options[operation] || false;
  }

  /**
   * 驗證密碼
   */
  validatePassword(password: string): boolean {
    if (!this._passwordHash) return true;
    return verifyPassword(password, this._passwordHash);
  }
}
