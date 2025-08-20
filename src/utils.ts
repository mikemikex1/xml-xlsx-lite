// 工具函數

// 欄位轉換工具
export const COL_A_CODE = "A".charCodeAt(0);
export const EXCEL_EPOCH = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30 (Excel's 1900 date system, including the 1900 leap-year bug)

/**
 * 將欄位名稱轉換為數字
 * e.g., A -> 1, Z -> 26, AA -> 27
 */
export function colToNumber(col: string): number {
  let n = 0;
  for (let i = 0; i < col.length; i++) {
    n = n * 26 + (col.charCodeAt(i) - COL_A_CODE + 1);
  }
  return n;
}

/**
 * 將數字轉換為欄位名稱
 */
export function numberToCol(n: number): string {
  let col = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    col = String.fromCharCode(COL_A_CODE + rem) + col;
    n = Math.floor((n - 1) / 26);
  }
  return col;
}

/**
 * 解析儲存格地址
 * "B12" -> { col: 2, row: 12 }
 */
export function parseAddress(addr: string): { col: number; row: number } {
  const m = /^([A-Z]+)(\d+)$/.exec(addr.toUpperCase());
  if (!m) throw new Error(`Invalid cell address: ${addr}`);
  return { col: colToNumber(m[1]), row: parseInt(m[2], 10) };
}

/**
 * 從行列號生成儲存格地址
 */
export function addrFromRC(row: number, col: number): string {
  return `${numberToCol(col)}${row}`;
}

/**
 * 檢查是否為日期類型
 */
export function isDate(val: any): val is Date {
  return val instanceof Date;
}

/**
 * 將 JavaScript Date 轉換為 Excel 序列號 (1900 系統)
 */
export function excelSerialFromDate(d: Date): number {
  const msPerDay = 24 * 60 * 60 * 1000;
  const diff = (Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - EXCEL_EPOCH.getTime()) / msPerDay;
  return diff;
}

/**
 * 取得儲存格類型
 */
export function getCellType(value: any): 'n' | 's' | 'b' | 'd' | null {
  if (value === null || value === undefined) return null;
  if (typeof value === "number") return "n";
  if (typeof value === "boolean") return "b";
  if (isDate(value)) return "n"; // we will write as serial number for now
  return "s"; // default: string
}

/**
 * XML 轉義工具
 */
export function escapeXmlText(str: any): string {
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

export function escapeXmlAttr(str: any): string {
  return escapeXmlText(str);
}

/**
 * 密碼雜湊工具（簡單實現）
 */
export function hashPassword(password: string): string {
  // 這裡使用簡單的雜湊，實際應用中應使用更安全的演算法
  let hash = 0;
  for (let i = 0; i < password.length; i++) {
    const char = password.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32bit integer
  }
  return Math.abs(hash).toString(16);
}

/**
 * 驗證密碼
 */
export function verifyPassword(password: string, hash: string): boolean {
  return hashPassword(password) === hash;
}

/**
 * 生成唯一 ID
 */
export function generateUniqueId(): string {
  return Date.now().toString(36) + Math.random().toString(36).substr(2);
}
