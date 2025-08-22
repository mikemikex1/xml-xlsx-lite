/**
 * API 相容性層
 * 提供薄包裝方法，避免使用者誤用
 */

import { addPivotToWorkbookBuffer } from '../pivot-builder';
import { WorkbookImpl as Workbook } from '../workbook';
import { createMethodNotImplementedError, createFileOperationError } from '../errors';

// Node.js 環境檢查
let fs: any;
let Buffer: any;
try {
  // @ts-ignore
  fs = require('fs');
  // @ts-ignore
  Buffer = require('buffer').Buffer;
} catch {
  // 瀏覽器環境
  fs = null;
  Buffer = null;
}

/**
 * 將工作簿寫入檔案（薄包裝）
 * 內部使用 writeBuffer() + fs.writeFileSync
 */
export async function writeFile(
  this: Workbook,
  filePath: string
): Promise<string> {
  if (!fs || !Buffer) {
    throw createFileOperationError(
      'writeFile',
      filePath,
      'File system operations are not available in this environment'
    );
  }
  
  try {
    const buffer = await this.writeBuffer();
    fs.writeFileSync(filePath, new Uint8Array(buffer));
    return filePath;
  } catch (error) {
    throw createFileOperationError(
      'writeFile',
      filePath,
      error instanceof Error ? error.message : String(error)
    );
  }
}

/**
 * 將工作簿寫入檔案並插入樞紐分析表（薄包裝）
 * 內部使用 writeBuffer() + addPivotToWorkbookBuffer + fs.writeFileSync
 */
export async function writeFileWithPivotTables(
  this: Workbook,
  filePath: string,
  pivotOptions: any
): Promise<string> {
  if (!fs || !Buffer) {
    throw createFileOperationError(
      'writeFileWithPivotTables',
      filePath,
      'File system operations are not available in this environment'
    );
  }
  
  try {
    let buffer = await this.writeBuffer();
    
    // 插入樞紐分析表
    if (pivotOptions) {
      // 轉換 ArrayBuffer 為 Buffer
      const nodeBuffer = Buffer.from(buffer);
      const newBuffer = await addPivotToWorkbookBuffer(nodeBuffer, pivotOptions);
      // 轉換回 ArrayBuffer
      buffer = newBuffer.buffer.slice(newBuffer.byteOffset, newBuffer.byteOffset + newBuffer.byteLength) as ArrayBuffer;
    }
    
    // 寫入檔案
    fs.writeFileSync(filePath, Buffer.from(buffer));
    return filePath;
  } catch (error) {
    throw createFileOperationError(
      'writeFileWithPivotTables',
      filePath,
      error instanceof Error ? error.message : String(error)
    );
  }
}

/**
 * 將工作簿寫入檔案並插入多個樞紐分析表（薄包裝）
 */
export async function writeFileWithMultiplePivots(
  this: Workbook,
  filePath: string,
  pivotOptionsArray: any[]
): Promise<string> {
  if (!fs || !Buffer) {
    throw createFileOperationError(
      'writeFileWithMultiplePivots',
      filePath,
      'File system operations are not available in this environment'
    );
  }
  
  try {
    let buffer = await this.writeBuffer();
    
    // 依序插入多個樞紐分析表
    for (const pivotOptions of pivotOptionsArray) {
      // 轉換 ArrayBuffer 為 Buffer
      const nodeBuffer = Buffer.from(buffer);
      const newBuffer = await addPivotToWorkbookBuffer(nodeBuffer, pivotOptions);
      // 轉換回 ArrayBuffer
      buffer = newBuffer.buffer.slice(newBuffer.byteOffset, newBuffer.byteOffset + newBuffer.byteLength) as ArrayBuffer;
    }
    
    // 寫入檔案
    fs.writeFileSync(filePath, Buffer.from(buffer));
    return filePath;
  } catch (error) {
    throw createFileOperationError(
      'writeFileWithMultiplePivots',
      filePath,
      error instanceof Error ? error.message : String(error)
    );
  }
}

/**
 * 攔截舊的 writeFile 方法，提供明確的錯誤訊息
 */
export function interceptOldWriteFile(workbook: Workbook): void {
  // 攔截舊的 writeFile 方法
  (workbook as any).writeFile = function() {
    throw createMethodNotImplementedError(
      'writeFile',
      'Use writeBuffer() + fs.writeFileSync instead.',
      'Example: const buf = await wb.writeBuffer(); fs.writeFileSync("out.xlsx", new Uint8Array(buf));'
    );
  };
}

/**
 * 將相容性方法綁定到 Workbook 原型
 */
export function bindCompatibilityMethods(): void {
  // 綁定新方法
  (Workbook as any).prototype.writeFile = writeFile;
  (Workbook as any).prototype.writeFileWithPivotTables = writeFileWithPivotTables;
  (Workbook as any).prototype.writeFileWithMultiplePivots = writeFileWithMultiplePivots;
}
