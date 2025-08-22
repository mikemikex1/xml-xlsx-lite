/**
 * xml-xlsx-lite 自訂錯誤類型
 * 提供統一的錯誤處理和建議
 */

export class XlsxLiteError extends Error {
  constructor(
    message: string,
    public code: 'METHOD_NOT_IMPLEMENTED' | 'INVALID_PIVOT_SPEC' | 'UNSUPPORTED_OPERATION' | 'INVALID_DATA' | 'FILE_OPERATION_FAILED',
    public suggestion?: string
  ) {
    super(message);
    this.name = 'XlsxLiteError';
  }
}

/**
 * 創建方法未實作錯誤
 */
export function createMethodNotImplementedError(
  methodName: string,
  alternative: string,
  example?: string
): XlsxLiteError {
  return new XlsxLiteError(
    `Method '${methodName}' is not implemented. ${alternative}`,
    'METHOD_NOT_IMPLEMENTED',
    example
  );
}

/**
 * 創建樞紐分析表規格錯誤
 */
export function createInvalidPivotSpecError(
  field: string,
  value: any,
  expected: string
): XlsxLiteError {
  return new XlsxLiteError(
    `Invalid pivot table specification: ${field} "${value}" is not valid. Expected: ${expected}`,
    'INVALID_PIVOT_SPEC',
    `Please check the ${field} configuration and ensure it matches the expected format.`
  );
}

/**
 * 創建不支援操作錯誤
 */
export function createUnsupportedOperationError(
  operation: string,
  reason: string,
  alternative?: string
): XlsxLiteError {
  return new XlsxLiteError(
    `Operation '${operation}' is not supported: ${reason}`,
    'UNSUPPORTED_OPERATION',
    alternative
  );
}

/**
 * 創建無效資料錯誤
 */
export function createInvalidDataError(
  field: string,
  value: any,
  rule: string
): XlsxLiteError {
  return new XlsxLiteError(
    `Invalid data: ${field} "${value}" violates rule: ${rule}`,
    'INVALID_DATA',
    `Please validate your data and ensure it meets the requirements.`
  );
}

/**
 * 創建檔案操作錯誤
 */
export function createFileOperationError(
  operation: string,
  filePath: string,
  details: string
): XlsxLiteError {
  return new XlsxLiteError(
    `File operation '${operation}' failed for '${filePath}': ${details}`,
    'FILE_OPERATION_FAILED',
    `Please check file permissions, disk space, and ensure the file path is valid.`
  );
}
