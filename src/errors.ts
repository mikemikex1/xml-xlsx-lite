/**
 * xml-xlsx-lite 錯誤處理系統
 */

export class InvalidAddressError extends Error {
  constructor(message: string, public details: { sample?: string }) {
    super(message);
    this.name = 'InvalidAddressError';
  }
}

export class UnsupportedTypeError extends Error {
  constructor(message: string, public details: { suggest?: string }) {
    super(message);
    this.name = 'UnsupportedTypeError';
  }
}

export class CorruptedFileError extends Error {
  constructor(message: string, public details: { file?: string; expected?: string }) {
    super(message);
    this.name = 'CorruptedFileError';
  }
}

export class UnsupportedFeatureWarning extends Error {
  constructor(message: string, public details: { feature?: string; alternative?: string }) {
    super(message);
    this.name = 'UnsupportedFeatureWarning';
  }
}

export class ValidationError extends Error {
  constructor(message: string, public details: { field?: string; value?: any; expected?: any }) {
    super(message);
    this.name = 'ValidationError';
  }
}

export class PerformanceWarning extends Error {
  constructor(message: string, public details: { suggestion?: string; threshold?: number }) {
    super(message);
    this.name = 'PerformanceWarning';
  }
}

/**
 * 錯誤代碼定義
 */
export enum ErrorCodes {
  INVALID_ADDRESS = 'INVALID_ADDRESS',
  UNSUPPORTED_TYPE = 'UNSUPPORTED_TYPE',
  CORRUPTED_FILE = 'CORRUPTED_FILE',
  UNSUPPORTED_FEATURE = 'UNSUPPORTED_FEATURE',
  VALIDATION_ERROR = 'VALIDATION_ERROR',
  PERFORMANCE_WARNING = 'PERFORMANCE_WARNING'
}

/**
 * 錯誤訊息模板
 */
export const ErrorMessages = {
  [ErrorCodes.INVALID_ADDRESS]: (address: string) => 
    `Invalid cell address: ${address}. Expected format: A1, B2, etc.`,
  
  [ErrorCodes.UNSUPPORTED_TYPE]: (type: string, value: any) => 
    `Unsupported data type: ${type} for value ${value}. Please convert to supported type.`,
  
  [ErrorCodes.CORRUPTED_FILE]: (file: string) => 
    `Corrupted or invalid file: ${file}. File may be damaged or in unsupported format.`,
  
  [ErrorCodes.UNSUPPORTED_FEATURE]: (feature: string) => 
    `Feature not yet implemented: ${feature}. This will be available in future versions.`,
  
  [ErrorCodes.VALIDATION_ERROR]: (field: string, value: any, expected: any) => 
    `Validation failed for ${field}: got ${value}, expected ${expected}.`,
  
  [ErrorCodes.PERFORMANCE_WARNING]: (message: string) => 
    `Performance warning: ${message}. Consider optimizing your data or using streaming mode.`
};

/**
 * 創建標準化錯誤
 */
export function createError(
  code: ErrorCodes, 
  details: Record<string, any> = {}, 
  customMessage?: string
): Error {
  let message: string;
  
  try {
    switch (code) {
      case ErrorCodes.INVALID_ADDRESS:
        message = customMessage || ErrorMessages[code](details.sample || '');
        break;
      case ErrorCodes.UNSUPPORTED_TYPE:
        message = customMessage || ErrorMessages[code](details.type || '', details.value);
        break;
      case ErrorCodes.CORRUPTED_FILE:
        message = customMessage || ErrorMessages[code](details.file || '');
        break;
      case ErrorCodes.UNSUPPORTED_FEATURE:
        message = customMessage || ErrorMessages[code](details.feature || '');
        break;
      case ErrorCodes.VALIDATION_ERROR:
        message = customMessage || ErrorMessages[code](details.field || '', details.value, details.expected);
        break;
      case ErrorCodes.PERFORMANCE_WARNING:
        message = customMessage || ErrorMessages[code](details.suggestion || '');
        break;
      default:
        message = customMessage || `Error occurred: ${code}`;
    }
  } catch (e) {
    message = customMessage || `Error occurred: ${code}`;
  }
  
  switch (code) {
    case ErrorCodes.INVALID_ADDRESS:
      return new InvalidAddressError(message, details);
    case ErrorCodes.UNSUPPORTED_TYPE:
      return new UnsupportedTypeError(message, details);
    case ErrorCodes.CORRUPTED_FILE:
      return new CorruptedFileError(message, details);
    case ErrorCodes.UNSUPPORTED_FEATURE:
      return new UnsupportedFeatureWarning(message, details);
    case ErrorCodes.VALIDATION_ERROR:
      return new ValidationError(message, details);
    case ErrorCodes.PERFORMANCE_WARNING:
      return new PerformanceWarning(message, details);
    default:
      return new Error(message);
  }
}
