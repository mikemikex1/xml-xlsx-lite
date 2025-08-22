/**
 * xml-xlsx-lite 效能優化器
 * 實現 sharedStrings 自動切換和大檔案處理優化
 */

import { XlsxLiteError } from './errors';

/**
 * 效能優化配置
 */
export interface PerformanceConfig {
  /** 啟用 sharedStrings 的閾值（字串數量） */
  sharedStringsThreshold: number;
  /** 啟用 sharedStrings 的重複率閾值（百分比） */
  repetitionRateThreshold: number;
  /** 大檔案處理的儲存格數量閾值 */
  largeFileThreshold: number;
  /** 啟用串流處理的檔案大小閾值（MB） */
  streamingThreshold: number;
  /** 快取大小限制（MB） */
  cacheSizeLimit: number;
  /** 是否啟用記憶體優化 */
  memoryOptimization: boolean;
}

/**
 * 效能統計資訊
 */
export interface PerformanceStats {
  /** 總儲存格數量 */
  totalCells: number;
  /** 字串儲存格數量 */
  stringCells: number;
  /** 唯一字串數量 */
  uniqueStrings: number;
  /** 字串重複率 */
  repetitionRate: number;
  /** 記憶體使用量（MB） */
  memoryUsage: number;
  /** 處理時間（毫秒） */
  processingTime: number;
  /** 建議的優化策略 */
  recommendedStrategy: string[];
}

/**
 * 效能優化器
 */
export class PerformanceOptimizer {
  private config: PerformanceConfig;
  private stats: PerformanceStats;

  constructor(config?: Partial<PerformanceConfig>) {
    this.config = {
      sharedStringsThreshold: 100,
      repetitionRateThreshold: 30,
      largeFileThreshold: 10000,
      streamingThreshold: 10,
      cacheSizeLimit: 100,
      memoryOptimization: true,
      ...config
    };

    this.stats = {
      totalCells: 0,
      stringCells: 0,
      uniqueStrings: 0,
      repetitionRate: 0,
      memoryUsage: 0,
      processingTime: 0,
      recommendedStrategy: []
    };
  }

  /**
   * 分析工作表效能
   */
  analyzeWorksheet(worksheet: any): PerformanceStats {
    const startTime = Date.now();
    
    // 收集統計資訊
    this.collectStats(worksheet);
    
    // 計算重複率
    this.calculateRepetitionRate();
    
    // 生成建議策略
    this.generateRecommendations();
    
    // 計算處理時間
    this.stats.processingTime = Date.now() - startTime;
    
    return { ...this.stats };
  }

  /**
   * 收集統計資訊
   */
  private collectStats(worksheet: any): void {
    this.stats.totalCells = 0;
    this.stats.stringCells = 0;
    this.stats.uniqueStrings = 0;
    
    const stringSet = new Set<string>();
    
    // 遍歷所有儲存格
    if (worksheet._cells) {
      for (const [address, cell] of worksheet._cells) {
        this.stats.totalCells++;
        
        if (cell.value && typeof cell.value === 'string') {
          this.stats.stringCells++;
          stringSet.add(cell.value);
        }
      }
    }
    
    this.stats.uniqueStrings = stringSet.size;
  }

  /**
   * 計算字串重複率
   */
  private calculateRepetitionRate(): void {
    if (this.stats.stringCells === 0) {
      this.stats.repetitionRate = 0;
      return;
    }
    
    this.stats.repetitionRate = ((this.stats.stringCells - this.stats.uniqueStrings) / this.stats.stringCells) * 100;
  }

  /**
   * 生成優化建議
   */
  private generateRecommendations(): void {
    this.stats.recommendedStrategy = [];
    
    // 檢查是否需要啟用 sharedStrings
    if (this.shouldUseSharedStrings()) {
      this.stats.recommendedStrategy.push('啟用 sharedStrings 以減少檔案大小');
    }
    
    // 檢查是否需要串流處理
    if (this.shouldUseStreaming()) {
      this.stats.recommendedStrategy.push('啟用串流處理以優化大檔案處理');
    }
    
    // 檢查記憶體優化
    if (this.shouldOptimizeMemory()) {
      this.stats.recommendedStrategy.push('啟用記憶體優化以減少記憶體使用');
    }
    
    // 如果沒有特殊建議，添加預設建議
    if (this.stats.recommendedStrategy.length === 0) {
      this.stats.recommendedStrategy.push('當前配置已是最佳化狀態');
    }
  }

  /**
   * 判斷是否應該使用 sharedStrings
   */
  shouldUseSharedStrings(): boolean {
    return this.stats.stringCells >= this.config.sharedStringsThreshold ||
           this.stats.repetitionRate >= this.config.repetitionRateThreshold;
  }

  /**
   * 判斷是否應該使用串流處理
   */
  shouldUseStreaming(): boolean {
    return this.stats.totalCells >= this.config.largeFileThreshold;
  }

  /**
   * 判斷是否應該優化記憶體
   */
  shouldOptimizeMemory(): boolean {
    return this.config.memoryOptimization && 
           this.stats.totalCells >= this.config.largeFileThreshold;
  }

  /**
   * 獲取優化的儲存格類型
   */
  getOptimizedCellType(value: any): 'inlineStr' | 's' {
    if (typeof value !== 'string') {
      return 'inlineStr';
    }
    
    // 如果啟用了 sharedStrings，使用 sharedStrings
    if (this.shouldUseSharedStrings()) {
      return 's';
    }
    
    // 否則使用 inlineStr
    return 'inlineStr';
  }

  /**
   * 創建效能警告
   */
  createPerformanceWarning(message: string, suggestion?: string): XlsxLiteError {
    return new XlsxLiteError(message, 'UNSUPPORTED_OPERATION', suggestion);
  }

  /**
   * 獲取效能配置
   */
  getConfig(): PerformanceConfig {
    return { ...this.config };
  }

  /**
   * 更新效能配置
   */
  updateConfig(newConfig: Partial<PerformanceConfig>): void {
    this.config = { ...this.config, ...newConfig };
  }

  /**
   * 重置統計資訊
   */
  resetStats(): void {
    this.stats = {
      totalCells: 0,
      stringCells: 0,
      uniqueStrings: 0,
      repetitionRate: 0,
      memoryUsage: 0,
      processingTime: 0,
      recommendedStrategy: []
    };
  }

  /**
   * 獲取記憶體使用量（模擬）
   */
  getMemoryUsage(): number {
    // 這裡可以實現實際的記憶體使用量計算
    // 目前使用模擬數據
    const baseMemory = this.stats.totalCells * 0.001; // 每個儲存格約 1KB
    const stringMemory = this.stats.stringCells * 0.002; // 字串儲存格額外記憶體
    
    this.stats.memoryUsage = Math.round((baseMemory + stringMemory) * 100) / 100;
    return this.stats.memoryUsage;
  }
}

/**
 * 串流處理器
 */
export class StreamingProcessor {
  private chunkSize: number;
  private progressCallback?: (progress: number) => void;

  constructor(chunkSize: number = 1000, progressCallback?: (progress: number) => void) {
    this.chunkSize = chunkSize;
    this.progressCallback = progressCallback;
  }

  /**
   * 分批處理資料
   */
  async processInChunks<T>(
    data: T[],
    processor: (chunk: T[]) => Promise<void>
  ): Promise<void> {
    const totalChunks = Math.ceil(data.length / this.chunkSize);
    
    for (let i = 0; i < totalChunks; i++) {
      const start = i * this.chunkSize;
      const end = Math.min(start + this.chunkSize, data.length);
      const chunk = data.slice(start, end);
      
      await processor(chunk);
      
      // 報告進度
      if (this.progressCallback) {
        const progress = ((i + 1) / totalChunks) * 100;
        this.progressCallback(progress);
      }
    }
  }

  /**
   * 設定進度回調
   */
  setProgressCallback(callback: (progress: number) => void): void {
    this.progressCallback = callback;
  }

  /**
   * 設定分塊大小
   */
  setChunkSize(size: number): void {
    this.chunkSize = size;
  }
}

/**
 * 快取管理器
 */
export class CacheManager {
  private cache: Map<string, any>;
  private maxSize: number;
  private accessCount: Map<string, number>;

  constructor(maxSize: number = 1000) {
    this.cache = new Map();
    this.maxSize = maxSize;
    this.accessCount = new Map();
  }

  /**
   * 獲取快取項目
   */
  get(key: string): any | undefined {
    const value = this.cache.get(key);
    if (value !== undefined) {
      this.accessCount.set(key, (this.accessCount.get(key) || 0) + 1);
    }
    return value;
  }

  /**
   * 設定快取項目
   */
  set(key: string, value: any): void {
    // 如果快取已滿，移除最少使用的項目
    if (this.cache.size >= this.maxSize) {
      this.evictLeastUsed();
    }
    
    this.cache.set(key, value);
    this.accessCount.set(key, 1);
  }

  /**
   * 移除最少使用的項目
   */
  private evictLeastUsed(): void {
    let leastUsedKey = '';
    let leastUsedCount = Infinity;
    
    for (const [key, count] of this.accessCount) {
      if (count < leastUsedCount) {
        leastUsedCount = count;
        leastUsedKey = key;
      }
    }
    
    if (leastUsedKey) {
      this.cache.delete(leastUsedKey);
      this.accessCount.delete(leastUsedKey);
    }
  }

  /**
   * 清除快取
   */
  clear(): void {
    this.cache.clear();
    this.accessCount.clear();
  }

  /**
   * 獲取快取統計
   */
  getStats(): { size: number; maxSize: number; hitRate: number } {
    return {
      size: this.cache.size,
      maxSize: this.maxSize,
      hitRate: 0 // 這裡可以實現實際的命中率計算
    };
  }
}
