/**
 * 進階效能優化測試
 * 展示實際的效能優化器功能
 */

const { Workbook, PerformanceOptimizer, StreamingProcessor, CacheManager } = require('../dist/index.js');
const fs = require('fs');

async function testAdvancedPerformance() {
  console.log('🧪 進階效能優化測試...');
  
  try {
    // 創建工作簿
    const wb = new Workbook();
    
    console.log('📝 1. 創建測試資料...');
    
    // 創建測試工作表
    const testWs = wb.getWorksheet('Performance Test');
    
    // 添加標題
    testWs.setCell('A1', '效能測試資料', { font: { bold: true, size: 16 } });
    testWs.setCell('A3', 'ID', { font: { bold: true } });
    testWs.setCell('B3', 'Category', { font: { bold: true } });
    testWs.setCell('C3', 'Status', { font: { bold: true } });
    testWs.setCell('D3', 'Description', { font: { bold: true } });
    testWs.setCell('E3', 'Value', { font: { bold: true } });
    
    // 生成測試資料
    const categories = ['A類', 'B類', 'C類', 'D類', 'E類'];
    const statuses = ['進行中', '已完成', '待處理', '已取消', '暫停'];
    const descriptions = [
      '重要任務', '一般任務', '緊急任務', '例行任務', '特殊任務',
      '系統維護', '資料備份', '效能監控', '安全檢查', '更新部署'
    ];
    
    console.log('📊 生成 2,000 筆測試資料...');
    
    for (let i = 0; i < 2000; i++) {
      const row = i + 4;
      const category = categories[i % categories.length];
      const status = statuses[i % statuses.length];
      const description = descriptions[i % descriptions.length] + (i % 100);
      const value = Math.floor(Math.random() * 1000) + 100;
      
      testWs.setCell(`A${row}`, i + 1);
      testWs.setCell(`B${row}`, category);
      testWs.setCell(`C${row}`, status);
      testWs.setCell(`D${row}`, description);
      testWs.setCell(`E${row}`, value);
      
      if ((i + 1) % 500 === 0) {
        console.log(`  已生成 ${i + 1} 筆資料...`);
      }
    }
    
    // 設定欄寬
    testWs.setColumnWidth('A', 10);
    testWs.setColumnWidth('B', 15);
    testWs.setColumnWidth('C', 15);
    testWs.setColumnWidth('D', 25);
    testWs.setColumnWidth('E', 15);
    
    console.log('✅ 測試資料生成完成');
    
    console.log('\n🔧 2. 測試效能優化器...');
    
    // 創建效能優化器
    const optimizer = new PerformanceOptimizer({
      sharedStringsThreshold: 50,
      repetitionRateThreshold: 25,
      largeFileThreshold: 1000,
      streamingThreshold: 1,
      cacheSizeLimit: 50,
      memoryOptimization: true
    });
    
    // 分析工作表效能
    const stats = optimizer.analyzeWorksheet(testWs);
    
    console.log('📊 效能分析結果:');
    console.log(`  總儲存格數量: ${stats.totalCells.toLocaleString()}`);
    console.log(`  字串儲存格數量: ${stats.stringCells.toLocaleString()}`);
    console.log(`  唯一字串數量: ${stats.uniqueStrings.toLocaleString()}`);
    console.log(`  字串重複率: ${stats.repetitionRate.toFixed(1)}%`);
    console.log(`  記憶體使用量: ${stats.memoryUsage} MB`);
    console.log(`  處理時間: ${stats.processingTime} ms`);
    
    console.log('\n💡 優化建議:');
    stats.recommendedStrategy.forEach((strategy, index) => {
      console.log(`  ${index + 1}. ${strategy}`);
    });
    
    // 測試優化決策
    console.log('\n🎯 優化決策測試:');
    console.log(`  是否使用 sharedStrings: ${optimizer.shouldUseSharedStrings() ? '是' : '否'}`);
    console.log(`  是否使用串流處理: ${optimizer.shouldUseStreaming() ? '是' : '否'}`);
    console.log(`  是否優化記憶體: ${optimizer.shouldOptimizeMemory() ? '是' : '否'}`);
    
    console.log('\n🔧 3. 測試串流處理器...');
    
    // 創建串流處理器
    const streamer = new StreamingProcessor(100, (progress) => {
      console.log(`  串流處理進度: ${progress.toFixed(1)}%`);
    });
    
    // 模擬串流處理
    const testData = Array.from({ length: 1000 }, (_, i) => `資料${i + 1}`);
    
    console.log('📊 開始串流處理...');
    await streamer.processInChunks(testData, async (chunk) => {
      // 模擬處理邏輯
      await new Promise(resolve => setTimeout(resolve, 10));
    });
    
    console.log('✅ 串流處理完成');
    
    console.log('\n🔧 4. 測試快取管理器...');
    
    // 創建快取管理器
    const cache = new CacheManager(100);
    
    // 測試快取功能
    for (let i = 0; i < 150; i++) {
      cache.set(`key${i}`, `value${i}`);
    }
    
    const cacheStats = cache.getStats();
    console.log('📊 快取統計:');
    console.log(`  當前大小: ${cacheStats.size}`);
    console.log(`  最大大小: ${cacheStats.maxSize}`);
    console.log(`  命中率: ${cacheStats.hitRate}%`);
    
    // 測試快取存取
    const testValue = cache.get('key50');
    console.log(`  測試快取存取: key50 = ${testValue}`);
    
    console.log('✅ 快取測試完成');
    
    console.log('\n💾 5. 輸出 Excel 檔案...');
    
    // 輸出檔案
    const buffer = await wb.writeBuffer();
    const filename = 'test-advanced-performance.xlsx';
    fs.writeFileSync(filename, new Uint8Array(buffer));
    
    console.log(`✅ Excel 檔案 ${filename} 已產生`);
    console.log('📊 檔案大小:', (buffer.byteLength / 1024).toFixed(2), 'KB');
    
    // 效能優化總結
    console.log('\n📈 效能優化總結:');
    
    if (stats.repetitionRate > 50) {
      console.log('✅ 高重複率檢測到，建議啟用 sharedStrings 以減少檔案大小');
    }
    
    if (stats.totalCells > 1000) {
      console.log('✅ 大檔案檢測到，建議啟用串流處理以優化處理效能');
    }
    
    if (stats.memoryUsage > 5) {
      console.log('✅ 高記憶體使用檢測到，建議啟用記憶體優化');
    }
    
    // 顯示配置資訊
    const config = optimizer.getConfig();
    console.log('\n⚙️ 當前效能配置:');
    console.log(`  sharedStrings 閾值: ${config.sharedStringsThreshold}`);
    console.log(`  重複率閾值: ${config.repetitionRateThreshold}%`);
    console.log(`  大檔案閾值: ${config.largeFileThreshold.toLocaleString()}`);
    console.log(`  串流閾值: ${config.streamingThreshold} MB`);
    console.log(`  快取大小限制: ${config.cacheSizeLimit} MB`);
    console.log(`  記憶體優化: ${config.memoryOptimization ? '啟用' : '停用'}`);
    
    console.log('\n🎯 進階效能優化測試完成！');
    console.log('所有效能優化功能都已成功測試。');
    
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error('錯誤詳情:', error.stack);
  }
}

// 執行測試
testAdvancedPerformance().catch(console.error);
