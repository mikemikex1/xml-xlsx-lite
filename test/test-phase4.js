const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testPhase4() {
  console.log('⚡ 測試 Phase 4: 效能優化');
  
  // 建立具有效能優化選項的工作簿
  const wb = new Workbook({
    memoryOptimization: true,
    chunkSize: 500,
    cacheEnabled: true,
    maxCacheSize: 5000
  });
  
  console.log('🔧 效能優化設定已啟用');
  
  // 測試大型資料集處理
  console.log('📊 測試大型資料集處理...');
  
  // 生成 10,000 筆測試資料
  const largeDataset = [];
  for (let i = 0; i < 10000; i++) {
    largeDataset.push([
      `產品${i + 1}`,
      Math.floor(Math.random() * 1000),
      Math.random() * 1000,
      new Date(2024, 0, 1 + (i % 365)),
      i % 2 === 0 ? '啟用' : '停用'
    ]);
  }
  
  const startTime = Date.now();
  
  // 使用優化的大型資料集方法
  await wb.addLargeDataset('大型資料測試', largeDataset, {
    startRow: 2,
    startCol: 1,
    chunkSize: 500
  });
  
  const processingTime = Date.now() - startTime;
  console.log(`✅ 大型資料集處理完成！耗時: ${processingTime}ms`);
  
  // 設定標題
  const ws = wb.getWorksheet('大型資料測試');
  ws.setCell('A1', '產品名稱', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
  });
  ws.setCell('B1', '數量', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
  });
  ws.setCell('C1', '價格', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
  });
  ws.setCell('D1', '日期', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
  });
  ws.setCell('E1', '狀態', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
  });
  
  // 設定欄寬
  ws.setColumnWidth('A', 20);
  ws.setColumnWidth('B', 15);
  ws.setColumnWidth('C', 15);
  ws.setColumnWidth('D', 15);
  ws.setColumnWidth('E', 15);
  
  // 測試記憶體統計
  console.log('\n📊 記憶體使用統計:');
  const memoryStats = wb.getMemoryStats();
  console.log('工作表數量:', memoryStats.sheets);
  console.log('總儲存格數量:', memoryStats.totalCells.toLocaleString());
  console.log('快取大小:', memoryStats.cacheSize);
  console.log('快取命中率:', (memoryStats.cacheHitRate * 100).toFixed(1) + '%');
  console.log('估算記憶體使用:', (memoryStats.memoryUsage / 1024 / 1024).toFixed(2) + ' MB');
  
  // 測試效能優化設定
  console.log('\n🔧 測試效能優化設定...');
  
  // 調整分塊大小
  wb.setChunkSize(1000);
  console.log('✅ 分塊大小已調整為 1000');
  
  // 調整快取大小
  wb.setMaxCacheSize(10000);
  console.log('✅ 快取大小限制已調整為 10000');
  
  // 測試記憶體回收
  console.log('\n🗑️ 測試記憶體回收...');
  const beforeGC = wb.getMemoryStats();
  wb.forceGarbageCollection();
  const afterGC = wb.getMemoryStats();
  
  console.log('記憶體回收前:', (beforeGC.memoryUsage / 1024 / 1024).toFixed(2) + ' MB');
  console.log('記憶體回收後:', (afterGC.memoryUsage / 1024 / 1024).toFixed(2) + ' MB');
  console.log('節省記憶體:', ((beforeGC.memoryUsage - afterGC.memoryUsage) / 1024 / 1024).toFixed(2) + ' MB');
  
  // 測試串流寫入
  console.log('\n🌊 測試串流寫入...');
  
  const writeStream = async (chunk) => {
    // 模擬串流寫入
    process.stdout.write(`\r串流寫入中... ${chunk.length} bytes`);
  };
  
  try {
    await wb.writeStream(writeStream);
    console.log('\n✅ 串流寫入測試完成');
  } catch (error) {
    console.log('\n⚠️ 串流寫入測試失敗:', error.message);
  }
  
  // 生成傳統 Excel 檔案
  console.log('\n💾 生成 Excel 檔案...');
  const buffer = await wb.writeBuffer();
  
  const filename = 'test-phase4.xlsx';
  fs.writeFileSync(filename, Buffer.from(buffer));
  console.log(`✅ Phase 4 測試完成！檔案已儲存為: ${filename}`);
  
  // 最終記憶體統計
  console.log('\n📊 最終記憶體統計:');
  const finalStats = wb.getMemoryStats();
  console.log('工作表數量:', finalStats.sheets);
  console.log('總儲存格數量:', finalStats.totalCells.toLocaleString());
  console.log('快取大小:', finalStats.cacheSize);
  console.log('估算記憶體使用:', (finalStats.memoryUsage / 1024 / 1024).toFixed(2) + ' MB');
  
  // 效能建議
  console.log('\n💡 效能優化建議:');
  if (finalStats.totalCells > 100000) {
    console.log('- 建議啟用記憶體優化');
    console.log('- 考慮使用串流處理');
    console.log('- 定期執行記憶體回收');
  }
  
  if (finalStats.cacheSize > finalStats.maxCacheSize * 0.8) {
    console.log('- 快取使用率較高，考慮增加快取大小限制');
  }
  
  console.log('\n🎯 Phase 4 效能優化功能測試完成！');
}

// 執行測試
testPhase4().catch(console.error);
