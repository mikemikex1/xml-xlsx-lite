const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testStepByStep() {
  console.log('🧪 逐步測試 - 找出問題所在');
  console.log('='.repeat(50));

  try {
    // 創建工作簿
    const workbook = new Workbook();
    console.log('✅ 工作簿創建成功');

    // 步驟 1: 基本功能
    console.log('\n📊 步驟 1: 基本功能');
    const basicSheet = workbook.getWorksheet('基本功能');
    basicSheet.setCell('A1', '產品名稱', { font: { bold: true, size: 14 } });
    basicSheet.setCell('B1', '數量', { font: { bold: true, size: 14 } });
    basicSheet.setCell('C1', '單價', { font: { bold: true, size: 14 } });
    basicSheet.setCell('D1', '總價', { font: { bold: true, size: 14 } });
    basicSheet.setCell('A2', '筆記型電腦');
    basicSheet.setCell('B2', 5);
    basicSheet.setCell('C2', 80000);
    basicSheet.setCell('D2', '=B2*C2');
    console.log('✅ 基本功能完成');

    // 步驟 2: 樣式支援
    console.log('\n🎨 步驟 2: 樣式支援');
    const styleSheet = workbook.getWorksheet('樣式測試');
    styleSheet.setCell('A1', '粗體文字', { font: { bold: true, size: 16, color: 'FF0000' } });
    styleSheet.setCell('A2', '斜體文字', { font: { italic: true, size: 14, color: '0000FF' } });
    styleSheet.setCell('A3', '底線文字', { font: { underline: true, size: 12 } });
    console.log('✅ 樣式支援完成');

    // 步驟 3: 進階功能
    console.log('\n🔧 步驟 3: 進階功能');
    const advancedSheet = workbook.getWorksheet('進階功能');
    advancedSheet.mergeCells('A1:C1');
    advancedSheet.setCell('A1', '合併儲存格標題', { 
      font: { bold: true, size: 16 },
      alignment: { horizontal: 'center' }
    });
    advancedSheet.setColumnWidth('A', 20);
    advancedSheet.setColumnWidth('B', 15);
    advancedSheet.setColumnWidth('C', 15);
    advancedSheet.setRowHeight(1, 30);
    advancedSheet.freezePanes(2, 1);
    console.log('✅ 進階功能完成');

    // 步驟 4: 效能優化
    console.log('\n⚡ 步驟 4: 效能優化');
    const perfSheet = workbook.getWorksheet('效能測試');
    const largeData = [];
    for (let i = 0; i < 100; i++) {
      largeData.push([
        `產品${i + 1}`,
        Math.floor(Math.random() * 1000),
        Math.floor(Math.random() * 10000) + 1000,
        Math.floor(Math.random() * 100) + 1
      ]);
    }
    
    await workbook.addLargeDataset('效能測試', largeData, {
      startRow: 2,
      startCol: 1,
      chunkSize: 50
    });
    console.log('✅ 效能優化完成');

    // 檢查當前狀態
    console.log('\n🔍 當前狀態檢查:');
    console.log(`工作表數量: ${workbook.getWorksheets().length}`);
    const sheetNames = workbook.getWorksheets().map(ws => ws.name);
    console.log(`工作表名稱: ${sheetNames.join(', ')}`);

    // 檢查每個工作表的資料
    for (const sheetName of sheetNames) {
      const sheet = workbook.getWorksheet(sheetName);
      let rowCount = 0;
      for (const [rowNum, rowMap] of sheet.rows()) {
        rowCount++;
      }
      console.log(`${sheetName}: ${rowCount} 行`);
    }

    // 生成 Excel 檔案
    console.log('\n💾 生成 Excel 檔案...');
    const buffer = await workbook.writeBuffer();
    fs.writeFileSync('test-step-by-step.xlsx', new Uint8Array(buffer));
    console.log('✅ Excel 檔案已生成: test-step-by-step.xlsx');

    // 檢查檔案大小
    const stats = fs.statSync('test-step-by-step.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    console.log('\n🎉 逐步測試完成！');

  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error(error.stack);
  }
}

testStepByStep();
