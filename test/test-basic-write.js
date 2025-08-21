const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testBasicWrite() {
  console.log('🧪 基本寫入測試 - 驗證 writeBuffer 方法');
  console.log('='.repeat(50));

  try {
    // 創建工作簿
    const workbook = new Workbook();
    console.log('✅ 工作簿創建成功');

    // 創建第一個工作表
    const sheet1 = workbook.getWorksheet('工作表1');
    sheet1.setCell('A1', '標題1', { font: { bold: true } });
    sheet1.setCell('A2', '資料1');
    sheet1.setCell('B2', 100);
    console.log('✅ 工作表1 創建完成');

    // 創建第二個工作表
    const sheet2 = workbook.getWorksheet('工作表2');
    sheet2.setCell('A1', '標題2', { font: { bold: true } });
    sheet2.setCell('A2', '資料2');
    sheet2.setCell('B2', 200);
    console.log('✅ 工作表2 創建完成');

    // 創建第三個工作表
    const sheet3 = workbook.getWorksheet('工作表3');
    sheet3.setCell('A1', '標題3', { font: { bold: true } });
    sheet3.setCell('A2', '資料3');
    sheet3.setCell('B2', 300);
    console.log('✅ 工作表3 創建完成');

    // 檢查工作表數量
    console.log(`📊 工作表數量: ${workbook.getWorksheets().length}`);
    const sheetNames = workbook.getWorksheets().map(ws => ws.name);
    console.log(`📋 工作表名稱: ${sheetNames.join(', ')}`);

    // 使用標準的 writeBuffer 方法
    console.log('\n💾 使用標準 writeBuffer 方法生成 Excel 檔案...');
    const buffer = await workbook.writeBuffer();
    fs.writeFileSync('test-basic-write.xlsx', new Uint8Array(buffer));
    console.log('✅ 標準 Excel 檔案已生成: test-basic-write.xlsx');

    // 檢查檔案大小
    const stats = fs.statSync('test-basic-write.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    console.log('\n🎉 基本寫入測試完成！');

  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error(error.stack);
  }
}

testBasicWrite();
