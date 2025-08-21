const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testDebug() {
  console.log('🔍 調試測試 - 診斷問題');
  console.log('='.repeat(40));

  try {
    // 創建工作簿
    const workbook = new Workbook();
    console.log('✅ 工作簿創建成功');

    // 創建一個簡單的工作表
    const sheet = workbook.getWorksheet('調試');
    console.log('✅ 工作表創建成功');

    // 設置一些儲存格
    sheet.setCell('A1', '標題', { font: { bold: true } });
    sheet.setCell('A2', '資料1');
    sheet.setCell('B2', 100);
    sheet.setCell('A3', '資料2');
    sheet.setCell('B3', 200);

    console.log('✅ 儲存格設置完成');

    // 檢查工作表的內部狀態
    console.log('\n🔍 工作表內部狀態檢查:');
    console.log(`工作表名稱: ${sheet.name}`);
    console.log(`工作表保護狀態: ${sheet.isProtected()}`);
    
    // 檢查儲存格
    console.log('\n📊 儲存格檢查:');
    console.log(`A1: ${sheet.getCell('A1').value}`);
    console.log(`A2: ${sheet.getCell('A2').value}`);
    console.log(`B2: ${sheet.getCell('B2').value}`);
    console.log(`A3: ${sheet.getCell('A3').value}`);
    console.log(`B3: ${sheet.getCell('B3').value}`);

    // 檢查 rows() 方法
    console.log('\n🔍 rows() 方法檢查:');
    let rowCount = 0;
    for (const [rowNum, rowMap] of sheet.rows()) {
      console.log(`行 ${rowNum}: ${rowMap.size} 個儲存格`);
      rowCount++;
    }
    console.log(`總行數: ${rowCount}`);

    // 檢查工作簿狀態
    console.log('\n🔍 工作簿狀態檢查:');
    console.log(`工作表數量: ${workbook.getWorksheets().length}`);
    const sheetNames = workbook.getWorksheets().map(ws => ws.name);
    console.log(`工作表名稱: ${sheetNames.join(', ')}`);

    // 生成 Excel 檔案
    console.log('\n💾 生成 Excel 檔案...');
    const buffer = await workbook.writeBuffer();
    fs.writeFileSync('test-debug.xlsx', new Uint8Array(buffer));
    console.log('✅ Excel 檔案已生成: test-debug.xlsx');

    // 檢查檔案大小
    const stats = fs.statSync('test-debug.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    console.log('\n🎉 調試測試完成！');

  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error(error.stack);
  }
}

testDebug();
