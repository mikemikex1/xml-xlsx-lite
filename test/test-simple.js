const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testSimple() {
  console.log('🧪 簡單測試 - 驗證基本功能');
  console.log('='.repeat(40));

  try {
    // 創建工作簿
    const workbook = new Workbook();
    console.log('✅ 工作簿創建成功');

    // 創建工作表
    const sheet = workbook.getWorksheet('測試');
    console.log('✅ 工作表創建成功');

    // 設置一些儲存格
    sheet.setCell('A1', '測試標題', { font: { bold: true } });
    sheet.setCell('A2', '數值1');
    sheet.setCell('B2', 100);
    sheet.setCell('A3', '數值2');
    sheet.setCell('B3', 200);
    sheet.setCell('A4', '總計');
    sheet.setCell('B4', '=B2+B3');

    console.log('✅ 儲存格設置完成');

    // 檢查儲存格值
    console.log('📊 儲存格值檢查:');
    console.log(`A1: ${sheet.getCell('A1').value}`);
    console.log(`B2: ${sheet.getCell('B2').value}`);
    console.log(`B3: ${sheet.getCell('B3').value}`);
    console.log(`B4: ${sheet.getCell('B4').value}`);

    // 生成 Excel 檔案
    const buffer = await workbook.writeBuffer();
    fs.writeFileSync('test-simple.xlsx', new Uint8Array(buffer));
    console.log('✅ Excel 檔案已生成: test-simple.xlsx');

    // 檢查檔案大小
    const stats = fs.statSync('test-simple.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    console.log('🎉 簡單測試完成！');

  } catch (error) {
    console.error('❌ 測試失敗:', error);
  }
}

testSimple();
