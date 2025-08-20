const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testBasic() {
  console.log('🧪 測試 Phase 1: 基本功能');
  console.log('=' .repeat(50));
  
  try {
    // 建立工作簿
    const workbook = new Workbook();
    
    // 建立工作表
    const sheet = workbook.getWorksheet('基本測試');
    
    // 測試基本儲存格操作
    sheet.setCell('A1', '產品名稱');
    sheet.setCell('B1', '數量');
    sheet.setCell('C1', '單價');
    sheet.setCell('D1', '總價');
    
    sheet.setCell('A2', 'iPhone 15');
    sheet.setCell('B2', 10);
    sheet.setCell('C2', 35000);
    sheet.setCell('D2', 350000);
    
    sheet.setCell('A3', 'MacBook Pro');
    sheet.setCell('B3', 5);
    sheet.setCell('C3', 80000);
    sheet.setCell('D3', 400000);
    
    // 測試不同資料類型
    sheet.setCell('A4', '日期測試');
    sheet.setCell('B4', new Date());
    sheet.setCell('C4', true);
    sheet.setCell('D4', false);
    
    console.log('✅ 基本儲存格操作完成');
    
    // 測試多工作表
    const sheet2 = workbook.getWorksheet('第二工作表');
    sheet2.setCell('A1', '第二工作表的資料');
    sheet2.setCell('B1', 42);
    
    console.log('✅ 多工作表支援完成');
    
    // 匯出檔案
    const buffer = await workbook.writeBuffer();
    fs.writeFileSync('test-basic.xlsx', new Uint8Array(buffer));
    
    console.log('✅ 檔案匯出完成: test-basic.xlsx');
    console.log(`✅ 工作表數量: ${workbook.getWorksheets().length}`);
    
    console.log('\n🎉 Phase 1 基本功能測試完成！');
    
  } catch (error) {
    console.error('❌ 測試失敗:', error.message);
    throw error;
  }
}

testBasic().catch(console.error);
