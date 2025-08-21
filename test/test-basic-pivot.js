const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testBasicPivot() {
  console.log('🧪 測試基本樞紐分析表功能');
  console.log('='.repeat(50));

  try {
    // 創建工作簿
    const workbook = new Workbook();
    console.log('✅ 工作簿創建成功');
    
    // 創建 Detail 工作表
    const detailSheet = workbook.getWorksheet('Detail');
    console.log('✅ Detail 工作表創建成功');
    
    // 設定標題行
    detailSheet.setCell('A1', 'Month', { font: { bold: true } });
    detailSheet.setCell('B1', 'Account', { font: { bold: true } });
    detailSheet.setCell('C1', 'Saving Amount (NTD)', { font: { bold: true } });
    
    // 設定測試資料
    const testData = [
      ['January', 'Account A', 1000],
      ['January', 'Account B', 2000],
      ['February', 'Account A', 1500],
      ['February', 'Account B', 2500]
    ];
    
    // 填入資料
    testData.forEach((row, index) => {
      const rowNum = index + 2;
      detailSheet.setCell(`A${rowNum}`, row[0]);
      detailSheet.setCell(`B${rowNum}`, row[1]);
      detailSheet.setCell(`C${rowNum}`, row[2]);
    });
    
    console.log('✅ 測試資料填入完成');
    
    // 創建樞紐分析表
    const pivotConfig = {
      name: 'Test Pivot',
      sourceRange: 'A1:C5',
      targetRange: 'E1:H10',
      fields: [
        {
          name: 'Month',
          sourceColumn: 'Month',
          type: 'row'
        },
        {
          name: 'Account',
          sourceColumn: 'Account',
          type: 'column'
        },
        {
          name: 'Saving Amount',
          sourceColumn: 'Saving Amount (NTD)',
          type: 'value',
          function: 'sum'
        }
      ]
    };
    
    console.log('🔄 創建樞紐分析表...');
    const pivotTable = workbook.createPivotTable(pivotConfig);
    console.log('✅ 樞紐分析表創建成功');
    
    // 重新整理樞紐分析表
    pivotTable.refresh();
    console.log('✅ 樞紐分析表重新整理完成');
    
    // 匯出到新工作表
    console.log('📋 匯出樞紐分析表...');
    const resultSheet = pivotTable.exportToWorksheet('工作表5');
    console.log('✅ 樞紐分析表已匯出到工作表5');
    
    // 儲存檔案
    const filename = 'test-basic-pivot.xlsx';
    await workbook.writeFile(filename);
    console.log(`💾 檔案已儲存: ${filename}`);
    
    // 顯示工作表清單
    const worksheets = workbook.getWorksheets();
    console.log(`📊 工作表數量: ${worksheets.length}`);
    worksheets.forEach((sheet, index) => {
      console.log(`  ${index + 1}. ${sheet.name}`);
    });
    
    console.log('\n🎉 基本樞紐分析表測試完成！');
    
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    throw error;
  }
}

// 執行測試
testBasicPivot().catch(console.error);
