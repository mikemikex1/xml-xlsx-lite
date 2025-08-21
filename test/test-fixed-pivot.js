const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testFixedPivot() {
  console.log('🧪 測試修正後的樞紐分析表功能');
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
    
    // 設定測試資料 - 使用更簡單的資料結構
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
    console.log('📋 測試資料:');
    testData.forEach((row, index) => {
      console.log(`  行 ${index + 2}: ${row.join(', ')}`);
    });
    
    // 創建樞紐分析表配置 - 簡化配置
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
      ],
      showRowSubtotals: false,
      showColumnSubtotals: false,
      showGrandTotals: false
    };
    
    console.log('\n🔄 創建樞紐分析表...');
    console.log('配置:', JSON.stringify(pivotConfig, null, 2));
    
    const pivotTable = workbook.createPivotTable(pivotConfig);
    console.log('✅ 樞紐分析表創建成功');
    
    // 重新整理樞紐分析表
    pivotTable.refresh();
    console.log('✅ 樞紐分析表重新整理完成');
    
    // 檢查處理後的資料
    const pivotData = pivotTable.getData();
    console.log('\n📊 樞紐分析表資料:');
    console.log('資料行數:', pivotData.length);
    pivotData.forEach((row, index) => {
      console.log(`  行 ${index}: [${row.join(', ')}]`);
    });
    
    // 匯出到新工作表
    console.log('\n📋 匯出樞紐分析表...');
    const resultSheet = pivotTable.exportToWorksheet('工作表5');
    console.log('✅ 樞紐分析表已匯出到工作表5');
    
    // 檢查工作表5 的內容
    console.log('\n🔍 檢查工作表5 內容:');
    let rowCount = 0;
    for (const [rowNum, rowMap] of resultSheet.rows()) {
      if (rowCount < 10) { // 只顯示前10行
        const rowData = [];
        for (let col = 0; col < 4; col++) {
          const cell = rowMap.get(col + 1);
          if (cell) {
            rowData.push(cell.value);
          } else {
            rowData.push('(空)');
          }
        }
        console.log(`  行 ${rowNum}: [${rowData.join(', ')}]`);
      }
      rowCount++;
    }
    console.log(`工作表5總行數: ${rowCount}`);
    
    // 手動創建預期的樞紐分析表結果到新工作表
    console.log('\n📊 手動創建預期結果...');
    const expectedSheet = workbook.getWorksheet('Expected Results');
    
    // 設定標題
    expectedSheet.setCell('A1', '預期結果 - 儲蓄金額彙總', {
      font: { bold: true, size: 16 },
      alignment: { horizontal: 'center' }
    });
    
    // 設定欄標題
    expectedSheet.setCell('A3', 'Month', { font: { bold: true } });
    expectedSheet.setCell('B3', 'Account A', { font: { bold: true } });
    expectedSheet.setCell('C3', 'Account B', { font: { bold: true } });
    
    // 計算並填入預期結果
    const expectedData = [
      ['January', 1000, 2000],
      ['February', 1500, 2500]
    ];
    
    expectedData.forEach((row, index) => {
      const rowNum = index + 4;
      expectedSheet.setCell(`A${rowNum}`, row[0]);
      expectedSheet.setCell(`B${rowNum}`, row[1], { numFmt: '#,##0' });
      expectedSheet.setCell(`C${rowNum}`, row[2], { numFmt: '#,##0' });
    });
    
    // 設定欄寬
    expectedSheet.setColumnWidth('A', 15);
    expectedSheet.setColumnWidth('B', 15);
    expectedSheet.setColumnWidth('C', 15);
    
    console.log('✅ 預期結果工作表創建完成');
    
    // 使用 writeBuffer 方法儲存檔案
    console.log('\n💾 使用 writeBuffer 方法儲存檔案...');
    const buffer = await workbook.writeBuffer();
    const filename = 'test-fixed-pivot.xlsx';
    fs.writeFileSync(filename, new Uint8Array(buffer));
    console.log(`✅ 檔案已儲存: ${filename}`);
    
    // 顯示檔案統計
    const stats = fs.statSync(filename);
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);
    
    // 顯示工作表清單
    const worksheets = workbook.getWorksheets();
    console.log(`📊 工作表數量: ${worksheets.length}`);
    console.log('\n📋 工作表清單:');
    worksheets.forEach((sheet, index) => {
      console.log(`  ${index + 1}. ${sheet.name}`);
    });
    
    console.log('\n🎉 修正後的樞紐分析表測試完成！');
    console.log('\n📝 請檢查生成的檔案，確認：');
    console.log('  1. Detail 工作表包含正確的測試資料');
    console.log('  2. 工作表5 包含樞紐分析表結果');
    console.log('  3. Expected Results 工作表顯示預期結果');
    console.log('  4. 資料一致性驗證');
    
    // 分析樞紐分析表資料問題
    console.log('\n🔍 樞紐分析表資料分析:');
    console.log('問題分析:');
    console.log('  1. 第一行標題不正確: 應該是 [Month, Account A, Account B]');
    console.log('  2. 資料行格式不正確: 應該有正確的行標題和數值');
    console.log('  3. 最後一行資料異常: 包含不正確的數值');
    
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error('錯誤堆疊:', error.stack);
    throw error;
  }
}

// 執行測試
testFixedPivot().catch(console.error);
