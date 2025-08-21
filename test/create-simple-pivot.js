const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function createSimplePivot() {
  console.log('🧪 創建簡單的樞紐分析表示範');
  console.log('='.repeat(50));

  try {
    // 創建工作簿
    const workbook = new Workbook();
    
    // 創建 Detail 工作表
    console.log('\n📊 創建 Detail 工作表...');
    const detailSheet = workbook.getWorksheet('Detail');
    
    // 設定標題行
    const headers = ['Month', 'Account', 'Saving Amount (NTD)'];
    headers.forEach((header, index) => {
      detailSheet.setCell(`${String.fromCharCode(65 + index)}1`, header, {
        font: { bold: true, size: 14 },
        fill: { type: 'pattern', patternType: 'solid', fgColor: '#E0E0E0' },
        alignment: { horizontal: 'center' }
      });
    });
    
    // 設定欄寬
    detailSheet.setColumnWidth('A', 15); // Month
    detailSheet.setColumnWidth('B', 20); // Account
    detailSheet.setColumnWidth('C', 18); // Saving Amount
    
    // 生成簡單的測試資料
    const months = ['January', 'February', 'March'];
    const accounts = ['Account A', 'Account B'];
    
    let rowIndex = 2;
    const testData = [];
    
    // 為每個月份和帳戶生成固定的儲蓄金額（便於驗證）
    for (const month of months) {
      for (const account of accounts) {
        const amount = (months.indexOf(month) + 1) * 1000 + (accounts.indexOf(account) + 1) * 100;
        
        detailSheet.setCell(`A${rowIndex}`, month);
        detailSheet.setCell(`B${rowIndex}`, account);
        detailSheet.setCell(`C${rowIndex}`, amount, {
          numFmt: '#,##0',
          alignment: { horizontal: 'right' }
        });
        
        testData.push({ month, account, amount });
        rowIndex++;
      }
    }
    
    console.log(`✅ Detail 工作表創建完成，共 ${testData.length} 筆資料`);
    
    // 顯示測試資料
    console.log('\n📋 測試資料:');
    testData.forEach(data => {
      console.log(`  ${data.month} - ${data.account}: ${data.amount}`);
    });
    
    // 創建樞紐分析表配置
    console.log('\n🔄 創建樞紐分析表...');
    const pivotConfig = {
      name: 'Savings Summary',
      sourceRange: `A1:C${rowIndex - 1}`,
      targetRange: 'E1:H20',
      fields: [
        {
          name: 'Month',
          sourceColumn: 'Month',
          type: 'row',
          showSubtotal: true
        },
        {
          name: 'Account',
          sourceColumn: 'Account',
          type: 'column',
          showSubtotal: true
        },
        {
          name: 'Saving Amount',
          sourceColumn: 'Saving Amount (NTD)',
          type: 'value',
          function: 'sum',
          numberFormat: '#,##0'
        }
      ],
      showGrandTotals: true,
      autoFormat: true
    };
    
    // 創建樞紐分析表
    const pivotTable = workbook.createPivotTable(pivotConfig);
    pivotTable.refresh();
    
    console.log('✅ 樞紐分析表創建完成');
    
    // 將樞紐分析表結果匯出到工作表5
    console.log('\n📋 匯出樞紐分析表到工作表5...');
    const pivotResultSheet = pivotTable.exportToWorksheet('工作表5');
    
    // 設定工作表5的標題
    pivotResultSheet.setCell('A1', '樞紐分析表結果 - 儲蓄金額彙總', {
      font: { bold: true, size: 16 },
      alignment: { horizontal: 'center' }
    });
    
    // 設定欄寬
    pivotResultSheet.setColumnWidth('A', 20);
    pivotResultSheet.setColumnWidth('B', 20);
    pivotResultSheet.setColumnWidth('C', 20);
    pivotResultSheet.setColumnWidth('D', 20);
    
    console.log('✅ 樞紐分析表已匯出到工作表5');
    
    // 手動創建預期的樞紐分析表結果
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
    expectedSheet.setCell('D3', 'Total', { font: { bold: true } });
    
    // 計算並填入預期結果
    let expectedRow = 4;
    for (const month of months) {
      const monthData = testData.filter(d => d.month === month);
      const accountA = monthData.find(d => d.account === 'Account A')?.amount || 0;
      const accountB = monthData.find(d => d.account === 'Account B')?.amount || 0;
      const total = accountA + accountB;
      
      expectedSheet.setCell(`A${expectedRow}`, month);
      expectedSheet.setCell(`B${expectedRow}`, accountA, { numFmt: '#,##0' });
      expectedSheet.setCell(`C${expectedRow}`, accountB, { numFmt: '#,##0' });
      expectedSheet.setCell(`D${expectedRow}`, total, { 
        numFmt: '#,##0',
        font: { bold: true }
      });
      
      expectedRow++;
    }
    
    // 設定欄寬
    expectedSheet.setColumnWidth('A', 15);
    expectedSheet.setColumnWidth('B', 15);
    expectedSheet.setColumnWidth('C', 15);
    expectedSheet.setColumnWidth('D', 15);
    
    console.log('✅ 預期結果工作表創建完成');
    
    // 儲存檔案
    const filename = 'simple-pivot-example.xlsx';
    await workbook.writeFile(filename);
    console.log(`\n💾 檔案已儲存: ${filename}`);
    
    // 顯示檔案統計
    const stats = fs.statSync(filename);
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);
    console.log(`📊 工作表數量: ${workbook.getWorksheets().length}`);
    
    // 顯示工作表清單
    console.log('\n📋 工作表清單:');
    workbook.getWorksheets().forEach((sheet, index) => {
      console.log(`${index + 1}. ${sheet.name}`);
    });
    
    console.log('\n🎉 簡單樞紐分析表示範完成！');
    console.log('\n📝 請檢查生成的檔案，確認：');
    console.log('  1. Detail 工作表包含正確的測試資料');
    console.log('  2. 工作表5 包含樞紐分析表結果');
    console.log('  3. Expected Results 工作表顯示預期結果');
    console.log('  4. 資料一致性驗證');
    
  } catch (error) {
    console.error('❌ 創建失敗:', error);
    throw error;
  }
}

// 執行創建
createSimplePivot().catch(console.error);
