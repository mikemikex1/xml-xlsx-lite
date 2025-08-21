const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function processAvapTemplate() {
  console.log('🧪 處理 AVAP Saving Report Template');
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
    
    // 生成測試資料
    const months = ['January', 'February', 'March', 'April', 'May', 'June'];
    const accounts = ['Account A', 'Account B', 'Account C', 'Account D'];
    
    let rowIndex = 2;
    const testData = [];
    
    // 為每個月份和帳戶生成隨機儲蓄金額
    for (const month of months) {
      for (const account of accounts) {
        const amount = Math.floor(Math.random() * 10000) + 1000; // 1000-11000 之間
        
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
          showSubtotal: true,
          sortOrder: 'asc'
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
          numberFormat: '#,##0',
          customName: 'Total Savings'
        }
      ],
      showRowHeaders: true,
      showColumnHeaders: true,
      showRowSubtotals: true,
      showColumnSubtotals: true,
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
    
    // 驗證 Detail 和工作表5 的資料一致性
    console.log('\n🔍 驗證資料一致性...');
    
    // 從 Detail 工作表計算預期結果
    const expectedResults = {};
    for (const data of testData) {
      if (!expectedResults[data.month]) {
        expectedResults[data.month] = {};
      }
      if (!expectedResults[data.month][data.account]) {
        expectedResults[data.month][data.account] = 0;
      }
      expectedResults[data.month][data.account] += data.amount;
    }
    
    // 從工作表5 讀取實際結果
    const actualResults = {};
    let pivotDataRow = 3; // 跳過標題行和樞紐標題行
    
    // 讀取樞紐分析表的資料
    while (true) {
      const monthCell = pivotResultSheet.getCell(`A${pivotDataRow}`);
      if (!monthCell.value || monthCell.value === 'Grand Total') break;
      
      const month = monthCell.value;
      if (!actualResults[month]) {
        actualResults[month] = {};
      }
      
      // 讀取每個帳戶的儲蓄金額
      for (let col = 1; col < 4; col++) { // Account A, B, C, D
        const accountCol = String.fromCharCode(66 + col); // B, C, D, E
        const amountCell = pivotResultSheet.getCell(`${accountCol}${pivotDataRow}`);
        
        if (amountCell.value && typeof amountCell.value === 'number') {
          const accountName = `Account ${String.fromCharCode(64 + col)}`; // A, B, C, D
          actualResults[month][accountName] = amountCell.value;
        }
      }
      
      pivotDataRow++;
    }
    
    // 比較預期和實際結果
    let isConsistent = true;
    const comparisonReport = [];
    
    for (const month of months) {
      for (const account of accounts) {
        const expected = expectedResults[month]?.[account] || 0;
        const actual = actualResults[month]?.[account] || 0;
        
        if (expected !== actual) {
          isConsistent = false;
          comparisonReport.push(`❌ ${month} - ${account}: 預期 ${expected}, 實際 ${actual}`);
        } else {
          comparisonReport.push(`✅ ${month} - ${account}: ${expected}`);
        }
      }
    }
    
    console.log('\n📊 資料一致性驗證結果:');
    console.log('-'.repeat(40));
    
    if (isConsistent) {
      console.log('🎉 所有資料完全一致！');
    } else {
      console.log('⚠️  發現資料不一致的情況：');
      comparisonReport.forEach(line => console.log(line));
    }
    
    // 顯示詳細比較報告
    console.log('\n📋 詳細比較報告:');
    console.log('-'.repeat(40));
    comparisonReport.forEach(line => console.log(line));
    
    // 儲存檔案
    const filename = 'avap-saving-report-processed.xlsx';
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
    
    console.log('\n🎉 AVAP Template 處理完成！');
    
  } catch (error) {
    console.error('❌ 處理失敗:', error);
    throw error;
  }
}

// 執行處理
processAvapTemplate().catch(console.error);
