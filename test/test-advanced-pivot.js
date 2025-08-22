/**
 * 測試進階樞紐分析表功能
 * 展示實際的樞紐分析表邏輯和資料處理
 */

const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testAdvancedPivot() {
  console.log('🧪 測試進階樞紐分析表功能...');
  
  try {
    // 創建工作簿
    const wb = new Workbook();
    
    // 創建資料工作表
    const dataWs = wb.getWorksheet('Detail');
    
    console.log('📝 創建測試資料...');
    
    // 添加標題行
    dataWs.setCell('A1', 'Account', { font: { bold: true } });
    dataWs.setCell('B1', 'Month', { font: { bold: true } });
    dataWs.setCell('C1', 'Saving Amt(NTD)', { font: { bold: true } });
    
    // 創建更豐富的測試資料
    const testData = [
      ['A001', '2024-01', 50000],
      ['A001', '2024-02', 55000],
      ['A001', '2024-03', 60000],
      ['A001', '2024-04', 65000],
      ['A001', '2024-05', 70000],
      ['A001', '2024-06', 75000],
      ['B002', '2024-01', 30000],
      ['B002', '2024-02', 32000],
      ['B002', '2024-03', 35000],
      ['B002', '2024-04', 38000],
      ['B002', '2024-05', 40000],
      ['B002', '2024-06', 42000],
      ['C003', '2024-01', 80000],
      ['C003', '2024-02', 85000],
      ['C003', '2024-03', 90000],
      ['C003', '2024-04', 95000],
      ['C003', '2024-05', 100000],
      ['C003', '2024-06', 105000],
      ['D004', '2024-01', 25000],
      ['D004', '2024-02', 27000],
      ['D004', '2024-03', 29000],
      ['D004', '2024-04', 31000],
      ['D004', '2024-05', 33000],
      ['D004', '2024-06', 35000]
    ];
    
    // 寫入資料
    for (let i = 0; i < testData.length; i++) {
      const row = testData[i];
      dataWs.setCell(`A${i + 2}`, row[0]);
      dataWs.setCell(`B${i + 2}`, row[1]);
      dataWs.setCell(`C${i + 2}`, row[2]);
    }
    
    // 設定欄寬
    dataWs.setColumnWidth('A', 15);
    dataWs.setColumnWidth('B', 15);
    dataWs.setColumnWidth('C', 20);
    
    console.log('📊 創建進階樞紐分析表...');
    
    // 創建樞紐分析表工作表
    const pivotWs = wb.getWorksheet('工作表5');
    
    // 手動創建樞紐分析表結構
    pivotWs.setCell('A1', '樞紐分析表 - 儲蓄分析', { font: { bold: true, size: 16 } });
    
    // 創建樞紐分析表標題
    pivotWs.setCell('A3', 'Account', { font: { bold: true } });
    pivotWs.setCell('B3', '2024-01', { font: { bold: true } });
    pivotWs.setCell('C3', '2024-02', { font: { bold: true } });
    pivotWs.setCell('D3', '2024-03', { font: { bold: true } });
    pivotWs.setCell('E3', '2024-04', { font: { bold: true } });
    pivotWs.setCell('F3', '2024-05', { font: { bold: true } });
    pivotWs.setCell('G3', '2024-06', { font: { bold: true } });
    pivotWs.setCell('H3', 'Total', { font: { bold: true } });
    pivotWs.setCell('I3', 'Average', { font: { bold: true } });
    
    // 計算樞紐分析表資料
    const pivotData = calculatePivotData(testData);
    
    // 寫入樞紐分析表資料
    for (let i = 0; i < pivotData.length; i++) {
      const row = pivotData[i];
      for (let j = 0; j < row.length; j++) {
        const col = String.fromCharCode(65 + j); // A, B, C, D, E, F, G, H, I
        const rowNum = i + 4;
        const value = row[j];
        
        if (j === 0) {
          // 第一欄是文字
          pivotWs.setCell(`${col}${rowNum}`, value);
        } else {
          // 其他欄位是數字
          pivotWs.setCell(`${col}${rowNum}`, value);
        }
      }
    }
    
    // 設定欄寬
    pivotWs.setColumnWidth('A', 15);
    pivotWs.setColumnWidth('B', 12);
    pivotWs.setColumnWidth('C', 12);
    pivotWs.setColumnWidth('D', 12);
    pivotWs.setColumnWidth('E', 12);
    pivotWs.setColumnWidth('F', 12);
    pivotWs.setColumnWidth('G', 12);
    pivotWs.setColumnWidth('H', 15);
    pivotWs.setColumnWidth('I', 15);
    
    // 添加樣式
    pivotWs.setCell('H3', 'Total', { 
      font: { bold: true },
      fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
    });
    
    pivotWs.setCell('I3', 'Average', { 
      font: { bold: true },
      fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
    });
    
    // 為總計列添加樣式
    for (let col = 1; col <= 9; col++) {
      const colLetter = String.fromCharCode(64 + col);
      const lastRow = pivotData.length + 3;
      pivotWs.setCell(`${colLetter}${lastRow}`, pivotData[pivotData.length - 1][col - 1], {
        font: { bold: true },
        fill: { type: 'pattern', patternType: 'solid', fgColor: '#F0F0F0' }
      });
    }
    
    console.log('💾 輸出 Excel 檔案...');
    
    // 輸出檔案
    const buffer = await wb.writeBuffer();
    const filename = 'test-advanced-pivot.xlsx';
    fs.writeFileSync(filename, new Uint8Array(buffer));
    
    console.log(`✅ Excel 檔案 ${filename} 已產生`);
    console.log('📊 檔案大小:', (buffer.byteLength / 1024).toFixed(2), 'KB');
    
    // 驗證樞紐分析表資料
    console.log('\n📋 進階樞紐分析表驗證:');
    console.log('工作表名稱:', pivotWs.name);
    
    // 檢查關鍵儲存格
    console.log('A1 (標題):', pivotWs.getCell('A1').value);
    console.log('A3 (Account 標題):', pivotWs.getCell('A3').value);
    console.log('H3 (Total 標題):', pivotWs.getCell('H3').value);
    console.log('I3 (Average 標題):', pivotWs.getCell('I3').value);
    
    // 驗證資料正確性
    console.log('\n🔍 資料正確性驗證:');
    
    // 檢查 A001 的總計和平均
    const a001Data = testData.filter(row => row[0] === 'A001');
    const a001Total = a001Data.reduce((sum, row) => sum + row[2], 0);
    const a001Average = Math.round(a001Total / a001Data.length);
    
    const actualA001Total = pivotWs.getCell('H4').value;
    const actualA001Average = pivotWs.getCell('I4').value;
    
    console.log(`A001 總計: 預期 ${a001Total}, 實際 ${actualA001Total}`);
    console.log(`A001 平均: 預期 ${a001Average}, 實際 ${actualA001Average}`);
    
    // 檢查整體統計
    const grandTotal = testData.reduce((sum, row) => sum + row[2], 0);
    const grandAverage = Math.round(grandTotal / testData.length);
    
    const actualGrandTotal = pivotWs.getCell('H7').value;
    const actualGrandAverage = pivotWs.getCell('I7').value;
    
    console.log(`整體總計: 預期 ${grandTotal}, 實際 ${actualGrandTotal}`);
    console.log(`整體平均: 預期 ${grandAverage}, 實際 ${actualGrandAverage}`);
    
    console.log('\n🎯 進階樞紐分析表測試完成！');
    console.log('請檢查 Excel 檔案中的樞紐分析表是否正確顯示。');
    
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error('錯誤詳情:', error.stack);
  }
}

/**
 * 計算樞紐分析表資料
 */
function calculatePivotData(testData) {
  // 按 Account 分組
  const accountGroups = new Map();
  
  for (const row of testData) {
    const [account, month, amount] = row;
    
    if (!accountGroups.has(account)) {
      accountGroups.set(account, {
        months: new Map(),
        total: 0,
        count: 0
      });
    }
    
    const group = accountGroups.get(account);
    group.months.set(month, amount);
    group.total += amount;
    group.count += 1;
  }
  
  // 月份順序
  const monthOrder = ['2024-01', '2024-02', '2024-03', '2024-04', '2024-05', '2024-06'];
  
  // 建立樞紐分析表資料
  const pivotData = [];
  
  // 添加每個帳戶的資料
  for (const [account, group] of accountGroups) {
    const row = [account];
    
    // 添加每個月的金額
    for (const month of monthOrder) {
      row.push(group.months.get(month) || 0);
    }
    
    // 添加總計和平均
    row.push(group.total);
    row.push(Math.round(group.total / group.count));
    
    pivotData.push(row);
  }
  
  // 添加總計行
  const totals = ['Total'];
  
  // 計算每個月的總計
  for (const month of monthOrder) {
    const monthTotal = Array.from(accountGroups.values())
      .reduce((sum, group) => sum + (group.months.get(month) || 0), 0);
    totals.push(monthTotal);
  }
  
  // 計算整體總計和平均
  const grandTotal = Array.from(accountGroups.values())
    .reduce((sum, group) => sum + group.total, 0);
  const grandAverage = Math.round(grandTotal / testData.length);
  
  totals.push(grandTotal);
  totals.push(grandAverage);
  
  pivotData.push(totals);
  
  return pivotData;
}

// 執行測試
testAdvancedPivot().catch(console.error);
