/**
 * 測試樞紐分析表功能
 */

const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testPivotTable() {
  console.log('🧪 測試樞紐分析表功能...');
  
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
    
    // 添加測試資料
    const testData = [
      ['A001', '2024-01', 50000],
      ['A001', '2024-02', 55000],
      ['A001', '2024-03', 60000],
      ['B002', '2024-01', 30000],
      ['B002', '2024-02', 32000],
      ['B002', '2024-03', 35000],
      ['C003', '2024-01', 80000],
      ['C003', '2024-02', 85000],
      ['C003', '2024-03', 90000],
      ['A001', '2024-04', 65000],
      ['B002', '2024-04', 38000],
      ['C003', '2024-04', 95000]
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
    
    console.log('📊 創建樞紐分析表...');
    
    // 創建樞紐分析表工作表
    const pivotWs = wb.getWorksheet('工作表5');
    
    // 手動創建樞紐分析表結構（模擬）
    pivotWs.setCell('A1', '樞紐分析表', { font: { bold: true, size: 16 } });
    
    // 創建樞紐分析表標題
    pivotWs.setCell('A3', 'Account', { font: { bold: true } });
    pivotWs.setCell('B3', '2024-01', { font: { bold: true } });
    pivotWs.setCell('C3', '2024-02', { font: { bold: true } });
    pivotWs.setCell('D3', '2024-03', { font: { bold: true } });
    pivotWs.setCell('E3', '2024-04', { font: { bold: true } });
    pivotWs.setCell('F3', 'Total', { font: { bold: true } });
    
    // 創建樞紐分析表資料
    const pivotData = [
      ['A001', 50000, 55000, 60000, 65000, 230000],
      ['B002', 30000, 32000, 35000, 38000, 135000],
      ['C003', 80000, 85000, 90000, 95000, 350000],
      ['Total', 160000, 172000, 185000, 198000, 715000]
    ];
    
    // 寫入樞紐分析表資料
    for (let i = 0; i < pivotData.length; i++) {
      const row = pivotData[i];
      for (let j = 0; j < row.length; j++) {
        const col = String.fromCharCode(65 + j); // A, B, C, D, E, F
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
    pivotWs.setColumnWidth('B', 15);
    pivotWs.setColumnWidth('C', 15);
    pivotWs.setColumnWidth('D', 15);
    pivotWs.setColumnWidth('E', 15);
    pivotWs.setColumnWidth('F', 15);
    
    // 添加樣式
    pivotWs.setCell('F3', 'Total', { 
      font: { bold: true },
      fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
    });
    
    // 為總計列添加樣式
    for (let col = 1; col <= 6; col++) {
      const colLetter = String.fromCharCode(64 + col);
      pivotWs.setCell(`${colLetter}7`, pivotData[3][col - 1], {
        font: { bold: true },
        fill: { type: 'pattern', patternType: 'solid', fgColor: '#F0F0F0' }
      });
    }
    
    console.log('💾 輸出 Excel 檔案...');
    
    // 輸出檔案
    const buffer = await wb.writeBuffer();
    const filename = 'test-pivot-table.xlsx';
    fs.writeFileSync(filename, new Uint8Array(buffer));
    
    console.log(`✅ Excel 檔案 ${filename} 已產生`);
    console.log('📊 檔案大小:', (buffer.byteLength / 1024).toFixed(2), 'KB');
    
    // 驗證樞紐分析表資料
    console.log('\n📋 樞紐分析表驗證:');
    console.log('工作表名稱:', pivotWs.name);
    
    // 檢查關鍵儲存格
    console.log('A1 (標題):', pivotWs.getCell('A1').value);
    console.log('A3 (Account 標題):', pivotWs.getCell('A3').value);
    console.log('B3 (2024-01 標題):', pivotWs.getCell('B3').value);
    console.log('A4 (A001):', pivotWs.getCell('A4').value);
    console.log('B4 (A001 2024-01 金額):', pivotWs.getCell('B4').value);
    console.log('F4 (A001 總計):', pivotWs.getCell('F4').value);
    
    // 驗證資料正確性
    console.log('\n🔍 資料正確性驗證:');
    
    // 檢查 A001 的總計
    const a001Total = 50000 + 55000 + 60000 + 65000;
    const actualA001Total = pivotWs.getCell('F4').value;
    console.log(`A001 總計: 預期 ${a001Total}, 實際 ${actualA001Total}`);
    
    // 檢查 2024-01 的總計
    const janTotal = 50000 + 30000 + 80000;
    const actualJanTotal = pivotWs.getCell('B7').value;
    console.log(`2024-01 總計: 預期 ${janTotal}, 實際 ${actualJanTotal}`);
    
    // 檢查整體總計
    const grandTotal = 230000 + 135000 + 350000;
    const actualGrandTotal = pivotWs.getCell('F7').value;
    console.log(`整體總計: 預期 ${grandTotal}, 實際 ${actualGrandTotal}`);
    
    console.log('\n🎯 樞紐分析表測試完成！');
    console.log('請檢查 Excel 檔案中的樞紐分析表是否正確顯示。');
    
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error('錯誤詳情:', error.stack);
  }
}

// 執行測試
testPivotTable().catch(console.error);
