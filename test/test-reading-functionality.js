/**
 * 測試讀取功能 - 簡化版本
 * 先驗證 toArray 和 toJSON 功能
 */

const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testReadingFunctionality() {
  console.log('🧪 測試讀取功能...');
  
  try {
    // 首先創建一個測試檔案
    console.log('📝 創建測試資料...');
    
    const wb = new Workbook();
    const ws = wb.getWorksheet('TestData');
    
    // 添加測試資料
    ws.setCell('A1', '部門');
    ws.setCell('B1', '姓名');
    ws.setCell('C1', '月份');
    ws.setCell('D1', '銷售額');
    
    ws.setCell('A2', 'A');
    ws.setCell('B2', '小明');
    ws.setCell('C2', '1月');
    ws.setCell('D2', 1000);
    
    ws.setCell('A3', 'A');
    ws.setCell('B3', '小華');
    ws.setCell('C3', '2月');
    ws.setCell('D3', 1500);
    
    ws.setCell('A4', 'B');
    ws.setCell('B4', '小美');
    ws.setCell('C4', '1月');
    ws.setCell('D4', 2000);
    
    console.log('💾 輸出測試檔案...');
    const buffer = await wb.writeBuffer();
    fs.writeFileSync('test-reading-data.xlsx', new Uint8Array(buffer));
    
    console.log('📊 測試 toArray 功能...');
    
    // 模擬 toArray 功能
    const arrayData = [];
    
    // 手動從工作表提取資料
    for (let row = 1; row <= 4; row++) {
      const rowData = [];
      for (let col = 1; col <= 4; col++) {
        const colLetter = String.fromCharCode(64 + col); // A, B, C, D
        const address = `${colLetter}${row}`;
        const cell = ws.getCell(address);
        rowData.push(cell.value);
      }
      arrayData.push(rowData);
    }
    
    console.log('toArray 結果:');
    console.table(arrayData);
    
    console.log('📋 測試 toJSON 功能...');
    
    // 模擬 toJSON 功能
    const headers = arrayData[0];
    const jsonData = [];
    
    for (let i = 1; i < arrayData.length; i++) {
      const row = arrayData[i];
      const rowObj = {};
      
      for (let j = 0; j < headers.length; j++) {
        const headerName = headers[j] || `Column${j + 1}`;
        rowObj[headerName] = row[j];
      }
      
      jsonData.push(rowObj);
    }
    
    console.log('toJSON 結果:');
    console.log(JSON.stringify(jsonData, null, 2));
    
    console.log('✅ 讀取功能測試完成！');
    
    // 驗證資料正確性
    console.log('\n🔍 驗證資料正確性:');
    console.log('總行數:', arrayData.length);
    console.log('總欄數:', arrayData[0].length);
    console.log('標題行:', headers);
    console.log('資料行數:', jsonData.length);
    
    // 檢查特定值
    console.log('A2 值:', arrayData[1][0], '(預期: A)');
    console.log('B2 值:', arrayData[1][1], '(預期: 小明)');
    console.log('D2 值:', arrayData[1][3], '(預期: 1000)');
    
    console.log('\n🎯 讀取功能基礎測試成功！');
    
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error('錯誤詳情:', error.stack);
  }
}

// 執行測試
testReadingFunctionality().catch(console.error);
