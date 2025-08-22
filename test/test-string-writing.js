/**
 * 測試字串寫入功能
 * 驗證 inlineStr 支援是否正常工作
 */

const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testStringWriting() {
  console.log('🧪 測試字串寫入功能...');
  
  try {
    // 創建工作簿
    const wb = new Workbook();
    
    // 創建測試工作表
    const ws = wb.getWorksheet('String Test');
    
    console.log('📝 寫入各種類型的資料...');
    
    // 測試數字
    ws.setCell('A1', '數字測試', { font: { bold: true } });
    ws.setCell('A2', 123);
    ws.setCell('A3', 456.78);
    ws.setCell('A4', -999);
    
    // 測試字串（這是關鍵測試）
    ws.setCell('B1', '字串測試', { font: { bold: true } });
    ws.setCell('B2', 'Hello World');
    ws.setCell('B3', '繁體中文測試');
    ws.setCell('B4', 'Emoji 測試 🚀🎉💻');
    ws.setCell('B5', '包含空格的字串 ');
    ws.setCell('B6', ' 前後都有空格 ');
    ws.setCell('B7', ''); // 空字串
    ws.setCell('B8', '特殊字符: & < > " \'');
    
    // 測試布林值
    ws.setCell('C1', '布林值測試', { font: { bold: true } });
    ws.setCell('C2', true);
    ws.setCell('C3', false);
    
    // 測試日期
    ws.setCell('D1', '日期測試', { font: { bold: true } });
    ws.setCell('D2', new Date('2024-01-01'));
    ws.setCell('D3', new Date('2024-12-31'));
    
    // 測試混合資料
    ws.setCell('E1', '混合資料測試', { font: { bold: true } });
    ws.setCell('E2', '部門');
    ws.setCell('E3', '姓名');
    ws.setCell('E4', '月份');
    ws.setCell('E5', '銷售額');
    
    ws.setCell('F2', 'A');
    ws.setCell('F3', '小明');
    ws.setCell('F4', '1月');
    ws.setCell('F5', 1000);
    
    ws.setCell('G2', 'B');
    ws.setCell('G3', '小華');
    ws.setCell('G4', '2月');
    ws.setCell('G5', 2000);
    
    // 設定欄寬
    ws.setColumnWidth('A', 15);
    ws.setColumnWidth('B', 20);
    ws.setColumnWidth('C', 15);
    ws.setColumnWidth('D', 15);
    ws.setColumnWidth('E', 15);
    ws.setColumnWidth('F', 10);
    ws.setColumnWidth('G', 10);
    
    console.log('💾 輸出 Excel 檔案...');
    
    // 使用 writeBuffer 方法
    const buffer = await wb.writeBuffer();
    const filename = 'test-string-writing.xlsx';
    fs.writeFileSync(filename, new Uint8Array(buffer));
    
    console.log(`✅ Excel 檔案 ${filename} 已產生`);
    console.log('📊 檔案大小:', (buffer.byteLength / 1024).toFixed(2), 'KB');
    
    // 驗證工作表內容
    console.log('\n📋 工作表內容驗證:');
    console.log('工作表名稱:', ws.name);
    console.log('儲存格 A1:', ws.getCell('A1').value);
    console.log('儲存格 B2:', ws.getCell('B2').value);
    console.log('儲存格 B3:', ws.getCell('B3').value);
    console.log('儲存格 B4:', ws.getCell('B4').value);
    console.log('儲存格 B5:', ws.getCell('B5').value);
    console.log('儲存格 B6:', ws.getCell('B6').value);
    console.log('儲存格 B7:', ws.getCell('B7').value);
    console.log('儲存格 B8:', ws.getCell('B8').value);
    
    console.log('\n🎯 測試完成！請檢查 Excel 檔案中的字串是否正常顯示。');
    
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error('錯誤詳情:', error.stack);
  }
}

// 執行測試
testStringWriting().catch(console.error);
