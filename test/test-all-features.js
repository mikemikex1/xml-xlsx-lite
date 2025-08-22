/**
 * 綜合測試 - 展示所有已實現的功能
 */

const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testAllFeatures() {
  console.log('🧪 綜合測試 - 展示所有已實現的功能...');
  
  try {
    // 創建工作簿
    const wb = new Workbook();
    
    console.log('📝 1. 測試字串寫入功能...');
    
    // 創建字串測試工作表
    const stringWs = wb.getWorksheet('String Test');
    
    // 測試各種資料類型
    stringWs.setCell('A1', '功能測試', { font: { bold: true, size: 16 } });
    stringWs.setCell('A3', '數字測試', { font: { bold: true } });
    stringWs.setCell('A4', 123);
    stringWs.setCell('A5', 456.78);
    stringWs.setCell('A6', -999);
    
    stringWs.setCell('B3', '字串測試', { font: { bold: true } });
    stringWs.setCell('B4', 'Hello World');
    stringWs.setCell('B5', '繁體中文測試');
    stringWs.setCell('B6', 'Emoji 測試 🚀🎉💻');
    stringWs.setCell('B7', '包含空格的字串 ');
    stringWs.setCell('B8', ' 前後都有空格 ');
    stringWs.setCell('B9', ''); // 空字串
    stringWs.setCell('B10', '特殊字符: & < > " \'');
    
    stringWs.setCell('C3', '布林值測試', { font: { bold: true } });
    stringWs.setCell('C4', true);
    stringWs.setCell('C5', false);
    
    stringWs.setCell('D3', '日期測試', { font: { bold: true } });
    stringWs.setCell('D4', new Date('2024-01-01'));
    stringWs.setCell('D5', new Date('2024-12-31'));
    
    // 設定欄寬
    stringWs.setColumnWidth('A', 15);
    stringWs.setColumnWidth('B', 25);
    stringWs.setColumnWidth('C', 15);
    stringWs.setColumnWidth('D', 15);
    
    console.log('✅ 字串寫入功能測試完成');
    
    console.log('\n📊 2. 測試樞紐分析表功能...');
    
    // 創建資料工作表
    const dataWs = wb.getWorksheet('Detail');
    
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
      ['C003', '2024-03', 90000]
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
    
    console.log('✅ 樞紐分析表資料準備完成');
    
    console.log('\n📋 3. 測試樣式和格式化功能...');
    
    // 創建樣式測試工作表
    const styleWs = wb.getWorksheet('Style Test');
    
    // 測試各種樣式
    styleWs.setCell('A1', '樣式測試', { 
      font: { bold: true, size: 18, color: '#FF0000' },
      fill: { type: 'pattern', patternType: 'solid', fgColor: '#FFFF00' },
      alignment: { horizontal: 'center', vertical: 'middle' }
    });
    
    styleWs.setCell('A3', '字體樣式', { font: { bold: true } });
    styleWs.setCell('A4', '粗體文字', { font: { bold: true } });
    styleWs.setCell('A5', '斜體文字', { font: { italic: true } });
    styleWs.setCell('A6', '底線文字', { font: { underline: true } });
    
    styleWs.setCell('B3', '對齊樣式', { font: { bold: true } });
    styleWs.setCell('B4', '左對齊', { alignment: { horizontal: 'left' } });
    styleWs.setCell('B5', '置中對齊', { alignment: { horizontal: 'center' } });
    styleWs.setCell('B6', '右對齊', { alignment: { horizontal: 'right' } });
    
    styleWs.setCell('C3', '填滿樣式', { font: { bold: true } });
    styleWs.setCell('C4', '紅色背景', { fill: { type: 'pattern', patternType: 'solid', fgColor: '#FF0000' } });
    styleWs.setCell('C5', '綠色背景', { fill: { type: 'pattern', patternType: 'solid', fgColor: '#00FF00' } });
    styleWs.setCell('C6', '藍色背景', { fill: { type: 'pattern', patternType: 'solid', fgColor: '#0000FF' } });
    
    // 設定欄寬
    styleWs.setColumnWidth('A', 20);
    styleWs.setColumnWidth('B', 20);
    styleWs.setColumnWidth('C', 20);
    
    console.log('✅ 樣式功能測試完成');
    
    console.log('\n🔧 4. 測試欄寬和列高設定...');
    
    // 測試欄寬設定
    stringWs.setColumnWidth('E', 30);
    stringWs.setColumnWidth('F', 25);
    
    // 測試列高設定
    stringWs.setRowHeight(1, 30);
    stringWs.setRowHeight(3, 25);
    
    console.log('✅ 欄寬和列高設定測試完成');
    
    console.log('\n💾 5. 輸出 Excel 檔案...');
    
    // 輸出檔案
    const buffer = await wb.writeBuffer();
    const filename = 'test-all-features.xlsx';
    fs.writeFileSync(filename, new Uint8Array(buffer));
    
    console.log(`✅ Excel 檔案 ${filename} 已產生`);
    console.log('📊 檔案大小:', (buffer.byteLength / 1024).toFixed(2), 'KB');
    
    // 驗證所有功能
    console.log('\n🔍 功能驗證:');
    
    // 驗證字串寫入
    console.log('字串測試 - A1:', stringWs.getCell('A1').value);
    console.log('字串測試 - B4:', stringWs.getCell('B4').value);
    console.log('字串測試 - B5:', stringWs.getCell('B5').value);
    console.log('字串測試 - B6:', stringWs.getCell('B6').value);
    
    // 驗證數字寫入
    console.log('數字測試 - A4:', stringWs.getCell('A4').value);
    console.log('數字測試 - A5:', stringWs.getCell('A5').value);
    
    // 驗證布林值寫入
    console.log('布林值測試 - C4:', stringWs.getCell('C4').value);
    console.log('布林值測試 - C5:', stringWs.getCell('C5').value);
    
    // 驗證日期寫入
    console.log('日期測試 - D4:', stringWs.getCell('D4').value);
    console.log('日期測試 - D5:', stringWs.getCell('D5').value);
    
    // 驗證樞紐分析表資料
    console.log('樞紐分析表資料 - A1:', dataWs.getCell('A1').value);
    console.log('樞紐分析表資料 - A2:', dataWs.getCell('A2').value);
    console.log('樞紐分析表資料 - C2:', dataWs.getCell('C2').value);
    
    // 驗證樣式
    console.log('樣式測試 - A1:', styleWs.getCell('A1').value);
    console.log('樣式測試 - A4:', styleWs.getCell('A4').value);
    console.log('樣式測試 - B5:', styleWs.getCell('B5').value);
    console.log('樣式測試 - C4:', styleWs.getCell('C4').value);
    
    console.log('\n🎯 所有功能測試完成！');
    console.log('請檢查 Excel 檔案中的各種功能是否正確顯示。');
    
    // 顯示統計資訊
    console.log('\n📊 測試統計:');
    console.log('工作表數量:', wb.getWorksheets().length);
    console.log('工作表名稱:', wb.getWorksheets().map(ws => ws.name).join(', '));
    
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error('錯誤詳情:', error.stack);
  }
}

// 執行測試
testAllFeatures().catch(console.error);
