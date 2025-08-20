const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testFormulas() {
  console.log('🧮 測試 Phase 3: 公式支援');
  
  const wb = new Workbook();
  const ws = wb.getWorksheet('公式測試');
  
  // 測試基本數學公式
  console.log('🔢 測試基本數學公式...');
  
  // 設定一些數值資料
  ws.setCell('A1', 10);
  ws.setCell('A2', 20);
  ws.setCell('A3', 30);
  ws.setCell('A4', 40);
  ws.setCell('A5', 50);
  
  // 設定公式
  ws.setFormula('A6', '=SUM(A1:A5)');
  ws.setFormula('A7', '=AVERAGE(A1:A5)');
  ws.setFormula('A8', '=MAX(A1:A5)');
  ws.setFormula('A9', '=MIN(A1:A5)');
  ws.setFormula('A10', '=COUNT(A1:A5)');
  
  // 測試邏輯公式
  console.log('🧠 測試邏輯公式...');
  ws.setCell('B1', 100);
  ws.setCell('B2', 200);
  ws.setFormula('B3', '=IF(B1>B2,"B1 較大","B2 較大")');
  ws.setFormula('B4', '=AND(B1>50,B2>150)');
  ws.setFormula('B5', '=OR(B1>150,B2>250)');
  
  // 測試文字公式
  console.log('📝 測試文字公式...');
  ws.setCell('C1', 'Hello');
  ws.setCell('C2', 'World');
  ws.setFormula('C3', '=CONCATENATE(C1," ",C2)');
  ws.setFormula('C4', '=LEFT(C1,3)');
  ws.setFormula('C5', '=RIGHT(C2,3)');
  ws.setFormula('C6', '=MID(C1,2,3)');
  
  // 測試日期公式
  console.log('📅 測試日期公式...');
  ws.setFormula('D1', '=TODAY()');
  ws.setFormula('D2', '=NOW()');
  ws.setFormula('D3', '=DATE(2024,1,1)');
  ws.setFormula('D4', '=YEAR(D1)');
  ws.setFormula('D5', '=MONTH(D1)');
  ws.setFormula('D6', '=DAY(D1)');
  
  // 測試查找公式
  console.log('🔍 測試查找公式...');
  ws.setCell('E1', '蘋果');
  ws.setCell('E2', '香蕉');
  ws.setCell('E3', '橙子');
  ws.setCell('F1', 5);
  ws.setCell('F2', 3);
  ws.setCell('F3', 7);
  ws.setFormula('E4', '=VLOOKUP("香蕉",E1:F3,2,FALSE)');
  ws.setFormula('E5', '=INDEX(F1:F3,2)');
  ws.setFormula('E6', '=MATCH("橙子",E1:E3,0)');
  
  // 測試複雜公式
  console.log('🎯 測試複雜公式...');
  ws.setFormula('G1', '=SUMIF(A1:A5,">25")');
  ws.setFormula('G2', '=COUNTIF(A1:A5,">30")');
  ws.setFormula('G3', '=IF(AND(A1>5,A2>15),"條件成立","條件不成立")');
  ws.setFormula('G4', '=SUM(A1:A5)*2+10');
  ws.setFormula('G5', '=ROUND(AVERAGE(A1:A5),2)');
  
  // 設定欄寬和樣式
  console.log('🎨 設定樣式和欄寬...');
  ws.setColumnWidth('A', 15);
  ws.setColumnWidth('B', 15);
  ws.setColumnWidth('C', 20);
  ws.setColumnWidth('D', 15);
  ws.setColumnWidth('E', 15);
  ws.setColumnWidth('F', 15);
  ws.setColumnWidth('G', 20);
  
  // 設定標題樣式
  ws.setCell('A1', '數值', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' },
    alignment: { horizontal: 'center' }
  });
  
  ws.setCell('B1', '邏輯測試', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' },
    alignment: { horizontal: 'center' }
  });
  
  ws.setCell('C1', '文字處理', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' },
    alignment: { horizontal: 'center' }
  });
  
  ws.setCell('D1', '日期函數', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' },
    alignment: { horizontal: 'center' }
  });
  
  ws.setCell('E1', '查找函數', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' },
    alignment: { horizontal: 'center' }
  });
  
  ws.setCell('G1', '複雜公式', {
    font: { bold: true, size: 14 },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' },
    alignment: { horizontal: 'center' }
  });
  
  // 顯示工作表資訊
  console.log('\n📊 工作表資訊:');
  console.log('A6 公式:', ws.getFormula('A6'));
  console.log('B3 公式:', ws.getFormula('B3'));
  console.log('C3 公式:', ws.getFormula('C3'));
  console.log('D1 公式:', ws.getFormula('D1'));
  console.log('E4 公式:', ws.getFormula('E4'));
  console.log('G1 公式:', ws.getFormula('G1'));
  
  // 生成 Excel 檔案
  console.log('\n💾 生成 Excel 檔案...');
  const buffer = await wb.writeBuffer();
  
  const filename = 'test-formulas.xlsx';
  fs.writeFileSync(filename, Buffer.from(buffer));
  console.log(`✅ 公式測試完成！檔案已儲存為: ${filename}`);
  
  // 顯示儲存格資訊
  console.log('\n📋 儲存格詳細資訊:');
  for (const [r, rowMap] of ws.rows()) {
    for (const [c, cell] of rowMap) {
      const addr = cell.address;
      const value = cell.value;
      const formula = cell.options.formula;
      const hasStyle = cell.options.font || cell.options.fill || cell.options.border || cell.options.alignment;
      console.log(`  ${addr}: ${value} ${formula ? `[公式: ${formula}]` : ''} ${hasStyle ? '✨' : ''}`);
    }
  }
}

// 執行測試
testFormulas().catch(console.error);
