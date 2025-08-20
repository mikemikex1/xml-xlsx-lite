const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testPhase3() {
  console.log('🧪 測試 Phase 3: 進階功能');
  
  const wb = new Workbook();
  const ws = wb.getWorksheet('進階功能測試');
  
  // 測試合併儲存格
  console.log('🔗 測試合併儲存格...');
  ws.setCell('A1', '合併標題', {
    font: { bold: true, size: 16 },
    alignment: { horizontal: 'center' }
  });
  ws.mergeCells('A1:C1');
  
  ws.setCell('A2', '左側標題', {
    font: { bold: true },
    alignment: { vertical: 'middle' }
  });
  ws.mergeCells('A2:A4');
  
  // 測試欄寬/列高設定
  console.log('📏 測試欄寬/列高設定...');
  ws.setColumnWidth('A', 15);
  ws.setColumnWidth('B', 20);
  ws.setColumnWidth('C', 25);
  ws.setColumnWidth('D', 30);
  
  ws.setRowHeight(1, 30);
  ws.setRowHeight(2, 25);
  ws.setRowHeight(3, 25);
  ws.setRowHeight(4, 25);
  
  // 測試凍結窗格
  console.log('❄️ 測試凍結窗格...');
  ws.freezePanes(1, 1); // 凍結第一行和第一列
  
  // 填充一些資料
  console.log('📊 填充測試資料...');
  ws.setCell('B2', '欄位 1', {
    font: { bold: true },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
  });
  ws.setCell('C2', '欄位 2', {
    font: { bold: true },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
  });
  ws.setCell('D2', '欄位 3', {
    font: { bold: true },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' }
  });
  
  ws.setCell('B3', '資料 1-1');
  ws.setCell('C3', '資料 1-2');
  ws.setCell('D3', '資料 1-3');
  
  ws.setCell('B4', '資料 2-1');
  ws.setCell('C4', '資料 2-2');
  ws.setCell('D4', '資料 2-3');
  
  // 測試邊框樣式
  console.log('🔲 添加邊框樣式...');
  ws.setCell('B2', '欄位 1', {
    font: { bold: true },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' },
    border: {
      top: { style: 'thin', color: '#000000' },
      bottom: { style: 'thin', color: '#000000' },
      left: { style: 'thin', color: '#000000' },
      right: { style: 'thin', color: '#000000' }
    }
  });
  
  ws.setCell('C2', '欄位 2', {
    font: { bold: true },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' },
    border: {
      top: { style: 'thin', color: '#000000' },
      bottom: { style: 'thin', color: '#000000' },
      left: { style: 'thin', color: '#000000' },
      right: { style: 'thin', color: '#000000' }
    }
  });
  
  ws.setCell('D2', '欄位 3', {
    font: { bold: true },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#E6F3FF' },
    border: {
      top: { style: 'thin', color: '#000000' },
      bottom: { style: 'thin', color: '#000000' },
      left: { style: 'thin', color: '#000000' },
      right: { style: 'thin', color: '#000000' }
    }
  });
  
  // 顯示工作表資訊
  console.log('\n📊 工作表資訊:');
  console.log('合併範圍:', ws.getMergedRanges());
  console.log('凍結窗格:', ws.getFreezePanes());
  console.log('A 欄寬度:', ws.getColumnWidth('A'));
  console.log('B 欄寬度:', ws.getColumnWidth('B'));
  console.log('C 欄寬度:', ws.getColumnWidth('C'));
  console.log('D 欄寬度:', ws.getColumnWidth('D'));
  console.log('第 1 列高度:', ws.getRowHeight(1));
  console.log('第 2 列高度:', ws.getRowHeight(2));
  
  // 生成 Excel 檔案
  console.log('\n💾 生成 Excel 檔案...');
  const buffer = await wb.writeBuffer();
  
  const filename = 'test-phase3.xlsx';
  fs.writeFileSync(filename, Buffer.from(buffer));
  console.log(`✅ Phase 3 測試完成！檔案已儲存為: ${filename}`);
  
  // 顯示儲存格資訊
  console.log('\n📋 儲存格詳細資訊:');
  for (const [r, rowMap] of ws.rows()) {
    for (const [c, cell] of rowMap) {
      const addr = cell.address;
      const value = cell.value;
      const hasStyle = cell.options.font || cell.options.fill || cell.options.border || cell.options.alignment;
      const isMerged = cell.options.mergeRange || cell.options.mergedInto;
      console.log(`  ${addr}: ${value} ${hasStyle ? '✨' : ''} ${isMerged ? '🔗' : ''}`);
    }
  }
}

// 執行測試
testPhase3().catch(console.error);
