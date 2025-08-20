const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testStyles() {
  console.log('🧪 測試 Phase 2: 樣式支援');
  
  const wb = new Workbook();
  const ws = wb.getWorksheet('樣式測試');
  
  // 測試字體樣式
  console.log('📝 測試字體樣式...');
  ws.setCell('A1', '標題', {
    font: {
      bold: true,
      size: 16,
      name: '微軟正黑體',
      color: '#FF0000'
    }
  });
  
  ws.setCell('A2', '斜體文字', {
    font: {
      italic: true,
      size: 14,
      color: '#0000FF'
    }
  });
  
  ws.setCell('A3', '底線文字', {
    font: {
      underline: true,
      strike: true
    }
  });
  
  // 測試對齊樣式
  console.log('📐 測試對齊樣式...');
  ws.setCell('B1', '左對齊', {
    alignment: {
      horizontal: 'left',
      vertical: 'top'
    }
  });
  
  ws.setCell('B2', '置中對齊', {
    alignment: {
      horizontal: 'center',
      vertical: 'middle'
    }
  });
  
  ws.setCell('B3', '右對齊', {
    alignment: {
      horizontal: 'right',
      vertical: 'bottom'
    }
  });
  
  ws.setCell('B4', '自動換行文字\n第二行\n第三行', {
    alignment: {
      horizontal: 'left',
      vertical: 'top',
      wrapText: true
    }
  });
  
  // 測試填滿樣式
  console.log('🎨 測試填滿樣式...');
  ws.setCell('C1', '紅色背景', {
    fill: {
      type: 'pattern',
      patternType: 'solid',
      fgColor: '#FF0000'
    }
  });
  
  ws.setCell('C2', '藍色背景', {
    fill: {
      type: 'pattern',
      patternType: 'solid',
      fgColor: '#0000FF'
    }
  });
  
  ws.setCell('C3', '網格圖案', {
    fill: {
      type: 'pattern',
      patternType: 'lightGrid',
      fgColor: '#FFFF00',
      bgColor: '#FFFFFF'
    }
  });
  
  // 測試邊框樣式
  console.log('🔲 測試邊框樣式...');
  ws.setCell('D1', '粗邊框', {
    border: {
      top: { style: 'thick', color: '#000000' },
      bottom: { style: 'thick', color: '#000000' },
      left: { style: 'thick', color: '#000000' },
      right: { style: 'thick', color: '#000000' }
    }
  });
  
  ws.setCell('D2', '虛線邊框', {
    border: {
      top: { style: 'dashed', color: '#FF0000' },
      bottom: { style: 'dotted', color: '#00FF00' }
    }
  });
  
  ws.setCell('D3', '雙線邊框', {
    border: {
      style: 'double',
      color: '#0000FF'
    }
  });
  
  // 測試組合樣式
  console.log('🎭 測試組合樣式...');
  ws.setCell('E1', '完整樣式', {
    font: {
      bold: true,
      italic: true,
      size: 18,
      color: '#FFFFFF'
    },
    fill: {
      type: 'pattern',
      patternType: 'solid',
      fgColor: '#000000'
    },
    border: {
      top: { style: 'thick', color: '#FF0000' },
      bottom: { style: 'thick', color: '#FF0000' },
      left: { style: 'thick', color: '#FF0000' },
      right: { style: 'thick', color: '#FF0000' }
    },
    alignment: {
      horizontal: 'center',
      vertical: 'middle'
    }
  });
  
  // 測試數字格式
  console.log('🔢 測試數字格式...');
  ws.setCell('F1', 1234.56, { numFmt: '#,##0.00' });
  ws.setCell('F2', 0.123, { numFmt: '0.00%' });
  ws.setCell('F3', new Date(), { numFmt: 'yyyy-mm-dd' });
  
  console.log('💾 生成 Excel 檔案...');
  const buffer = await wb.writeBuffer();
  
  const filename = 'test-styles.xlsx';
  fs.writeFileSync(filename, Buffer.from(buffer));
  console.log(`✅ 樣式測試完成！檔案已儲存為: ${filename}`);
  
  // 顯示工作表資訊
  console.log('\n📊 工作表資訊:');
  for (const [r, rowMap] of ws.rows()) {
    for (const [c, cell] of rowMap) {
      const addr = cell.address;
      const value = cell.value;
      const hasStyle = cell.options.font || cell.options.fill || cell.options.border || cell.options.alignment;
      console.log(`  ${addr}: ${value} ${hasStyle ? '✨' : ''}`);
    }
  }
}

// 執行測試
testStyles().catch(console.error);
