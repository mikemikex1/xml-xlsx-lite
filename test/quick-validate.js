const XLSX = require('xlsx');
const fs = require('fs');

// 快速驗證函數
function quickValidate(filePath, description) {
  try {
    if (!fs.existsSync(filePath)) {
      console.log(`❌ ${description}: 檔案不存在`);
      return false;
    }
    
    const workbook = XLSX.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    
    console.log(`✅ ${description}: ${sheetNames.length} 個工作表`);
    return true;
  } catch (error) {
    console.log(`❌ ${description}: ${error.message}`);
    return false;
  }
}

// 主驗證函數
function quickValidateAll() {
  console.log('🚀 快速驗證所有 Excel 檔案');
  console.log('=' .repeat(50));
  
  const files = [
    { path: 'test-basic.xlsx', desc: 'Phase 1: 基本功能' },
    { path: 'test-styles.xlsx', desc: 'Phase 2: 樣式支援' },
    { path: 'test-phase3.xlsx', desc: 'Phase 3: 進階功能' },
    { path: 'test-pivot-table.xlsx', desc: 'Phase 5: Pivot Table' },
    { path: 'test-phase6.xlsx', desc: 'Phase 6: 保護和圖表' },
    { path: 'comprehensive-test.xlsx', desc: '綜合功能測試' }
  ];
  
  let passed = 0;
  let total = files.length;
  
  files.forEach(file => {
    if (quickValidate(file.path, file.desc)) {
      passed++;
    }
  });
  
  console.log('\n' + '=' .repeat(50));
  console.log(`📊 快速驗證結果: ${passed}/${total} 通過`);
  
  if (passed === total) {
    console.log('🎉 所有檔案驗證通過！');
  } else {
    console.log('⚠️ 部分檔案驗證失敗');
  }
}

// 執行快速驗證
quickValidateAll();
