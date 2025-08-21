const XLSX = require('xlsx');
const fs = require('fs');

function checkPivotOnly() {
  console.log('🔍 檢查 Pivot Table 測試 Excel 檔案');
  console.log('='.repeat(50));
  try {
    if (!fs.existsSync('test-pivot-only.xlsx')) {
      console.log('❌ 檔案不存在: test-pivot-only.xlsx');
      return;
    }
    console.log('✅ 檔案存在: test-pivot-only.xlsx');
    const stats = fs.statSync('test-pivot-only.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);
    
    const workbook = XLSX.readFile('test-pivot-only.xlsx');
    console.log('✅ Excel 檔案讀取成功');
    
    const sheetNames = workbook.SheetNames;
    console.log(`📋 工作表數量: ${sheetNames.length}`);
    console.log('📋 工作表名稱:', sheetNames);
    
    for (const sheetName of sheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      console.log(`\n📊 ${sheetName}: ${data.length} 行資料`);
      for (let i = 0; i < Math.min(data.length, 5); i++) {
        console.log(`  行 ${i + 1}:`, data[i]);
      }
      if (data.length > 5) {
        console.log(`  ... 還有 ${data.length - 5} 行`);
      }
    }
    
    console.log('\n🎉 檢查完成！');
  } catch (error) {
    console.error('❌ 檢查失敗:', error);
  }
}

checkPivotOnly();
