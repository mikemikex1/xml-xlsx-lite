const XLSX = require('xlsx');
const fs = require('fs');

function checkStepByStep() {
  console.log('🔍 檢查逐步測試 Excel 檔案');
  console.log('='.repeat(50));

  try {
    // 檢查檔案是否存在
    if (!fs.existsSync('test-step-by-step.xlsx')) {
      console.log('❌ 檔案不存在: test-step-by-step.xlsx');
      return;
    }

    console.log('✅ 檔案存在: test-step-by-step.xlsx');
    
    // 檢查檔案大小
    const stats = fs.statSync('test-step-by-step.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    // 讀取 Excel 檔案
    const workbook = XLSX.readFile('test-step-by-step.xlsx');
    console.log('✅ Excel 檔案讀取成功');

    // 檢查工作表
    const sheetNames = workbook.SheetNames;
    console.log(`📋 工作表數量: ${sheetNames.length}`);
    console.log('📋 工作表名稱:', sheetNames);

    // 檢查每個工作表
    for (const sheetName of sheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      console.log(`\n📊 ${sheetName}: ${data.length} 行資料`);
      
      // 顯示前幾行資料
      for (let i = 0; i < Math.min(data.length, 3); i++) {
        console.log(`  行 ${i + 1}:`, data[i]);
      }
      if (data.length > 3) {
        console.log(`  ... 還有 ${data.length - 3} 行`);
      }
    }

    console.log('\n🎉 檢查完成！');

  } catch (error) {
    console.error('❌ 檢查失敗:', error);
  }
}

checkStepByStep();
