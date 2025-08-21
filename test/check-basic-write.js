const XLSX = require('xlsx');
const fs = require('fs');

function checkBasicWrite() {
  console.log('🔍 檢查基本寫入測試 Excel 檔案');
  console.log('='.repeat(50));

  try {
    // 檢查檔案是否存在
    if (!fs.existsSync('test-basic-write.xlsx')) {
      console.log('❌ 檔案不存在: test-basic-write.xlsx');
      return;
    }

    console.log('✅ 檔案存在: test-basic-write.xlsx');
    
    // 檢查檔案大小
    const stats = fs.statSync('test-basic-write.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    // 讀取 Excel 檔案
    const workbook = XLSX.readFile('test-basic-write.xlsx');
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
      
      // 顯示資料
      for (let i = 0; i < data.length; i++) {
        console.log(`  行 ${i + 1}:`, data[i]);
      }
    }

    console.log('\n🎉 檢查完成！');

  } catch (error) {
    console.error('❌ 檢查失敗:', error);
  }
}

checkBasicWrite();
