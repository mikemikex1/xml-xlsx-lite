const XLSX = require('xlsx');
const fs = require('fs');

function checkDebug() {
  console.log('🔍 檢查調試測試 Excel 檔案');
  console.log('='.repeat(50));

  try {
    // 檢查檔案是否存在
    if (!fs.existsSync('test-debug.xlsx')) {
      console.log('❌ 檔案不存在: test-debug.xlsx');
      return;
    }

    console.log('✅ 檔案存在: test-debug.xlsx');
    
    // 檢查檔案大小
    const stats = fs.statSync('test-debug.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    // 讀取 Excel 檔案
    const workbook = XLSX.readFile('test-debug.xlsx');
    console.log('✅ Excel 檔案讀取成功');

    // 檢查工作表
    const sheetNames = workbook.SheetNames;
    console.log(`📋 工作表數量: ${sheetNames.length}`);
    console.log('📋 工作表名稱:', sheetNames);

    // 檢查調試工作表
    if (workbook.Sheets['調試']) {
      const debugData = XLSX.utils.sheet_to_json(workbook.Sheets['調試'], { header: 1 });
      console.log(`✅ 調試工作表: ${debugData.length} 行資料`);
      
      // 顯示資料
      for (let i = 0; i < debugData.length; i++) {
        console.log(`  行 ${i + 1}:`, debugData[i]);
      }
    } else {
      console.log('❌ 調試工作表不存在');
    }

    console.log('\n🎉 檢查完成！');

  } catch (error) {
    console.error('❌ 檢查失敗:', error);
  }
}

checkDebug();
