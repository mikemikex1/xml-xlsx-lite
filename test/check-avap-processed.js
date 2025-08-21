const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function checkAvapProcessed() {
  console.log('🔍 檢查 AVAP 處理後的檔案');
  console.log('='.repeat(50));

  try {
    // 檢查檔案是否存在
    if (!fs.existsSync('avap-saving-report-processed.xlsx')) {
      console.log('❌ 檔案不存在: avap-saving-report-processed.xlsx');
      return;
    }

    console.log('✅ 檔案存在');
    
    // 顯示檔案資訊
    const stats = fs.statSync('avap-saving-report-processed.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);
    
    // 創建工作簿來讀取檔案
    const workbook = new Workbook();
    
    // 獲取工作表
    const worksheets = workbook.getWorksheets();
    console.log(`📊 工作表數量: ${worksheets.length}`);
    
    // 顯示工作表清單
    console.log('\n📋 工作表清單:');
    worksheets.forEach((sheet, index) => {
      console.log(`${index + 1}. ${sheet.name}`);
    });
    
    // 檢查 Detail 工作表
    const detailSheet = workbook.getWorksheet('Detail');
    if (detailSheet) {
      console.log('\n📊 Detail 工作表內容:');
      console.log('-'.repeat(30));
      
      let rowCount = 0;
      for (const [rowNum, rowMap] of detailSheet.rows()) {
        if (rowCount < 10) { // 只顯示前10行
          const rowData = [];
          for (let col = 0; col < 3; col++) {
            const cell = rowMap.get(col + 1);
            if (cell) {
              rowData.push(cell.value);
            } else {
              rowData.push('(空)');
            }
          }
          console.log(`行 ${rowNum}: [${rowData.join(', ')}]`);
        }
        rowCount++;
      }
      console.log(`... 總共 ${rowCount} 行`);
    }
    
    // 檢查工作表5
    const sheet5 = workbook.getWorksheet('工作表5');
    if (sheet5) {
      console.log('\n📊 工作表5 內容:');
      console.log('-'.repeat(30));
      
      let rowCount = 0;
      for (const [rowNum, rowMap] of sheet5.rows()) {
        if (rowCount < 10) { // 只顯示前10行
          const rowData = [];
          for (let col = 0; col < 4; col++) {
            const cell = rowMap.get(col + 1);
            if (cell) {
              rowData.push(cell.value);
            } else {
              rowData.push('(空)');
            }
          }
          console.log(`行 ${rowNum}: [${rowData.join(', ')}]`);
        }
        rowCount++;
      }
      console.log(`... 總共 ${rowCount} 行`);
    }
    
    console.log('\n🎉 檢查完成！');
    
  } catch (error) {
    console.error('❌ 檢查失敗:', error);
  }
}

// 執行檢查
checkAvapProcessed().catch(console.error);
