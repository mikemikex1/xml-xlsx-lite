const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function debugPivot() {
  console.log('🔍 調試 Pivot Table 問題');
  console.log('='.repeat(50));

  try {
    const workbook = new Workbook();
    console.log('✅ 工作簿創建成功');

    // 創建一個簡單的工作表
    const sheet = workbook.getWorksheet('測試資料');
    sheet.setCell('A1', '產品', { font: { bold: true } });
    sheet.setCell('B1', '銷售額', { font: { bold: true } });
    sheet.setCell('A2', '產品A', 1000);
    sheet.setCell('B2', 5000);
    sheet.setCell('A3', '產品B', 2000);
    sheet.setCell('B3', 8000);

    console.log('✅ 測試資料創建完成');

    // 檢查工作表內部狀態
    console.log('\n🔍 工作表內部狀態檢查:');
    console.log(`工作表名稱: ${sheet.name}`);
    console.log(`工作表保護狀態: ${sheet.isProtected()}`);

    console.log('\n📊 儲存格檢查:');
    console.log(`A1: ${sheet.getCell('A1').value}`);
    console.log(`B1: ${sheet.getCell('B1').value}`);
    console.log(`A2: ${sheet.getCell('A2').value}`);
    console.log(`B2: ${sheet.getCell('B2').value}`);
    console.log(`A3: ${sheet.getCell('A3').value}`);
    console.log(`B3: ${sheet.getCell('B3').value}`);

    console.log('\n🔍 rows() 方法檢查:');
    let rowCount = 0;
    for (const [rowNum, rowMap] of sheet.rows()) {
      console.log(`行 ${rowNum}: ${rowMap.size} 個儲存格`);
      rowCount++;
    }
    console.log(`總行數: ${rowCount}`);

    // 創建 Pivot Table
    const pivotConfig = {
      name: '簡單分析表',
      sourceRange: 'A1:B3',
      targetRange: 'D1:F10',
      fields: [
        { name: '產品', sourceColumn: '產品', type: 'row' },
        { name: '銷售額', sourceColumn: '銷售額', type: 'value', function: 'sum' }
      ]
    };

    const pivotTable = workbook.createPivotTable(pivotConfig);
    console.log('\n✅ Pivot Table 創建成功');

    // 匯出到新工作表
    const exportSheet = pivotTable.exportToWorksheet('Pivot匯出');
    console.log('✅ Pivot Table 匯出成功');

    // 檢查匯出工作表
    console.log('\n🔍 匯出工作表檢查:');
    console.log(`匯出工作表名稱: ${exportSheet.name}`);
    let exportRowCount = 0;
    for (const [rowNum, rowMap] of exportSheet.rows()) {
      console.log(`匯出行 ${rowNum}: ${rowMap.size} 個儲存格`);
      exportRowCount++;
    }
    console.log(`匯出總行數: ${exportRowCount}`);

    console.log('\n🔍 工作簿狀態檢查:');
    console.log(`工作表數量: ${workbook.getWorksheets().length}`);
    const sheetNames = workbook.getWorksheets().map(ws => ws.name);
    console.log(`工作表名稱: ${sheetNames.join(', ')}`);

    console.log('\n💾 生成 Excel 檔案...');
    const buffer = await workbook.writeBufferWithPivotTables();
    fs.writeFileSync('debug-pivot.xlsx', new Uint8Array(buffer));
    console.log('✅ Excel 檔案已生成: debug-pivot.xlsx');
    const stats = fs.statSync('debug-pivot.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    console.log('\n🎉 調試完成！');

  } catch (error) {
    console.error('❌ 調試失敗:', error);
    console.error(error.stack);
  }
}

debugPivot();
