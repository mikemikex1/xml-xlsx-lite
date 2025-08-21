const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testPivotOnly() {
  console.log('🎯 測試 Pivot Table 功能 - 使用 writeBufferWithPivotTables');
  console.log('='.repeat(60));

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
    console.log('✅ Pivot Table 創建成功');

    // 匯出到新工作表
    const exportSheet = pivotTable.exportToWorksheet('Pivot匯出');
    console.log('✅ Pivot Table 匯出成功');

    console.log('\n💾 使用 writeBufferWithPivotTables 生成檔案...');
    const buffer = await workbook.writeBufferWithPivotTables();
    fs.writeFileSync('test-pivot-only.xlsx', new Uint8Array(buffer));
    console.log('✅ Excel 檔案已生成: test-pivot-only.xlsx');

    const stats = fs.statSync('test-pivot-only.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    console.log('\n🎉 Pivot Table 測試完成！');

  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error(error.stack);
  }
}

testPivotOnly();
