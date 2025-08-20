const { Workbook, PivotField, PivotTableConfig } = require('../dist/index.js');
const fs = require('fs');

async function testPivotTable() {
  console.log('🎯 測試 Phase 5: Pivot Table 支援');
  
  // 建立工作簿
  const wb = new Workbook();
  
  // 建立資料工作表
  console.log('📊 建立資料工作表...');
  const dataWs = wb.getWorksheet('銷售資料');
  
  // 設定標題
  dataWs.setCell('A1', '產品', { font: { bold: true } });
  dataWs.setCell('B1', '地區', { font: { bold: true } });
  dataWs.setCell('C1', '月份', { font: { bold: true } });
  dataWs.setCell('D1', '銷售額', { font: { bold: true } });
  
  // 生成測試資料
  const products = ['筆記型電腦', '平板電腦', '智慧型手機', '耳機', '鍵盤', '滑鼠'];
  const regions = ['北區', '中區', '南區', '東區'];
  const months = ['1月', '2月', '3月', '4月', '5月', '6月'];
  
  let row = 2;
  for (let i = 0; i < 500; i++) {
    dataWs.setCell(`A${row}`, products[i % products.length]);
    dataWs.setCell(`B${row}`, regions[i % regions.length]);
    dataWs.setCell(`C${row}`, months[i % months.length]);
    dataWs.setCell(`D${row}`, Math.floor(Math.random() * 10000) + 1000);
    row++;
  }
  
  // 設定欄寬
  dataWs.setColumnWidth('A', 15);
  dataWs.setColumnWidth('B', 12);
  dataWs.setColumnWidth('C', 10);
  dataWs.setColumnWidth('D', 15);
  
  console.log(`✅ 已建立 ${row - 2} 筆測試資料`);
  
  // 建立 Pivot Table
  console.log('\n🎯 建立 Pivot Table...');
  
  // 定義 Pivot Table 欄位
  const fields = [
    {
      name: '產品',
      sourceColumn: '產品',
      type: 'row',
      showSubtotal: true,
      showGrandTotal: true
    },
    {
      name: '地區',
      sourceColumn: '地區',
      type: 'column',
      showSubtotal: true,
      showGrandTotal: true
    },
    {
      name: '月份',
      sourceColumn: '月份',
      type: 'filter',
      filterValues: ['1月', '2月', '3月']
    },
    {
      name: '銷售額',
      sourceColumn: '銷售額',
      type: 'value',
      function: 'sum',
      numberFormat: '#,##0',
      customName: '總銷售額'
    },
    {
      name: '銷售筆數',
      sourceColumn: '銷售額',
      type: 'value',
      function: 'count',
      customName: '銷售筆數'
    }
  ];
  
  // 建立 Pivot Table 配置
  const pivotConfig = {
    name: '銷售分析表',
    sourceRange: 'A1:D501',
    targetRange: 'F1:J30',
    fields: fields,
    showRowHeaders: true,
    showColumnHeaders: true,
    showRowSubtotals: true,
    showColumnSubtotals: true,
    showGrandTotals: true,
    autoFormat: true,
    compactRows: false,
    outlineData: true
  };
  
  // 建立 Pivot Table
  const pivotTable = wb.createPivotTable(pivotConfig);
  console.log('✅ Pivot Table 建立完成');
  
  // 測試 Pivot Table 功能
  console.log('\n🔧 測試 Pivot Table 功能...');
  
  // 取得欄位資訊
  const productField = pivotTable.getField('產品');
  console.log('產品欄位:', productField);
  
  // 應用篩選
  console.log('\n🔍 應用月份篩選...');
  pivotTable.applyFilter('月份', ['1月', '2月']);
  console.log('✅ 月份篩選已應用');
  
  // 取得 Pivot Table 資料
  console.log('\n📊 取得 Pivot Table 資料...');
  const pivotData = pivotTable.getData();
  console.log(`✅ 取得 ${pivotData.length} 行資料`);
  
  // 顯示前幾行資料
  console.log('\n📋 Pivot Table 資料預覽:');
  for (let i = 0; i < Math.min(5, pivotData.length); i++) {
    console.log(`  行 ${i + 1}:`, pivotData[i]);
  }
  
  // 測試欄位管理
  console.log('\n🔧 測試欄位管理...');
  
  // 添加新欄位
  const newField = {
    name: '平均銷售額',
    sourceColumn: '銷售額',
    type: 'value',
    function: 'average',
    numberFormat: '#,##0.00',
    customName: '平均銷售額'
  };
  
  pivotTable.addField(newField);
  console.log('✅ 新欄位已添加');
  
  // 重新整理 Pivot Table
  console.log('\n🔄 重新整理 Pivot Table...');
  pivotTable.refresh();
  console.log('✅ Pivot Table 已重新整理');
  
  // 取得更新後的資料
  const updatedData = pivotTable.getData();
  console.log(`✅ 更新後資料: ${updatedData.length} 行`);
  
  // 測試 Pivot Table 管理
  console.log('\n📋 Pivot Table 管理測試...');
  
  // 列出所有 Pivot Table
  const allPivotTables = wb.getAllPivotTables();
  console.log(`總共有 ${allPivotTables.length} 個 Pivot Table`);
  
  // 取得特定 Pivot Table
  const retrievedPivotTable = wb.getPivotTable('銷售分析表');
  if (retrievedPivotTable) {
    console.log('✅ 成功取得 Pivot Table:', retrievedPivotTable.name);
  }
  
  // 測試欄位重新排序
  console.log('\n🔄 測試欄位重新排序...');
  pivotTable.reorderFields(['產品', '地區', '銷售額', '銷售筆數', '平均銷售額']);
  console.log('✅ 欄位重新排序完成');
  
  // 清除篩選
  console.log('\n🧹 清除所有篩選...');
  pivotTable.clearFilters();
  console.log('✅ 篩選已清除');
  
  // 重新整理
  pivotTable.refresh();
  
  // 匯出到新工作表
  console.log('\n📤 匯出 Pivot Table 到新工作表...');
  const exportWs = pivotTable.exportToWorksheet('Pivot_Table_匯出');
  console.log('✅ Pivot Table 已匯出到工作表:', exportWs.name);
  
  // 設定匯出工作表的樣式
  exportWs.setColumnWidth('A', 20);
  exportWs.setColumnWidth('B', 15);
  exportWs.setColumnWidth('C', 15);
  exportWs.setColumnWidth('D', 15);
  exportWs.setColumnWidth('E', 15);
  
  // 生成 Excel 檔案
  console.log('\n💾 生成 Excel 檔案...');
  const buffer = await wb.writeBuffer();
  
  const filename = 'test-pivot-table.xlsx';
  fs.writeFileSync(filename, Buffer.from(buffer));
  console.log(`✅ Pivot Table 測試完成！檔案已儲存為: ${filename}`);
  
  // 顯示最終統計
  console.log('\n📊 最終統計:');
  console.log('工作表數量:', wb.getAllPivotTables().length);
  console.log('Pivot Table 數量:', wb.getAllPivotTables().length);
  
  // 顯示 Pivot Table 資訊
  for (const pt of wb.getAllPivotTables()) {
    console.log(`\n🎯 Pivot Table: ${pt.name}`);
    console.log(`  來源範圍: ${pt.config.sourceRange}`);
    console.log(`  目標範圍: ${pt.config.targetRange}`);
    console.log(`  欄位數量: ${pt.config.fields.length}`);
    console.log(`  資料行數: ${pt.getData().length}`);
  }
  
  console.log('\n🎯 Phase 5 Pivot Table 支援功能測試完成！');
}

// 執行測試
testPivotTable().catch(console.error);
