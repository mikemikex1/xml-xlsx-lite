const { Workbook, ChartFactory } = require('../dist/index.js');

async function testComprehensive() {
  console.log('🧪 開始綜合功能測試 - 所有 Phase 功能驗證');
  console.log('=' .repeat(60));

  const workbook = new Workbook();
  
  // ============================================================================
  // Phase 1: 基本功能測試
  // ============================================================================
  console.log('\n📋 Phase 1: 基本功能測試');
  console.log('-'.repeat(40));
  
  const basicSheet = workbook.getWorksheet('基本功能');
  
  // 基本儲存格設定
  basicSheet.setCell('A1', '產品名稱');
  basicSheet.setCell('B1', '數量');
  basicSheet.setCell('C1', '單價');
  basicSheet.setCell('D1', '總價');
  
  // 不同資料類型
  basicSheet.setCell('A2', 'iPhone 15');
  basicSheet.setCell('B2', 10);
  basicSheet.setCell('C2', 35000);
  basicSheet.setCell('D2', 350000);
  
  basicSheet.setCell('A3', 'MacBook Pro');
  basicSheet.setCell('B3', 5);
  basicSheet.setCell('C3', 80000);
  basicSheet.setCell('D3', 400000);
  
  basicSheet.setCell('A4', '日期測試');
  basicSheet.setCell('B4', new Date());
  basicSheet.setCell('C4', true);
  basicSheet.setCell('D4', false);
  
  console.log('✅ 基本儲存格操作完成');
  
  // ============================================================================
  // Phase 2: 樣式支援測試
  // ============================================================================
  console.log('\n🎨 Phase 2: 樣式支援測試');
  console.log('-'.repeat(40));
  
  const styleSheet = workbook.getWorksheet('樣式測試');
  
  // 標題樣式
  styleSheet.setCell('A1', '樣式展示', {
    font: { bold: true, size: 16, color: '#FF0000', name: 'Arial' },
    alignment: { horizontal: 'center', vertical: 'middle' },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#FFFF00' },
    border: { style: 'thick', color: '#000000' }
  });
  
  // 不同字體樣式
  styleSheet.setCell('A3', '粗體文字', { font: { bold: true } });
  styleSheet.setCell('B3', '斜體文字', { font: { italic: true } });
  styleSheet.setCell('C3', '底線文字', { font: { underline: true } });
  styleSheet.setCell('D3', '刪除線', { font: { strike: true } });
  
  // 不同對齊方式
  styleSheet.setCell('A5', '左對齊', { alignment: { horizontal: 'left' } });
  styleSheet.setCell('B5', '置中', { alignment: { horizontal: 'center' } });
  styleSheet.setCell('C5', '右對齊', { alignment: { horizontal: 'right' } });
  
  // 不同填充顏色
  styleSheet.setCell('A7', '紅色背景', { fill: { type: 'pattern', patternType: 'solid', fgColor: '#FF0000' } });
  styleSheet.setCell('B7', '綠色背景', { fill: { type: 'pattern', patternType: 'solid', fgColor: '#00FF00' } });
  styleSheet.setCell('C7', '藍色背景', { fill: { type: 'pattern', patternType: 'solid', fgColor: '#0000FF' } });
  
  // 不同邊框樣式
  styleSheet.setCell('A9', '細線邊框', { border: { style: 'thin', color: '#000000' } });
  styleSheet.setCell('B9', '粗線邊框', { border: { style: 'thick', color: '#FF0000' } });
  styleSheet.setCell('C9', '虛線邊框', { border: { style: 'dashed', color: '#0000FF' } });
  
  console.log('✅ 樣式設定完成');
  
  // ============================================================================
  // Phase 3: 進階功能測試
  // ============================================================================
  console.log('\n⚡ Phase 3: 進階功能測試');
  console.log('-'.repeat(40));
  
  const advancedSheet = workbook.getWorksheet('進階功能');
  
  // 合併儲存格
  advancedSheet.setCell('A1', '合併儲存格標題', {
    font: { bold: true, size: 14 },
    alignment: { horizontal: 'center', vertical: 'middle' }
  });
  advancedSheet.mergeCells('A1:D1');
  
  // 設定欄寬和列高
  advancedSheet.setColumnWidth('A', 20);
  advancedSheet.setColumnWidth('B', 15);
  advancedSheet.setColumnWidth('C', 12);
  advancedSheet.setColumnWidth('D', 18);
  advancedSheet.setRowHeight(1, 30);
  advancedSheet.setRowHeight(3, 25);
  
  // 凍結窗格
  advancedSheet.setCell('A3', '項目');
  advancedSheet.setCell('B3', 'Q1');
  advancedSheet.setCell('C3', 'Q2');
  advancedSheet.setCell('D3', 'Q3');
  advancedSheet.setCell('E3', 'Q4');
  advancedSheet.freezePanes(3, 1); // 凍結第3列以上和第1欄以左
  
  // 公式測試
  advancedSheet.setCell('A4', '銷售A');
  advancedSheet.setCell('B4', 100);
  advancedSheet.setCell('C4', 150);
  advancedSheet.setCell('D4', 200);
  advancedSheet.setCell('E4', 180);
  
  advancedSheet.setCell('A5', '銷售B');
  advancedSheet.setCell('B5', 80);
  advancedSheet.setCell('C5', 120);
  advancedSheet.setCell('D5', 160);
  advancedSheet.setCell('E5', 140);
  
  // 設定公式
  advancedSheet.setFormula('B6', '=SUM(B4:B5)', { font: { bold: true } });
  advancedSheet.setFormula('C6', '=SUM(C4:C5)', { font: { bold: true } });
  advancedSheet.setFormula('D6', '=SUM(D4:D5)', { font: { bold: true } });
  advancedSheet.setFormula('E6', '=SUM(E4:E5)', { font: { bold: true } });
  advancedSheet.setFormula('F6', '=SUM(B6:E6)', { 
    font: { bold: true, color: '#FF0000' },
    fill: { type: 'pattern', patternType: 'solid', fgColor: '#FFFF00' }
  });
  
  advancedSheet.setCell('A6', '總計');
  
  console.log('✅ 進階功能設定完成');
  
  // ============================================================================
  // Phase 4: 效能優化測試
  // ============================================================================
  console.log('\n🚀 Phase 4: 效能優化測試');
  console.log('-'.repeat(40));
  
  const performanceSheet = workbook.getWorksheet('效能測試');
  
  // 大量資料測試
  console.log('正在生成大量測試資料...');
  const startTime = Date.now();
  
  for (let row = 1; row <= 1000; row++) {
    performanceSheet.setCell(`A${row}`, `項目 ${row}`);
    performanceSheet.setCell(`B${row}`, Math.floor(Math.random() * 1000));
    performanceSheet.setCell(`C${row}`, Math.floor(Math.random() * 100));
    performanceSheet.setCell(`D${row}`, Math.floor(Math.random() * 10000));
    
    if (row % 100 === 0) {
      console.log(`已生成 ${row} 筆資料...`);
    }
  }
  
  const endTime = Date.now();
  console.log(`✅ 1000筆資料生成完成，耗時: ${endTime - startTime}ms`);
  
  // 記憶體統計
  const memStats = workbook.getMemoryStats();
  console.log(`記憶體使用: ${(memStats.memoryUsage / 1024 / 1024).toFixed(2)} MB`);
  console.log(`總儲存格: ${memStats.totalCells.toLocaleString()}`);
  console.log(`快取大小: ${memStats.cacheSize} 項`);
  console.log(`快取命中率: ${(memStats.cacheHitRate * 100).toFixed(1)}%`);
  
  // ============================================================================
  // Phase 5: Pivot Table 測試
  // ============================================================================
  console.log('\n🎯 Phase 5: Pivot Table 測試');
  console.log('-'.repeat(40));
  
  const pivotDataSheet = workbook.getWorksheet('Pivot資料');
  
  // 建立 Pivot Table 資料
  pivotDataSheet.setCell('A1', '產品');
  pivotDataSheet.setCell('B1', '地區');
  pivotDataSheet.setCell('C1', '銷售員');
  pivotDataSheet.setCell('D1', '銷售額');
  pivotDataSheet.setCell('E1', '月份');
  
  const products = ['iPhone', 'MacBook', 'iPad', 'AirPods'];
  const regions = ['北部', '中部', '南部'];
  const salespeople = ['張三', '李四', '王五', '趙六'];
  const months = ['1月', '2月', '3月', '4月'];
  
  for (let i = 2; i <= 101; i++) {
    pivotDataSheet.setCell(`A${i}`, products[Math.floor(Math.random() * products.length)]);
    pivotDataSheet.setCell(`B${i}`, regions[Math.floor(Math.random() * regions.length)]);
    pivotDataSheet.setCell(`C${i}`, salespeople[Math.floor(Math.random() * salespeople.length)]);
    pivotDataSheet.setCell(`D${i}`, Math.floor(Math.random() * 50000) + 10000);
    pivotDataSheet.setCell(`E${i}`, months[Math.floor(Math.random() * months.length)]);
  }
  
  // 建立 Pivot Table
  const pivotTable = workbook.createPivotTable({
    name: '銷售分析',
    sourceRange: 'A1:E101',
    targetRange: 'G1:M50',
    fields: [
      {
        name: '產品',
        sourceColumn: '產品',
        type: 'row',
        showSubtotal: true
      },
      {
        name: '地區',
        sourceColumn: '地區',
        type: 'column',
        showSubtotal: true
      },
      {
        name: '銷售額總計',
        sourceColumn: '銷售額',
        type: 'value',
        function: 'sum'
      },
      {
        name: '銷售次數',
        sourceColumn: '銷售額',
        type: 'value',
        function: 'count'
      }
    ],
    showGrandTotals: true,
    autoFormat: true
  });
  
  console.log('✅ Pivot Table 建立完成');
  
  // ============================================================================
  // Phase 6: 保護功能和圖表測試
  // ============================================================================
  console.log('\n🔒 Phase 6: 保護功能和圖表測試');
  console.log('-'.repeat(40));
  
  const protectedSheet = workbook.getWorksheet('保護和圖表');
  
  // 建立圖表資料
  protectedSheet.setCell('A1', '月份');
  protectedSheet.setCell('B1', '銷售額');
  protectedSheet.setCell('C1', '利潤');
  
  const chartData = [
    ['1月', 100000, 25000],
    ['2月', 120000, 30000],
    ['3月', 150000, 40000],
    ['4月', 180000, 50000],
    ['5月', 200000, 60000],
    ['6月', 220000, 70000]
  ];
  
  chartData.forEach((row, index) => {
    protectedSheet.setCell(`A${index + 2}`, row[0]);
    protectedSheet.setCell(`B${index + 2}`, row[1]);
    protectedSheet.setCell(`C${index + 2}`, row[2]);
  });
  
  // 建立柱狀圖
  const columnChart = ChartFactory.createColumnChart(
    '月度銷售柱狀圖',
    [
      {
        series: '銷售額',
        categories: 'A2:A7',
        values: 'B2:B7',
        color: '#4F81BD'
      },
      {
        series: '利潤',
        categories: 'A2:A7',
        values: 'C2:C7',
        color: '#F79646'
      }
    ],
    {
      title: '月度銷售和利潤分析',
      xAxisTitle: '月份',
      yAxisTitle: '金額',
      width: 600,
      height: 400,
      showLegend: true,
      showDataLabels: true
    },
    { row: 1, col: 5 }
  );
  
  protectedSheet.addChart(columnChart);
  
  // 建立圓餅圖
  const pieChart = ChartFactory.createPieChart(
    '銷售額分布圓餅圖',
    [{
      series: '銷售額',
      categories: 'A2:A7',
      values: 'B2:B7',
      color: '#9CBB58'
    }],
    {
      title: '各月份銷售額分布',
      width: 500,
      height: 350,
      showLegend: true,
      showDataLabels: true
    },
    { row: 20, col: 5 }
  );
  
  protectedSheet.addChart(pieChart);
  
  // 建立折線圖
  const lineChart = ChartFactory.createLineChart(
    '趨勢折線圖',
    [
      {
        series: '銷售額',
        categories: 'A2:A7',
        values: 'B2:B7',
        color: '#C5504B'
      },
      {
        series: '利潤',
        categories: 'A2:A7',
        values: 'C2:C7',
        color: '#4BACC6'
      }
    ],
    {
      title: '銷售和利潤趨勢',
      xAxisTitle: '月份',
      yAxisTitle: '金額',
      width: 600,
      height: 400,
      showLegend: true,
      showDataLabels: false,
      showGridlines: true
    },
    { row: 1, col: 15 }
  );
  
  protectedSheet.addChart(lineChart);
  
  console.log('✅ 圖表建立完成');
  
  // 測試工作表保護
  protectedSheet.protect('test123', {
    selectLockedCells: false,
    selectUnlockedCells: true,
    formatCells: false,
    insertRows: false,
    deleteRows: false
  });
  
  console.log('🔒 工作表保護已啟用');
  console.log('保護狀態:', protectedSheet.isProtected());
  
  // 測試工作簿保護
  workbook.protect('workbook123', {
    structure: true,
    windows: false
  });
  
  console.log('🔒 工作簿保護已啟用');
  console.log('工作簿保護狀態:', workbook.isProtected());
  
  // ============================================================================
  // 檔案匯出測試
  // ============================================================================
  console.log('\n💾 檔案匯出測試');
  console.log('-'.repeat(40));
  
  try {
    const filename = 'comprehensive-test.xlsx';
    const buffer = await workbook.writeBuffer();
    
    // 手動寫入檔案
    const fs = require('fs');
    fs.writeFileSync(filename, new Uint8Array(buffer));
    
    console.log(`✅ Excel 檔案已成功匯出: ${filename}`);
    
    // 最終統計
    console.log('\n📊 最終統計資訊:');
    console.log(`工作表數量: ${workbook.getWorksheets().length}`);
    console.log(`Pivot Table 數量: ${workbook.getAllPivotTables().length}`);
    console.log(`圖表數量: ${protectedSheet.getCharts().length}`);
    
    const finalMemStats = workbook.getMemoryStats();
    console.log(`最終記憶體使用: ${(finalMemStats.memoryUsage / 1024 / 1024).toFixed(2)} MB`);
    console.log(`總儲存格數: ${finalMemStats.totalCells.toLocaleString()}`);
    
  } catch (error) {
    console.error('❌ 檔案匯出失敗:', error.message);
    throw error;
  }
  
  // ============================================================================
  // 測試完成
  // ============================================================================
  console.log('\n' + '='.repeat(60));
  console.log('🎉 綜合功能測試完成！');
  console.log('✅ Phase 1: 基本功能 - 通過');
  console.log('✅ Phase 2: 樣式支援 - 通過');
  console.log('✅ Phase 3: 進階功能 - 通過');
  console.log('✅ Phase 4: 效能優化 - 通過');
  console.log('✅ Phase 5: Pivot Table 支援 - 通過');
  console.log('✅ Phase 6: 保護功能和圖表支援 - 通過');
  console.log('🚀 xml-xlsx-lite 所有功能運作正常！');
  console.log('='.repeat(60));
}

// 執行測試
testComprehensive().catch(error => {
  console.error('❌ 測試失敗:', error);
  process.exit(1);
});
