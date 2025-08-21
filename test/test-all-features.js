const { Workbook, ChartFactory } = require('../dist/index.js');
const fs = require('fs');

async function testAllFeatures() {
  console.log('🚀 測試所有功能 - 完整驗證');
  console.log('='.repeat(60));

  try {
    // 創建工作簿
    const workbook = new Workbook({
      memoryOptimization: true,
      chunkSize: 500,
      cacheEnabled: true
    });

    console.log('✅ 工作簿創建成功');

    // ===== Phase 1: 基本功能測試 =====
    console.log('\n📊 Phase 1: 基本功能測試');
    console.log('-'.repeat(40));

    const basicSheet = workbook.getWorksheet('基本功能');
    
    // 基本資料設定
    basicSheet.setCell('A1', '產品名稱', { font: { bold: true, size: 14 } });
    basicSheet.setCell('B1', '數量', { font: { bold: true, size: 14 } });
    basicSheet.setCell('C1', '單價', { font: { bold: true, size: 14 } });
    basicSheet.setCell('D1', '總價', { font: { bold: true, size: 14 } });

    basicSheet.setCell('A2', '筆記型電腦');
    basicSheet.setCell('B2', 5);
    basicSheet.setCell('C2', 80000);
    basicSheet.setFormula('D2', '=B2*C2');

    basicSheet.setCell('A3', '平板電腦');
    basicSheet.setCell('B3', 3);
    basicSheet.setCell('C3', 25000);
    basicSheet.setFormula('D3', '=B3*C3');

    console.log('✅ 基本資料設定完成');

    // ===== Phase 2: 樣式支援測試 =====
    console.log('\n🎨 Phase 2: 樣式支援測試');
    console.log('-'.repeat(40));

    const styleSheet = workbook.getWorksheet('樣式測試');
    
    // 字體樣式
    styleSheet.setCell('A1', '粗體文字', { font: { bold: true, size: 16, color: 'FF0000' } });
    styleSheet.setCell('A2', '斜體文字', { font: { italic: true, size: 14, color: '0000FF' } });
    styleSheet.setCell('A3', '底線文字', { font: { underline: true, size: 12 } });
    
    // 對齊樣式
    styleSheet.setCell('B1', '左對齊', { alignment: { horizontal: 'left' } });
    styleSheet.setCell('B2', '置中對齊', { alignment: { horizontal: 'center' } });
    styleSheet.setCell('B3', '右對齊', { alignment: { horizontal: 'right' } });
    
    // 填滿樣式
    styleSheet.setCell('C1', '淺灰背景', { fill: { type: 'pattern', patternType: 'solid', fgColor: 'E0E0E0' } });
    styleSheet.setCell('C2', '深灰背景', { fill: { type: 'pattern', patternType: 'solid', fgColor: '808080' } });
    
    // 邊框樣式
    styleSheet.setCell('D1', '細邊框', { border: { style: 'thin' } });
    styleSheet.setCell('D2', '粗邊框', { border: { style: 'thick' } });
    styleSheet.setCell('D3', '雙線邊框', { border: { style: 'double' } });

    console.log('✅ 樣式支援測試完成');

    // ===== Phase 3: 進階功能測試 =====
    console.log('\n🔧 Phase 3: 進階功能測試');
    console.log('-'.repeat(40));

    const advancedSheet = workbook.getWorksheet('進階功能');
    
    // 合併儲存格
    advancedSheet.mergeCells('A1:C1');
    advancedSheet.setCell('A1', '合併儲存格標題', { 
      font: { bold: true, size: 16 },
      alignment: { horizontal: 'center' }
    });
    
    // 欄寬和列高設定
    advancedSheet.setColumnWidth('A', 20);
    advancedSheet.setColumnWidth('B', 15);
    advancedSheet.setColumnWidth('C', 15);
    advancedSheet.setRowHeight(1, 30);
    
    // 凍結窗格
    advancedSheet.freezePanes(2, 1);
    
    // 公式支援
    advancedSheet.setCell('A2', '數值1');
    advancedSheet.setCell('B2', 100);
    advancedSheet.setFormula('C2', '=B2*2');
    
    advancedSheet.setCell('A3', '數值2');
    advancedSheet.setCell('B3', 200);
    advancedSheet.setFormula('C3', '=SUM(B2:B3)');

    console.log('✅ 進階功能測試完成');

    // ===== Phase 4: 效能優化測試 =====
    console.log('\n⚡ Phase 4: 效能優化測試');
    console.log('-'.repeat(40));

    const perfSheet = workbook.getWorksheet('效能測試');
    
    // 大型資料集測試
    const largeData = [];
    for (let i = 0; i < 1000; i++) {
      largeData.push([
        `產品${i + 1}`,
        Math.floor(Math.random() * 1000),
        Math.floor(Math.random() * 10000) + 1000,
        Math.floor(Math.random() * 100) + 1
      ]);
    }
    
    await workbook.addLargeDataset('效能測試', largeData, {
      startRow: 2,
      startCol: 1,
      chunkSize: 100
    });

    // 記憶體統計
    const memStats = workbook.getMemoryStats();
    console.log(`📊 記憶體使用統計:`);
    console.log(`  工作表數量: ${memStats.sheets}`);
    console.log(`  總儲存格數: ${memStats.totalCells.toLocaleString()}`);
    console.log(`  快取大小: ${memStats.cacheSize}`);
    console.log(`  快取命中率: ${(memStats.cacheHitRate * 100).toFixed(1)}%`);
    console.log(`  記憶體使用: ${(memStats.memoryUsage / 1024 / 1024).toFixed(2)} MB`);

    console.log('✅ 效能優化測試完成');

    // ===== Phase 5: Pivot Table 支援測試 =====
    console.log('\n🎯 Phase 5: Pivot Table 支援測試');
    console.log('-'.repeat(40));

    const pivotSheet = workbook.getWorksheet('Pivot資料');
    
    // 創建測試資料
    const products = ['筆記型電腦', '平板電腦', '智慧型手機', '耳機', '鍵盤', '滑鼠'];
    const regions = ['北區', '中區', '南區', '東區'];
    const months = ['1月', '2月', '3月', '4月', '5月', '6月'];
    
    // 添加標題行
    pivotSheet.setCell('A1', '產品', { font: { bold: true } });
    pivotSheet.setCell('B1', '地區', { font: { bold: true } });
    pivotSheet.setCell('C1', '月份', { font: { bold: true } });
    pivotSheet.setCell('D1', '銷售額', { font: { bold: true } });
    
    // 添加測試資料
    for (let i = 0; i < 200; i++) {
      const row = i + 2;
      const product = products[i % products.length];
      const region = regions[i % regions.length];
      const month = months[i % months.length];
      const sales = Math.floor(Math.random() * 10000) + 1000;
      
      pivotSheet.setCell(`A${row}`, product);
      pivotSheet.setCell(`B${row}`, region);
      pivotSheet.setCell(`C${row}`, month);
      pivotSheet.setCell(`D${row}`, sales);
    }

    // 創建 Pivot Table
    const pivotConfig = {
      name: '銷售分析表',
      sourceRange: 'A1:D201',
      targetRange: 'F1:J30',
      fields: [
        { name: '產品', sourceColumn: '產品', type: 'row', showSubtotal: true, showGrandTotal: true },
        { name: '地區', sourceColumn: '地區', type: 'column', showSubtotal: false, showGrandTotal: true },
        { name: '銷售額', sourceColumn: '銷售額', type: 'value', function: 'sum', customName: '銷售額總計' }
      ],
      showRowHeaders: true,
      showColumnHeaders: true,
      showRowSubtotals: true,
      showColumnSubtotals: false,
      showGrandTotals: true,
      autoFormat: true,
      compactRows: true,
      outlineData: true,
      mergeLabels: true
    };

    const pivotTable = workbook.createPivotTable(pivotConfig);
    console.log('✅ Pivot Table 創建成功');

    // 測試 Pivot Table 功能
    const pivotData = pivotTable.getData();
    console.log(`📊 Pivot Table 資料: ${pivotData.length} 行`);

    // 匯出 Pivot Table 到新工作表
    const exportSheet = pivotTable.exportToWorksheet('Pivot_Table_匯出');
    console.log('✅ Pivot Table 匯出成功');

    console.log('✅ Pivot Table 支援測試完成');

    // ===== Phase 6: 保護和圖表功能測試 =====
    console.log('\n🔒 Phase 6: 保護和圖表功能測試');
    console.log('-'.repeat(40));

    const protectedSheet = workbook.getWorksheet('保護和圖表');
    
    // 圖表支援 - 在保護之前添加資料
    const chartData = [
      ['月份', '銷售額', '成本', '利潤'],
      ['1月', 50000, 35000, 15000],
      ['2月', 60000, 40000, 20000],
      ['3月', 45000, 32000, 13000],
      ['4月', 70000, 48000, 22000],
      ['5月', 55000, 38000, 17000],
      ['6月', 80000, 55000, 25000]
    ];

    // 添加圖表資料
    for (let i = 0; i < chartData.length; i++) {
      for (let j = 0; j < chartData[i].length; j++) {
        const address = `${String.fromCharCode(65 + j)}${i + 1}`;
        const value = chartData[i][j];
        if (typeof value === 'number') {
          protectedSheet.setCell(address, value);
        } else {
          protectedSheet.setCell(address, value, { font: { bold: true } });
        }
      }
    }

    // 創建柱狀圖
    const columnChart = ChartFactory.createColumnChart('銷售分析圖', [], {
      title: '月度銷售分析',
      width: 600,
      height: 400,
      xAxisTitle: '月份',
      yAxisTitle: '金額',
      showLegend: true,
      showDataLabels: true
    });

    columnChart.addSeries({ series: '銷售額', xRange: 'A2:A7', yRange: 'B2:B7' });
    columnChart.addSeries({ series: '成本', xRange: 'A2:A7', yRange: 'C2:C7' });
    columnChart.addSeries({ series: '利潤', xRange: 'A2:A7', yRange: 'D2:D7' });

    protectedSheet.addChart(columnChart);
    console.log('✅ 圖表創建成功');

    // 創建圓餅圖
    const pieChart = ChartFactory.createPieChart('利潤分布圖', [], {
      title: '利潤分布',
      width: 400,
      height: 300,
      showLegend: true,
      showDataLabels: true
    });

    pieChart.addSeries({ series: '利潤', xRange: 'A2:A7', yRange: 'D2:D7' });
    pieChart.moveTo(650, 50);

    protectedSheet.addChart(pieChart);
    console.log('✅ 圓餅圖創建成功');

    // 工作表保護 - 在添加圖表後設定
    protectedSheet.protect('password123', {
      selectLockedCells: false,
      selectUnlockedCells: true,
      formatCells: false,
      formatColumns: false,
      formatRows: false,
      insertColumns: false,
      insertRows: false,
      insertHyperlinks: false,
      deleteColumns: false,
      deleteRows: false,
      sort: false,
      autoFilter: false,
      pivotTables: false
    });
    console.log('✅ 工作表保護設定完成');

    // 工作簿保護
    workbook.protect('workbook123', {
      structure: true,
      windows: true
    });
    console.log('✅ 工作簿保護設定完成');

    console.log('✅ 保護和圖表功能測試完成');

    // ===== 生成 Excel 檔案 =====
    console.log('\n💾 生成 Excel 檔案');
    console.log('-'.repeat(40));

    try {
      // 嘗試使用動態 Pivot Table 方法
      console.log('🎯 嘗試生成包含動態 Pivot Table 的檔案...');
      const buffer = await workbook.writeBufferWithPivotTables();
      fs.writeFileSync('test-all-features.xlsx', new Uint8Array(buffer));
      console.log('✅ 動態 Pivot Table Excel 檔案已生成: test-all-features.xlsx');
    } catch (error) {
      console.log('⚠️ 動態 Pivot Table 生成失敗，使用標準方法:', error.message);
      const buffer = await workbook.writeBuffer();
      fs.writeFileSync('test-all-features.xlsx', new Uint8Array(buffer));
      console.log('✅ 標準 Excel 檔案已生成: test-all-features.xlsx');
    }

    // ===== 最終統計 =====
    console.log('\n📊 最終統計');
    console.log('-'.repeat(40));
    console.log(`工作表數量: ${workbook.getWorksheets().length}`);
    console.log(`Pivot Table 數量: ${workbook.getAllPivotTables().length}`);
    console.log(`圖表數量: ${protectedSheet.getCharts().length}`);
    console.log(`工作簿保護: ${workbook.isProtected() ? '是' : '否'}`);
    console.log(`工作表保護: ${protectedSheet.isProtected() ? '是' : '否'}`);

    // 顯示工作表名稱
    const sheetNames = workbook.getWorksheets().map(ws => ws.name);
    console.log(`工作表名稱: ${sheetNames.join(', ')}`);

    console.log('\n🎉 所有功能測試完成！');
    console.log('📝 請檢查生成的 test-all-features.xlsx 檔案');

  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error(error.stack);
  }
}

// 執行測試
testAllFeatures();
