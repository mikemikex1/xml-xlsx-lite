const { Workbook, ChartFactory } = require('../dist/index.js');
const fs = require('fs');

async function testPhase6() {
  console.log('🔒 測試 Phase 6: 工作表保護和圖表支援');
  
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
  for (let i = 0; i < 100; i++) {
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
  
  // 測試工作表保護
  console.log('\n🔒 測試工作表保護...');
  
  try {
    // 保護工作表
    dataWs.protect('password123', {
      selectLockedCells: false,
      selectUnlockedCells: true,
      formatCells: false,
      insertRows: false,
      deleteRows: false
    });
    console.log('✅ 工作表保護已啟用');
    
    // 檢查保護狀態
    console.log('工作表保護狀態:', dataWs.isProtected());
    console.log('保護選項:', dataWs.getProtectionOptions());
    
    // 嘗試修改受保護的儲存格（應該失敗）
    try {
      dataWs.setCell('A1', '測試修改');
      console.log('❌ 保護失敗：應該無法修改儲存格');
    } catch (error) {
      console.log('✅ 保護成功：無法修改受保護的儲存格');
    }
    
    // 解除保護
    dataWs.unprotect('password123');
    console.log('✅ 工作表保護已解除');
    
    // 再次嘗試修改（應該成功）
    dataWs.setCell('A1', '保護解除後可修改');
    console.log('✅ 保護解除後可以修改儲存格');
    
  } catch (error) {
    console.log('❌ 工作表保護測試失敗:', error.message);
  }
  
  // 測試圖表支援
  console.log('\n📊 測試圖表支援...');
  
  try {
    // 建立柱狀圖
    const columnChart = ChartFactory.createColumnChart(
      '銷售額柱狀圖',
      [
        {
          series: '銷售額',
          categories: 'A2:A7',
          values: 'D2:D7',
          color: '#FF6B6B'
        }
      ],
      {
        title: '產品銷售額分析',
        xAxisTitle: '產品',
        yAxisTitle: '銷售額',
        width: 500,
        height: 300,
        showLegend: true,
        showDataLabels: true
      },
      { row: 1, col: 6 }
    );
    
    // 添加圖表到工作表
    dataWs.addChart(columnChart);
    console.log('✅ 柱狀圖已添加');
    
    // 建立圓餅圖
    const pieChart = ChartFactory.createPieChart(
      '地區銷售圓餅圖',
      [
        {
          series: '地區銷售',
          categories: 'B2:B5',
          values: 'D2:D5',
          color: '#4ECDC4'
        }
      ],
      {
        title: '各地區銷售佔比',
        width: 400,
        height: 300,
        showLegend: true,
        showDataLabels: true
      },
      { row: 15, col: 6 }
    );
    
    dataWs.addChart(pieChart);
    console.log('✅ 圓餅圖已添加');
    
    // 建立折線圖
    const lineChart = ChartFactory.createLineChart(
      '月份趨勢折線圖',
      [
        {
          series: '銷售趨勢',
          categories: 'C2:C7',
          values: 'D2:D7',
          color: '#45B7D1'
        }
      ],
      {
        title: '銷售額月份趨勢',
        xAxisTitle: '月份',
        yAxisTitle: '銷售額',
        width: 600,
        height: 300,
        showLegend: true,
        showGridlines: true
      },
      { row: 1, col: 12 }
    );
    
    dataWs.addChart(lineChart);
    console.log('✅ 折線圖已添加');
    
    // 檢查圖表
    const charts = dataWs.getCharts();
    console.log(`總共有 ${charts.length} 個圖表`);
    
    for (const chart of charts) {
      console.log(`圖表: ${chart.name}, 類型: ${chart.type}`);
    }
    
    // 測試圖表管理
    const retrievedChart = dataWs.getChart('銷售額柱狀圖');
    if (retrievedChart) {
      console.log('✅ 成功取得圖表:', retrievedChart.name);
      
      // 更新圖表選項
      retrievedChart.updateOptions({
        title: '更新後的銷售額分析',
        width: 550,
        height: 350
      });
      console.log('✅ 圖表選項已更新');
      
      // 移動圖表位置
      retrievedChart.moveTo(20, 6);
      console.log('✅ 圖表位置已移動');
    }
    
  } catch (error) {
    console.log('❌ 圖表支援測試失敗:', error.message);
  }
  
  // 測試工作簿保護
  console.log('\n🔒 測試工作簿保護...');
  
  try {
    // 保護工作簿
    wb.protect('workbook123', {
      structure: true,
      windows: false
    });
    console.log('✅ 工作簿保護已啟用');
    
    // 檢查保護狀態
    console.log('工作簿保護狀態:', wb.isProtected());
    console.log('保護選項:', wb.getProtectionOptions());
    
    // 解除保護
    wb.unprotect('workbook123');
    console.log('✅ 工作簿保護已解除');
    
  } catch (error) {
    console.log('❌ 工作簿保護測試失敗:', error.message);
  }
  
  // 生成 Excel 檔案
  console.log('\n💾 生成 Excel 檔案...');
  const buffer = await wb.writeBuffer();
  
  const filename = 'test-phase6.xlsx';
  fs.writeFileSync(filename, Buffer.from(buffer));
  console.log(`✅ Phase 6 測試完成！檔案已儲存為: ${filename}`);
  
  // 顯示最終統計
  console.log('\n📊 最終統計:');
  console.log('工作表數量:', 1);
  console.log('圖表數量:', dataWs.getCharts().length);
  console.log('工作表保護狀態:', dataWs.isProtected());
  console.log('工作簿保護狀態:', wb.isProtected());
  
  console.log('\n🎯 Phase 6 工作表保護和圖表支援功能測試完成！');
}

// 執行測試
testPhase6().catch(console.error);
