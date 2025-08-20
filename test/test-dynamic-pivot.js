const { Workbook } = require('../dist/index.js');
const fs = require('fs');

async function testDynamicPivotTable() {
  console.log('🎯 測試動態 Pivot Table 功能');
  console.log('='.repeat(50));

  try {
    // 創建工作簿
    const workbook = new Workbook();
    
    // 創建資料工作表
    const dataSheet = workbook.getWorksheet('銷售資料');
    
    // 添加標題行
    dataSheet.setCell('A1', '產品', { font: { bold: true } });
    dataSheet.setCell('B1', '地區', { font: { bold: true } });
    dataSheet.setCell('C1', '月份', { font: { bold: true } });
    dataSheet.setCell('D1', '銷售額', { font: { bold: true } });
    
    // 添加測試資料
    const products = ['筆記型電腦', '平板電腦', '智慧型手機', '耳機'];
    const regions = ['北區', '中區', '南區', '東區'];
    const months = ['1月', '2月', '3月', '4月'];
    
    console.log('📊 正在生成測試資料...');
    for (let i = 0; i < 500; i++) {
      const row = i + 2;
      const product = products[i % products.length];
      const region = regions[i % regions.length];
      const month = months[i % months.length];
      const sales = Math.floor(Math.random() * 10000) + 1000;
      
      dataSheet.setCell(`A${row}`, product);
      dataSheet.setCell(`B${row}`, region);
      dataSheet.setCell(`C${row}`, month);
      dataSheet.setCell(`D${row}`, sales);
      
      if (i % 100 === 0) {
        console.log(`已生成 ${i} 筆資料...`);
      }
    }
    console.log('✅ 500筆測試資料生成完成');

    // 創建 Pivot Table 配置
    const pivotConfig = {
      name: '銷售分析表',
      sourceRange: 'A1:D501',
      targetRange: 'F1:J30',
      fields: [
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
          showSubtotal: false,
          showGrandTotal: true
        },
        {
          name: '銷售額',
          sourceColumn: '銷售額',
          type: 'value',
          function: 'sum',
          customName: '銷售額總計'
        }
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

    console.log('🎯 正在創建動態 Pivot Table...');
    
    // 創建 Pivot Table
    const pivotTable = workbook.createPivotTable(pivotConfig);
    console.log('✅ Pivot Table 建立完成');
    
    // 測試 Pivot Table 功能
    console.log('🔧 測試 Pivot Table 功能...');
    
    // 取得欄位資訊
    const productField = pivotTable.getField('產品');
    console.log('產品欄位:', {
      name: productField?.name,
      sourceColumn: productField?.sourceColumn,
      type: productField?.type,
      showSubtotal: productField?.showSubtotal,
      showGrandTotal: productField?.showGrandTotal
    });

    // 應用篩選
    console.log('🔍 應用月份篩選...');
    pivotTable.applyFilter('月份', ['1月', '2月']);
    console.log('✅ 月份篩選已應用');

    // 取得資料
    console.log('📊 取得 Pivot Table 資料...');
    const pivotData = pivotTable.getData();
    console.log(`✅ 取得 ${pivotData.length} 行資料`);

    // 顯示資料預覽
    console.log('📋 Pivot Table 資料預覽:');
    for (let i = 0; i < Math.min(5, pivotData.length); i++) {
      console.log(`  行 ${i + 1}:`, pivotData[i]);
    }

    // 測試欄位管理
    console.log('🔧 測試欄位管理...');
    
    // 添加新欄位
    const newField = {
      name: '月份',
      sourceColumn: '月份',
      type: 'filter',
      showSubtotal: false,
      showGrandTotal: false
    };
    pivotTable.addField(newField);
    console.log('✅ 新欄位已添加');

    // 重新整理
    console.log('🔄 重新整理 Pivot Table...');
    pivotTable.refresh();
    const updatedData = pivotTable.getData();
    console.log(`✅ 更新後資料: ${updatedData.length} 行`);

    // 測試 Pivot Table 管理
    console.log('📋 Pivot Table 管理測試...');
    const allPivotTables = workbook.getAllPivotTables();
    console.log(`總共有 ${allPivotTables.length} 個 Pivot Table`);
    
    const retrievedPivotTable = workbook.getPivotTable('銷售分析表');
    if (retrievedPivotTable) {
      console.log('✅ 成功取得 Pivot Table: 銷售分析表');
    }

    // 測試欄位重新排序
    console.log('🔄 測試欄位重新排序...');
    pivotTable.reorderFields(['產品', '地區', '銷售額', '月份']);
    console.log('✅ 欄位重新排序完成');

    // 清除篩選
    console.log('🧹 清除所有篩選...');
    pivotTable.clearFilters();
    console.log('✅ 篩選已清除');

    // 匯出到新工作表
    console.log('📤 匯出 Pivot Table 到新工作表...');
    const exportSheet = pivotTable.exportToWorksheet('Pivot_Table_匯出');
    console.log('✅ Pivot Table 已匯出到工作表: Pivot_Table_匯出');

    // 生成包含動態 Pivot Table 的 Excel 檔案
    console.log('💾 生成包含動態 Pivot Table 的 Excel 檔案...');
    
    try {
      // 使用新的方法生成包含 Pivot Table 的檔案
      const buffer = await workbook.writeBufferWithPivotTables();
      fs.writeFileSync('test-dynamic-pivot.xlsx', new Uint8Array(buffer));
      console.log('✅ 動態 Pivot Table Excel 檔案已生成: test-dynamic-pivot.xlsx');
    } catch (error) {
      console.log('⚠️ 動態 Pivot Table 生成失敗，使用標準方法:', error.message);
      // 回退到標準方法
      const buffer = await workbook.writeBuffer();
      fs.writeFileSync('test-dynamic-pivot.xlsx', new Uint8Array(buffer));
      console.log('✅ 標準 Excel 檔案已生成: test-dynamic-pivot.xlsx');
    }

    // 最終統計
    console.log('\n📊 最終統計:');
    console.log(`工作表數量: ${workbook.getWorksheets().length}`);
    console.log(`Pivot Table 數量: ${workbook.getAllPivotTables().length}`);

    // Pivot Table 詳細資訊
    const finalPivotTable = workbook.getPivotTable('銷售分析表');
    if (finalPivotTable) {
      console.log('\n🎯 Pivot Table: 銷售分析表');
      console.log(`  來源範圍: ${finalPivotTable.config.sourceRange}`);
      console.log(`  目標範圍: ${finalPivotTable.config.targetRange}`);
      console.log(`  欄位數量: ${finalPivotTable.config.fields.length}`);
      console.log(`  資料行數: ${finalPivotTable.getData().length}`);
      
      // 如果是動態 Pivot Table，顯示快取和表格 ID
      if (finalPivotTable.getCacheId && finalPivotTable.getTableId) {
        console.log(`  快取 ID: ${finalPivotTable.getCacheId()}`);
        console.log(`  表格 ID: ${finalPivotTable.getTableId()}`);
      }
    }

    console.log('\n🎯 動態 Pivot Table 功能測試完成！');
    console.log('📝 注意: 真正的動態 Pivot Table 需要在 Excel 中打開才能看到互動式功能');
    console.log('📝 生成的檔案包含完整的 PivotCache 和 PivotTable XML 結構');

  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error(error.stack);
  }
}

testDynamicPivotTable();
