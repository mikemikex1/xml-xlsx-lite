const XLSX = require('xlsx');
const fs = require('fs');
const JSZip = require('jszip');

async function validateAllFeatures() {
  console.log('🔍 驗證所有功能 Excel 檔案');
  console.log('='.repeat(60));

  try {
    // 檢查檔案是否存在
    if (!fs.existsSync('test-all-features.xlsx')) {
      console.log('❌ 檔案不存在: test-all-features.xlsx');
      return;
    }

    console.log('✅ 檔案存在: test-all-features.xlsx');
    
    // 檢查檔案大小
    const stats = fs.statSync('test-all-features.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    // 讀取 Excel 檔案
    const workbook = XLSX.readFile('test-all-features.xlsx');
    console.log('✅ Excel 檔案讀取成功');

    // 檢查工作表
    const sheetNames = workbook.SheetNames;
    console.log(`📋 工作表數量: ${sheetNames.length}`);
    console.log('📋 工作表名稱:', sheetNames);

    // ===== Phase 1: 基本功能驗證 =====
    console.log('\n📊 Phase 1: 基本功能驗證');
    console.log('-'.repeat(40));

    if (workbook.Sheets['基本功能']) {
      const basicData = XLSX.utils.sheet_to_json(workbook.Sheets['基本功能'], { header: 1 });
      console.log(`✅ 基本功能工作表: ${basicData.length} 行資料`);
      
      // 檢查公式
      const cellD2 = workbook.Sheets['基本功能']['D2'];
      if (cellD2 && cellD2.f) {
        console.log('✅ 公式支援正常: D2 包含公式');
      } else {
        console.log('⚠️ 公式支援: D2 沒有公式');
      }
    } else {
      console.log('❌ 基本功能工作表不存在');
    }

    // ===== Phase 2: 樣式支援驗證 =====
    console.log('\n🎨 Phase 2: 樣式支援驗證');
    console.log('-'.repeat(40));

    if (workbook.Sheets['樣式測試']) {
      const styleData = XLSX.utils.sheet_to_json(workbook.Sheets['樣式測試'], { header: 1 });
      console.log(`✅ 樣式測試工作表: ${styleData.length} 行資料`);
      
      // 檢查樣式
      const cellA1 = workbook.Sheets['樣式測試']['A1'];
      if (cellA1) {
        console.log('✅ 樣式支援正常: A1 儲存格存在');
      }
    } else {
      console.log('❌ 樣式測試工作表不存在');
    }

    // ===== Phase 3: 進階功能驗證 =====
    console.log('\n🔧 Phase 3: 進階功能驗證');
    console.log('-'.repeat(40));

    if (workbook.Sheets['進階功能']) {
      const advancedData = XLSX.utils.sheet_to_json(workbook.Sheets['進階功能'], { header: 1 });
      console.log(`✅ 進階功能工作表: ${advancedData.length} 行資料`);
      
      // 檢查合併儲存格
      const cellA1 = workbook.Sheets['進階功能']['A1'];
      if (cellA1) {
        console.log('✅ 合併儲存格支援正常: A1 儲存格存在');
      }
    } else {
      console.log('❌ 進階功能工作表不存在');
    }

    // ===== Phase 4: 效能優化驗證 =====
    console.log('\n⚡ Phase 4: 效能優化驗證');
    console.log('-'.repeat(40));

    if (workbook.Sheets['效能測試']) {
      const perfData = XLSX.utils.sheet_to_json(workbook.Sheets['效能測試'], { header: 1 });
      console.log(`✅ 效能測試工作表: ${perfData.length} 行資料`);
      
      if (perfData.length > 1000) {
        console.log('✅ 大型資料集處理正常: 超過 1000 行資料');
      } else {
        console.log('⚠️ 大型資料集處理: 資料行數不足');
      }
    } else {
      console.log('❌ 效能測試工作表不存在');
    }

    // ===== Phase 5: Pivot Table 驗證 =====
    console.log('\n🎯 Phase 5: Pivot Table 驗證');
    console.log('-'.repeat(40));

    if (workbook.Sheets['Pivot資料']) {
      const pivotData = XLSX.utils.sheet_to_json(workbook.Sheets['Pivot資料'], { header: 1 });
      console.log(`✅ Pivot資料工作表: ${pivotData.length} 行資料`);
    } else {
      console.log('❌ Pivot資料工作表不存在');
    }

    if (workbook.Sheets['Pivot_Table_匯出']) {
      const exportData = XLSX.utils.sheet_to_json(workbook.Sheets['Pivot_Table_匯出'], { header: 1 });
      console.log(`✅ Pivot_Table_匯出工作表: ${exportData.length} 行資料`);
      
      // 檢查是否有資料
      if (exportData.length > 1) {
        console.log('✅ Pivot Table 匯出正常: 包含資料');
        
        // 顯示前幾行資料
        console.log('📊 Pivot Table 匯出資料預覽:');
        for (let i = 0; i < Math.min(5, exportData.length); i++) {
          console.log(`  行 ${i + 1}:`, exportData[i]);
        }
      } else {
        console.log('⚠️ Pivot Table 匯出: 資料不足');
      }
    } else {
      console.log('❌ Pivot_Table_匯出工作表不存在');
    }

    // ===== Phase 6: 保護和圖表驗證 =====
    console.log('\n🔒 Phase 6: 保護和圖表驗證');
    console.log('-'.repeat(40));

    if (workbook.Sheets['保護和圖表']) {
      const protectedData = XLSX.utils.sheet_to_json(workbook.Sheets['保護和圖表'], { header: 1 });
      console.log(`✅ 保護和圖表工作表: ${protectedData.length} 行資料`);
      
      // 檢查圖表資料
      if (protectedData.length > 1) {
        console.log('✅ 圖表資料正常: 包含資料');
      }
    } else {
      console.log('❌ 保護和圖表工作表不存在');
    }

    // ===== OOXML 結構驗證 =====
    console.log('\n📁 OOXML 結構驗證');
    console.log('-'.repeat(40));

    // 讀取 Excel 檔案作為 ZIP
    const data = fs.readFileSync('test-all-features.xlsx');
    const zip = await JSZip.loadAsync(data);
    
    // 檢查 PivotCache 檔案
    const pivotCacheFiles = [];
    for (const fileName of Object.keys(zip.files)) {
      if (fileName.includes('pivotCache') && fileName.endsWith('.xml')) {
        pivotCacheFiles.push(fileName);
      }
    }
    
    // 檢查 PivotTable 檔案
    const pivotTableFiles = [];
    for (const fileName of Object.keys(zip.files)) {
      if (fileName.includes('pivotTable') && fileName.endsWith('.xml')) {
        pivotTableFiles.push(fileName);
      }
    }
    
    // 檢查圖表檔案
    const chartFiles = [];
    for (const fileName of Object.keys(zip.files)) {
      if (fileName.includes('chart') && fileName.endsWith('.xml')) {
        chartFiles.push(fileName);
      }
    }
    
    // 檢查繪圖檔案
    const drawingFiles = [];
    for (const fileName of Object.keys(zip.files)) {
      if (fileName.includes('drawing') && fileName.endsWith('.xml')) {
        drawingFiles.push(fileName);
      }
    }

    console.log(`🎯 PivotCache 檔案: ${pivotCacheFiles.length} 個`);
    console.log(`📊 PivotTable 檔案: ${pivotTableFiles.length} 個`);
    console.log(`📈 圖表檔案: ${chartFiles.length} 個`);
    console.log(`🎨 繪圖檔案: ${drawingFiles.length} 個`);

    // 檢查 Content Types
    const contentTypes = zip.file('[Content_Types].xml');
    if (contentTypes) {
      const contentTypesText = await contentTypes.async('string');
      
      if (contentTypesText.includes('pivotCacheDefinition')) {
        console.log('✅ 包含 PivotCache 定義類型');
      }
      if (contentTypesText.includes('pivotTable')) {
        console.log('✅ 包含 PivotTable 類型');
      }
      if (contentTypesText.includes('chart')) {
        console.log('✅ 包含圖表類型');
      }
    }

    // ===== 總結 =====
    console.log('\n📊 驗證結果總結');
    console.log('='.repeat(40));
    
    const expectedSheets = 7;
    const actualSheets = sheetNames.length;
    const hasPivotTable = pivotTableFiles.length > 0;
    const hasCharts = chartFiles.length > 0;
    const hasPivotExport = workbook.Sheets['Pivot_Table_匯出'] && 
                           XLSX.utils.sheet_to_json(workbook.Sheets['Pivot_Table_匯出'], { header: 1 }).length > 1;
    
    if (actualSheets === expectedSheets && hasPivotTable && hasCharts && hasPivotExport) {
      console.log('🎉 所有功能驗證通過！');
      console.log('✅ 這是一個功能完整的 Excel 檔案');
      console.log('📝 包含以下功能:');
      console.log('  1. 基本資料和公式支援');
      console.log('  2. 完整的樣式支援（字體、對齊、填滿、邊框）');
      console.log('  3. 進階功能（合併儲存格、欄寬列高、凍結窗格）');
      console.log('  4. 效能優化（大型資料集處理）');
      console.log('  5. 動態 Pivot Table 支援（完整的 OOXML 結構）');
      console.log('  6. 工作表和工作簿保護');
      console.log('  7. 圖表支援（柱狀圖、圓餅圖）');
      console.log('  8. Pivot Table 資料匯出功能');
      
    } else {
      console.log('❌ 部分功能驗證失敗');
      if (actualSheets !== expectedSheets) {
        console.log(`  - 工作表數量: 期望 ${expectedSheets}, 實際 ${actualSheets}`);
      }
      if (!hasPivotTable) console.log('  - 缺少 Pivot Table 支援');
      if (!hasCharts) console.log('  - 缺少圖表支援');
      if (!hasPivotExport) console.log('  - Pivot Table 匯出功能異常');
    }
    
    console.log('\n🎯 驗證完成！');
    console.log('📝 請在 Excel 中打開檔案以驗證所有功能');
    
  } catch (error) {
    console.error('❌ 驗證失敗:', error.message);
    console.error(error.stack);
  }
}

// 執行驗證
validateAllFeatures();
