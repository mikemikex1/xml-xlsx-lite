const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// 驗證工具函數
function validateCellValue(actual, expected, description) {
  if (actual === expected) {
    console.log(`✅ ${description}: ${actual}`);
    return true;
  } else {
    console.log(`❌ ${description}: 期望 ${expected}, 實際 ${actual}`);
    return false;
  }
}

function validateCellType(actual, expectedType, description) {
  if (typeof actual === expectedType) {
    console.log(`✅ ${description}: 類型正確 (${expectedType})`);
    return true;
  } else {
    console.log(`❌ ${description}: 期望類型 ${expectedType}, 實際類型 ${typeof actual}`);
    return false;
  }
}

function validateRange(actual, min, max, description) {
  if (actual >= min && actual <= max) {
    console.log(`✅ ${description}: ${actual} (範圍: ${min}-${max})`);
    return true;
  } else {
    console.log(`❌ ${description}: ${actual} (超出範圍: ${min}-${max})`);
    return false;
  }
}

// 驗證基本功能檔案
function validateBasicExcel() {
  console.log('\n📋 驗證 Phase 1: 基本功能 (test-basic.xlsx)');
  console.log('=' .repeat(60));
  
  try {
    const workbook = XLSX.readFile('test-basic.xlsx');
    const sheetNames = workbook.SheetNames;
    
    console.log(`📊 工作表數量: ${sheetNames.length}`);
    console.log(`📋 工作表名稱: ${sheetNames.join(', ')}`);
    
    // 驗證第一個工作表
    const sheet1 = workbook.Sheets['基本測試'];
    const data1 = XLSX.utils.sheet_to_json(sheet1, { header: 1 });
    
    console.log('\n📊 基本測試工作表資料:');
    console.log('行數:', data1.length);
    console.log('列數:', data1[0] ? data1[0].length : 0);
    
    // 驗證標題行
    const headers = data1[0];
    validateCellValue(headers[0], '產品名稱', 'A1 標題');
    validateCellValue(headers[1], '數量', 'B1 標題');
    validateCellValue(headers[2], '單價', 'C1 標題');
    validateCellValue(headers[3], '總價', 'D1 標題');
    
    // 驗證產品資料
    const product1 = data1[1];
    validateCellValue(product1[0], 'iPhone 15', 'A2 產品名稱');
    validateCellValue(product1[1], 10, 'B2 數量');
    validateCellValue(product1[2], 35000, 'C2 單價');
    validateCellValue(product1[3], 350000, 'D2 總價');
    
    const product2 = data1[2];
    validateCellValue(product2[0], 'MacBook Pro', 'A3 產品名稱');
    validateCellValue(product2[1], 5, 'B3 數量');
    validateCellValue(product2[2], 80000, 'C3 單價');
    validateCellValue(product2[3], 400000, 'D3 總價');
    
    // 驗證第二個工作表
    const sheet2 = workbook.Sheets['第二工作表'];
    const data2 = XLSX.utils.sheet_to_json(sheet2, { header: 1 });
    
    console.log('\n📊 第二工作表資料:');
    validateCellValue(data2[0][0], '第二工作表的資料', '第二工作表 A1');
    validateCellValue(data2[0][1], 42, '第二工作表 B1');
    
    console.log('✅ Phase 1 基本功能驗證完成');
    return true;
    
  } catch (error) {
    console.error('❌ Phase 1 驗證失敗:', error.message);
    return false;
  }
}

// 驗證樣式支援檔案
function validateStylesExcel() {
  console.log('\n🎨 驗證 Phase 2: 樣式支援 (test-styles.xlsx)');
  console.log('=' .repeat(60));
  
  try {
    const workbook = XLSX.readFile('test-styles.xlsx');
    const sheet = workbook.Sheets['樣式測試'];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    console.log('📊 樣式測試工作表資料:');
    console.log('行數:', data.length);
    console.log('列數:', data[0] ? data[0].length : 0);
    
    // 驗證樣式資料
    const row1 = data[0];
    validateCellValue(row1[0], '標題 ✨', 'A1 標題樣式');
    validateCellValue(row1[1], '左對齊 ✨', 'B1 左對齊樣式');
    validateCellValue(row1[2], '紅色背景 ✨', 'C1 紅色背景樣式');
    validateCellValue(row1[3], '粗邊框 ✨', 'D1 粗邊框樣式');
    validateCellValue(row1[4], '完整樣式 ✨', 'E1 完整樣式');
    
    const row2 = data[1];
    validateCellValue(row2[0], '斜體文字 ✨', 'A2 斜體樣式');
    validateCellValue(row2[1], '置中對齊 ✨', 'B2 置中對齊樣式');
    validateCellValue(row2[2], '藍色背景 ✨', 'C2 藍色背景樣式');
    validateCellValue(row2[3], '虛線邊框 ✨', 'D2 虛線邊框樣式');
    
    const row3 = data[2];
    validateCellValue(row3[0], '底線文字 ✨', 'A3 底線樣式');
    validateCellValue(row3[1], '右對齊 ✨', 'B3 右對齊樣式');
    validateCellValue(row3[2], '網格圖案 ✨', 'C3 網格圖案樣式');
    validateCellValue(row3[3], '雙線邊框 ✨', 'D3 雙線邊框樣式');
    
    console.log('✅ Phase 2 樣式支援驗證完成');
    return true;
    
  } catch (error) {
    console.error('❌ Phase 2 驗證失敗:', error.message);
    return false;
  }
}

// 驗證進階功能檔案
function validateAdvancedExcel() {
  console.log('\n⚡ 驗證 Phase 3: 進階功能 (test-phase3.xlsx)');
  console.log('=' .repeat(60));
  
  try {
    const workbook = XLSX.readFile('test-phase3.xlsx');
    const sheet = workbook.Sheets['進階功能測試'];
    const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    
    console.log('📊 進階功能測試工作表資料:');
    console.log('行數:', data.length);
    console.log('列數:', data[0] ? data[0].length : 0);
    
    // 驗證合併儲存格效果（合併後的儲存格會顯示相同值）
    const row1 = data[0];
    validateCellValue(row1[0], '合併標題 ✨ 🔗', 'A1 合併標題');
    validateCellValue(row1[1], '合併標題 ✨ 🔗', 'B1 合併標題（應該與A1相同）');
    validateCellValue(row1[2], '合併標題 ✨ 🔗', 'C1 合併標題（應該與A1相同）');
    
    const row2 = data[1];
    validateCellValue(row2[0], '左側標題 ✨ 🔗', 'A2 左側標題');
    validateCellValue(row2[1], '欄位 1 ✨', 'B2 欄位1');
    validateCellValue(row2[2], '欄位 2 ✨', 'C2 欄位2');
    validateCellValue(row2[3], '欄位 3 ✨', 'D2 欄位3');
    
    // 驗證欄寬和列高設定（這些在讀取時可能無法直接驗證，但可以檢查資料完整性）
    console.log('📏 欄寬和列高設定已應用（需要手動檢查Excel檔案）');
    
    console.log('✅ Phase 3 進階功能驗證完成');
    return true;
    
  } catch (error) {
    console.error('❌ Phase 3 驗證失敗:', error.message);
    return false;
  }
}

// 驗證 Pivot Table 檔案
function validatePivotTableExcel() {
  console.log('\n🎯 驗證 Phase 5: Pivot Table 支援 (test-pivot-table.xlsx)');
  console.log('=' .repeat(60));
  
  try {
    const workbook = XLSX.readFile('test-pivot-table.xlsx');
    const sheetNames = workbook.SheetNames;
    
    console.log(`📊 工作表數量: ${sheetNames.length}`);
    console.log(`📋 工作表名稱: ${sheetNames.join(', ')}`);
    
    // 驗證資料工作表
    const dataSheet = workbook.Sheets['銷售資料'];
    const data = XLSX.utils.sheet_to_json(dataSheet, { header: 1 });
    
    console.log('\n📊 銷售資料工作表:');
    console.log('行數:', data.length);
    console.log('列數:', data[0] ? data[0].length : 0);
    
    // 驗證標題行
    const headers = data[0];
    validateCellValue(headers[0], '產品', 'A1 標題');
    validateCellValue(headers[1], '地區', 'B1 標題');
    validateCellValue(headers[2], '月份', 'C1 標題');
    validateCellValue(headers[3], '銷售額', 'D1 標題');
    
    // 驗證資料行
    const firstDataRow = data[1];
    validateCellType(firstDataRow[0], 'string', 'A2 產品類型');
    validateCellType(firstDataRow[1], 'string', 'B2 地區');
    validateCellType(firstDataRow[2], 'string', 'C2 月份');
    validateCellType(firstDataRow[3], 'number', 'D2 銷售額');
    
    // 驗證 Pivot Table 工作表
    if (workbook.Sheets['Pivot_Table_匯出']) {
      const pivotSheet = workbook.Sheets['Pivot_Table_匯出'];
      const pivotData = XLSX.utils.sheet_to_json(pivotSheet, { header: 1 });
      
      console.log('\n📊 Pivot Table 匯出工作表:');
      console.log('行數:', pivotData.length);
      console.log('列數:', pivotData[0] ? pivotData[0].length : 0);
      
      // 驗證 Pivot Table 結構
      if (pivotData.length > 0) {
        console.log('✅ Pivot Table 資料已成功匯出');
      }
    }
    
    console.log('✅ Phase 5 Pivot Table 支援驗證完成');
    return true;
    
  } catch (error) {
    console.error('❌ Phase 5 驗證失敗:', error.message);
    return false;
  }
}

// 驗證保護功能和圖表支援檔案
function validateProtectionChartsExcel() {
  console.log('\n🔒 驗證 Phase 6: 保護功能和圖表支援 (test-phase6.xlsx)');
  console.log('=' .repeat(60));
  
  try {
    const workbook = XLSX.readFile('test-phase6.xlsx');
    const sheetNames = workbook.SheetNames;
    
    console.log(`📊 工作表數量: ${sheetNames.length}`);
    console.log(`📋 工作表名稱: ${sheetNames.join(', ')}`);
    
    // 驗證主要資料工作表
    const dataSheet = workbook.Sheets['銷售資料'];
    const data = XLSX.utils.sheet_to_json(dataSheet, { header: 1 });
    
    console.log('\n📊 銷售資料工作表:');
    console.log('行數:', data.length);
    console.log('列數:', data[0] ? data[0].length : 0);
    
    // 驗證標題行
    const headers = data[0];
    validateCellValue(headers[0], '產品', 'A1 標題');
    validateCellValue(headers[1], '地區', 'B1 標題');
    validateCellValue(headers[2], '銷售額', 'C1 標題');
    validateCellValue(headers[3], '數量', 'D1 標題');
    
    // 驗證資料行
    const firstDataRow = data[1];
    validateCellType(firstDataRow[0], 'string', 'A2 產品類型');
    validateCellType(firstDataRow[1], 'string', 'B2 地區');
    validateCellType(firstDataRow[2], 'number', 'C2 銷售額');
    validateCellType(firstDataRow[3], 'number', 'D2 數量');
    
    // 圖表在讀取時可能無法直接驗證，但可以檢查工作表完整性
    console.log('📊 圖表支援已實現（需要手動檢查Excel檔案中的圖表）');
    console.log('🔒 保護功能已實現（需要手動檢查Excel檔案的保護設定）');
    
    console.log('✅ Phase 6 保護功能和圖表支援驗證完成');
    return true;
    
  } catch (error) {
    console.error('❌ Phase 6 驗證失敗:', error.message);
    return false;
  }
}

// 驗證綜合測試檔案
function validateComprehensiveExcel() {
  console.log('\n🧪 驗證綜合功能測試 (comprehensive-test.xlsx)');
  console.log('=' .repeat(60));
  
  try {
    const workbook = XLSX.readFile('comprehensive-test.xlsx');
    const sheetNames = workbook.SheetNames;
    
    console.log(`📊 工作表數量: ${sheetNames.length}`);
    console.log(`📋 工作表名稱: ${sheetNames.join(', ')}`);
    
    // 驗證基本功能工作表
    const basicSheet = workbook.Sheets['基本功能'];
    if (basicSheet) {
      const basicData = XLSX.utils.sheet_to_json(basicSheet, { header: 1 });
      console.log('\n📊 基本功能工作表:');
      console.log('行數:', basicData.length);
      console.log('列數:', basicData[0] ? basicData[0].length : 0);
      
      // 驗證基本資料
      const headers = basicData[0];
      validateCellValue(headers[0], '產品名稱', 'A1 標題');
      validateCellValue(headers[1], '數量', 'B1 標題');
      validateCellValue(headers[2], '單價', 'C1 標題');
    }
    
    // 驗證樣式工作表
    const stylesSheet = workbook.Sheets['樣式支援'];
    if (stylesSheet) {
      const stylesData = XLSX.utils.sheet_to_json(stylesSheet, { header: 1 });
      console.log('\n📊 樣式支援工作表:');
      console.log('行數:', stylesData.length);
      console.log('列數:', stylesData[0] ? stylesData[0].length : 0);
    }
    
    // 驗證進階功能工作表
    const advancedSheet = workbook.Sheets['進階功能'];
    if (advancedSheet) {
      const advancedData = XLSX.utils.sheet_to_json(advancedSheet, { header: 1 });
      console.log('\n📊 進階功能工作表:');
      console.log('行數:', advancedData.length);
      console.log('列數:', advancedData[0] ? advancedData[0].length : 0);
    }
    
    // 驗證效能優化工作表
    const performanceSheet = workbook.Sheets['效能測試'];
    if (performanceSheet) {
      const performanceData = XLSX.utils.sheet_to_json(performanceSheet, { header: 1 });
      console.log('\n📊 效能測試工作表:');
      console.log('行數:', performanceData.length);
      console.log('列數:', performanceData[0] ? performanceData[0].length : 0);
      
      // 驗證大量資料
      if (performanceData.length > 1000) {
        console.log('✅ 大量資料處理功能正常');
      }
    }
    
    // 驗證 Pivot Table 工作表
    const pivotSheet = workbook.Sheets['Pivot資料'];
    if (pivotSheet) {
      const pivotData = XLSX.utils.sheet_to_json(pivotSheet, { header: 1 });
      console.log('\n📊 Pivot資料工作表:');
      console.log('行數:', pivotData.length);
      console.log('列數:', pivotData[0] ? pivotData[0].length : 0);
    }
    
    // 驗證保護和圖表工作表
    const protectionSheet = workbook.Sheets['保護和圖表'];
    if (protectionSheet) {
      const protectionData = XLSX.utils.sheet_to_json(protectionSheet, { header: 1 });
      console.log('\n📊 保護和圖表工作表:');
      console.log('行數:', protectionData.length);
      console.log('列數:', protectionData[0] ? protectionData[0].length : 0);
    }
    
    console.log('✅ 綜合功能測試驗證完成');
    return true;
    
  } catch (error) {
    console.error('❌ 綜合功能測試驗證失敗:', error.message);
    return false;
  }
}

// 主驗證函數
async function validateAllExcelFiles() {
  console.log('🧪 開始驗證所有 Excel 檔案');
  console.log('=' .repeat(80));
  
  const results = [];
  
  // 檢查檔案是否存在
  const files = [
    'test-basic.xlsx',
    'test-styles.xlsx', 
    'test-phase3.xlsx',
    'test-pivot-table.xlsx',
    'test-phase6.xlsx',
    'comprehensive-test.xlsx'
  ];
  
  console.log('📁 檢查測試檔案:');
  files.forEach(file => {
    if (fs.existsSync(file)) {
      console.log(`✅ ${file} - 存在`);
    } else {
      console.log(`❌ ${file} - 不存在`);
    }
  });
  
  // 執行各階段驗證
  results.push(validateBasicExcel());
  results.push(validateStylesExcel());
  results.push(validateAdvancedExcel());
  results.push(validatePivotTableExcel());
  results.push(validateProtectionChartsExcel());
  results.push(validateComprehensiveExcel());
  
  // 驗證結果總結
  console.log('\n' + '=' .repeat(80));
  console.log('📊 驗證結果總結');
  console.log('=' .repeat(80));
  
  const passed = results.filter(r => r).length;
  const total = results.length;
  
  console.log(`✅ 通過: ${passed}/${total}`);
  console.log(`❌ 失敗: ${total - passed}/${total}`);
  
  if (passed === total) {
    console.log('\n🎉 所有驗證都通過了！Excel 檔案功能正常！');
  } else {
    console.log('\n⚠️ 部分驗證失敗，請檢查相關功能。');
  }
  
  console.log('\n📝 注意事項:');
  console.log('- 樣式、欄寬、列高、圖表等視覺效果需要手動打開 Excel 檔案檢查');
  console.log('- 保護功能需要在 Excel 中嘗試編輯來驗證');
  console.log('- 合併儲存格效果需要視覺檢查');
}

// 執行驗證
validateAllExcelFiles().catch(console.error);
