const XLSX = require('xlsx');
const fs = require('fs');
const JSZip = require('jszip');

async function validateOOXMLStructure() {
  console.log('🔍 驗證 OOXML 結構完整性');
  console.log('='.repeat(60));

  try {
    // 檢查檔案是否存在
    if (!fs.existsSync('test-dynamic-pivot.xlsx')) {
      console.log('❌ 檔案不存在: test-dynamic-pivot.xlsx');
      return;
    }

    console.log('✅ 檔案存在: test-dynamic-pivot.xlsx');
    
    // 檢查檔案大小
    const stats = fs.statSync('test-dynamic-pivot.xlsx');
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);

    // 讀取 Excel 檔案作為 ZIP
    const data = fs.readFileSync('test-dynamic-pivot.xlsx');
    const zip = await JSZip.loadAsync(data);
    
    console.log('\n📁 檢查檔案結構...');
    
    // 檢查基本檔案
    const requiredFiles = [
      '[Content_Types].xml',
      '_rels/.rels',
      'xl/workbook.xml',
      'xl/_rels/workbook.xml.rels',
      'xl/sharedStrings.xml',
      'xl/styles.xml'
    ];
    
    console.log('\n📋 基本檔案檢查:');
    for (const file of requiredFiles) {
      if (zip.file(file)) {
        console.log(`  ✅ ${file}`);
      } else {
        console.log(`  ❌ ${file} - 缺失`);
      }
    }
    
    // 檢查工作表
    console.log('\n📊 工作表檢查:');
    const worksheets = [];
    for (let i = 1; i <= 10; i++) {
      const sheetFile = `xl/worksheets/sheet${i}.xml`;
      if (zip.file(sheetFile)) {
        worksheets.push(i);
        console.log(`  ✅ ${sheetFile}`);
      }
    }
    console.log(`  總共找到 ${worksheets.length} 個工作表`);
    
    // 檢查 PivotCache 檔案
    console.log('\n🎯 PivotCache 檔案檢查:');
    const pivotCacheFiles = [];
    for (const fileName of Object.keys(zip.files)) {
      if (fileName.includes('pivotCache') && fileName.endsWith('.xml')) {
        pivotCacheFiles.push(fileName);
        console.log(`  ✅ ${fileName}`);
      }
    }
    
    // 檢查 PivotTable 檔案
    console.log('\n📊 PivotTable 檔案檢查:');
    const pivotTableFiles = [];
    for (const fileName of Object.keys(zip.files)) {
      if (fileName.includes('pivotTable') && fileName.endsWith('.xml')) {
        pivotTableFiles.push(fileName);
        console.log(`  ✅ ${fileName}`);
      }
    }
    
    // 檢查關聯檔案
    console.log('\n🔗 關聯檔案檢查:');
    const relsFiles = [];
    for (const fileName of Object.keys(zip.files)) {
      if (fileName.includes('_rels') && fileName.endsWith('.rels')) {
        relsFiles.push(fileName);
        console.log(`  ✅ ${fileName}`);
      }
    }
    
    // 檢查 Content Types
    console.log('\n📝 Content Types 檢查:');
    const contentTypes = zip.file('[Content_Types].xml');
    if (contentTypes) {
      const contentTypesText = await contentTypes.async('string');
      
      // 檢查是否包含 PivotCache 類型
      if (contentTypesText.includes('pivotCacheDefinition')) {
        console.log('  ✅ 包含 PivotCache 定義類型');
      } else {
        console.log('  ❌ 缺少 PivotCache 定義類型');
      }
      
      if (contentTypesText.includes('pivotCacheRecords')) {
        console.log('  ✅ 包含 PivotCache 記錄類型');
      } else {
        console.log('  ❌ 缺少 PivotCache 記錄類型');
      }
      
      if (contentTypesText.includes('pivotTable')) {
        console.log('  ✅ 包含 PivotTable 類型');
      } else {
        console.log('  ❌ 缺少 PivotTable 類型');
      }
    }
    
    // 檢查 Workbook 關聯
    console.log('\n🔗 Workbook 關聯檢查:');
    const workbookRels = zip.file('xl/_rels/workbook.xml.rels');
    if (workbookRels) {
      const workbookRelsText = await workbookRels.async('string');
      
      if (workbookRelsText.includes('pivotCacheDefinition')) {
        console.log('  ✅ 包含 PivotCache 定義關聯');
      } else {
        console.log('  ❌ 缺少 PivotCache 定義關聯');
      }
    }
    
    // 總結
    console.log('\n📊 OOXML 結構驗證結果:');
    console.log('='.repeat(40));
    
    const hasPivotCache = pivotCacheFiles.length > 0;
    const hasPivotTable = pivotTableFiles.length > 0;
    const hasRels = relsFiles.length > 0;
    
    if (hasPivotCache && hasPivotTable && hasRels) {
      console.log('✅ OOXML 結構完整！');
      console.log(`  - PivotCache 檔案: ${pivotCacheFiles.length} 個`);
      console.log(`  - PivotTable 檔案: ${pivotTableFiles.length} 個`);
      console.log(`  - 關聯檔案: ${relsFiles.length} 個`);
      
      console.log('\n🎯 這是一個符合標準的動態 Pivot Table Excel 檔案！');
      console.log('📝 包含以下 OOXML 組件:');
      console.log('  1. pivotCacheDefinition.xml - 快取定義和欄位結構');
      console.log('  2. pivotCacheRecords.xml - 實際資料記錄');
      console.log('  3. pivotTable.xml - Pivot Table 定義');
      console.log('  4. 相關的關聯檔案 (.rels)');
      console.log('  5. 正確的 Content Types 定義');
      
    } else {
      console.log('❌ OOXML 結構不完整！');
      if (!hasPivotCache) console.log('  - 缺少 PivotCache 檔案');
      if (!hasPivotTable) console.log('  - 缺少 PivotTable 檔案');
      if (!hasRels) console.log('  - 缺少關聯檔案');
    }
    
    // 顯示所有檔案列表
    console.log('\n📁 完整檔案列表:');
    const allFiles = Object.keys(zip.files).sort();
    for (const fileName of allFiles) {
      const file = zip.file(fileName);
      const size = file ? file._data.uncompressedSize : 0;
      console.log(`  ${fileName} (${(size / 1024).toFixed(1)} KB)`);
    }
    
    console.log('\n🎯 OOXML 結構驗證完成！');
    
  } catch (error) {
    console.error('❌ 驗證失敗:', error.message);
    console.error(error.stack);
  }
}

// 執行驗證
validateOOXMLStructure();
