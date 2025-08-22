/**
 * 測試動態樞紐分析表功能
 * 展示如何在既有 Excel 檔案上動態插入樞紐分析表
 */

const { Workbook, addPivotToWorkbookBuffer, CreatePivotOptions } = require('../dist/index.js');
const fs = require('fs');

async function testDynamicPivot() {
  console.log('🧪 測試動態樞紐分析表功能...');
  
  try {
    // 步驟 1: 創建基礎工作簿（包含資料和空白樞紐分析表工作表）
    console.log('📝 1. 創建基礎工作簿...');
    
    const wb = new Workbook();
    
    // 創建資料工作表
    const dataWs = wb.getWorksheet('數據');
    
    // 添加標題行
    dataWs.setCell('A1', '部門', { font: { bold: true } });
    dataWs.setCell('B1', '月份', { font: { bold: true } });
    dataWs.setCell('C1', '產品', { font: { bold: true } });
    dataWs.setCell('D1', '銷售額', { font: { bold: true } });
    
    // 添加測試資料
    const testData = [
      ['IT', '一月', '軟體', 50000],
      ['IT', '一月', '硬體', 30000],
      ['IT', '二月', '軟體', 60000],
      ['IT', '二月', '硬體', 35000],
      ['HR', '一月', '培訓', 20000],
      ['HR', '一月', '招募', 15000],
      ['HR', '二月', '培訓', 25000],
      ['HR', '二月', '招募', 18000],
      ['財務', '一月', '審計', 40000],
      ['財務', '一月', '稅務', 25000],
      ['財務', '二月', '審計', 45000],
      ['財務', '二月', '稅務', 30000]
    ];
    
    // 寫入資料
    for (let i = 0; i < testData.length; i++) {
      const row = testData[i];
      dataWs.setCell(`A${i + 2}`, row[0]);
      dataWs.setCell(`B${i + 2}`, row[1]);
      dataWs.setCell(`C${i + 2}`, row[2]);
      dataWs.setCell(`D${i + 2}`, row[3], { numFmt: '#,##0' });
    }
    
    // 設定欄寬
    dataWs.setColumnWidth('A', 15);
    dataWs.setColumnWidth('B', 12);
    dataWs.setColumnWidth('C', 15);
    dataWs.setColumnWidth('D', 15);
    
    // 創建空白樞紐分析表工作表
    const pivotWs = wb.getWorksheet('Pivot');
    
    // 添加標題
    pivotWs.setCell('A1', '樞紐分析表', { font: { bold: true, size: 16 } });
    pivotWs.setCell('A2', '（此處將插入動態樞紐分析表）', { font: { italic: true, color: '808080' } });
    
    // 設定欄寬
    pivotWs.setColumnWidth('A', 30);
    
    console.log('✅ 基礎工作簿創建完成');
    
    // 步驟 2: 輸出基礎 Excel 檔案
    console.log('\n💾 2. 輸出基礎 Excel 檔案...');
    
    const baseBuffer = await wb.writeBuffer();
    const baseFilename = 'base-workbook.xlsx';
    fs.writeFileSync(baseFilename, new Uint8Array(baseBuffer));
    
    console.log(`✅ 基礎檔案 ${baseFilename} 已產生`);
    console.log('📊 檔案大小:', (baseBuffer.byteLength / 1024).toFixed(2), 'KB');
    
    // 步驟 3: 使用動態樞紐分析表建構器
    console.log('\n🔧 3. 動態插入樞紐分析表...');
    
         const pivotOptions = {
      sourceSheet: "數據",
      sourceRange: "A1:D13",         // 含標題列
      targetSheet: "Pivot",
      anchorCell: "A3",
      layout: {
        rows: [{ name: "部門" }],
        cols: [{ name: "月份" }],
        values: [
          { 
            name: "銷售額", 
            agg: "sum", 
            displayName: "銷售額合計",
            numFmtId: 0
          }
        ],
      },
      refreshOnLoad: true,
      styleName: "PivotStyleMedium9",
    };
    
    console.log('📋 樞紐分析表配置:');
    console.log(`  來源工作表: ${pivotOptions.sourceSheet}`);
    console.log(`  來源範圍: ${pivotOptions.sourceRange}`);
    console.log(`  目標工作表: ${pivotOptions.targetSheet}`);
    console.log(`  錨點儲存格: ${pivotOptions.anchorCell}`);
    console.log(`  行欄位: ${pivotOptions.layout.rows?.map(f => f.name).join(', ') || '無'}`);
    console.log(`  列欄位: ${pivotOptions.layout.cols?.map(f => f.name).join(', ') || '無'}`);
    console.log(`  值欄位: ${pivotOptions.layout.values.map(v => `${v.name}(${v.agg})`).join(', ')}`);
    
    // 動態插入樞紐分析表
    const enhancedBuffer = await addPivotToWorkbookBuffer(baseBuffer, pivotOptions);
    
    console.log('✅ 樞紐分析表插入完成');
    
    // 步驟 4: 輸出最終檔案
    console.log('\n💾 4. 輸出最終 Excel 檔案...');
    
    const finalFilename = 'dynamic-pivot-workbook.xlsx';
    fs.writeFileSync(finalFilename, new Uint8Array(enhancedBuffer));
    
    console.log(`✅ 最終檔案 ${finalFilename} 已產生`);
    console.log('📊 檔案大小:', (enhancedBuffer.byteLength / 1024).toFixed(2), 'KB');
    console.log('📈 檔案大小變化:', ((enhancedBuffer.byteLength - baseBuffer.byteLength) / 1024).toFixed(2), 'KB');
    
    // 步驟 5: 驗證結果
    console.log('\n🔍 5. 驗證結果...');
    
    // 檢查檔案是否存在
    if (fs.existsSync(finalFilename)) {
      console.log('✅ 最終檔案存在');
      
      // 檢查檔案大小
      const stats = fs.statSync(finalFilename);
      console.log(`✅ 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);
      
      // 檢查檔案是否為有效的 ZIP 檔案（XLSX 本質上是 ZIP）
      try {
        const testZip = require('jszip');
        const testBuffer = fs.readFileSync(finalFilename);
        const zip = await testZip.loadAsync(testBuffer);
        
        // 檢查是否包含樞紐分析表相關檔案
        const hasPivotCache = zip.file(/pivotCache\/pivotCacheDefinition.*\.xml/).length > 0;
        const hasPivotTable = zip.file(/pivotTables\/pivotTable.*\.xml/).length > 0;
        const hasContentTypes = zip.file('[Content_Types].xml').length > 0;
        
        console.log('✅ 檔案結構驗證:');
        console.log(`  樞紐分析表快取定義: ${hasPivotCache ? '✅' : '❌'}`);
        console.log(`  樞紐分析表定義: ${hasPivotTable ? '✅' : '❌'}`);
        console.log(`  Content Types: ${hasContentTypes ? '✅' : '❌'}`);
        
        if (hasPivotCache && hasPivotTable && hasContentTypes) {
          console.log('🎉 所有必要檔案都已正確創建！');
        }
        
      } catch (zipError) {
        console.log('⚠️ 無法驗證 ZIP 結構:', zipError.message);
      }
      
    } else {
      console.log('❌ 最終檔案不存在');
    }
    
    console.log('\n🎯 動態樞紐分析表測試完成！');
    console.log('請打開 Excel 檔案檢查樞紐分析表是否正確顯示。');
    console.log('樞紐分析表應該出現在 Pivot 工作表的 A3 位置。');
    console.log('修改「數據」工作表的資料後，可以在樞紐分析表上按右鍵選擇「重新整理」來更新。');
    
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    console.error('錯誤詳情:', error.stack);
  }
}

// 執行測試
testDynamicPivot().catch(console.error);
