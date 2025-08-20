const fs = require('fs');
const path = require('path');

// 測試基本功能
async function testBasic() {
  console.log('🧪 測試基本功能...');
  
  try {
    // 動態導入（ESM 模組）
    const { Workbook } = await import('../dist/index.esm.js');
    
    const wb = new Workbook();
    const ws = wb.getWorksheet("測試工作表");
    
    // 測試各種資料型別
    ws.setCell("A1", 123);
    ws.setCell("B2", "Hello World");
    ws.setCell("C3", true);
    ws.setCell("D4", new Date());
    ws.setCell("E5", "中文測試");
    
    console.log('✅ 儲存格設定成功');
    
    // 測試讀取
    const cellA1 = ws.getCell("A1");
    console.log('📊 A1 儲存格:', cellA1.value, '型別:', cellA1.type);
    
    // 生成檔案
    const buffer = await wb.writeBuffer();
    console.log('📁 檔案大小:', buffer.byteLength, 'bytes');
    
    // 儲存到測試目錄
    const testDir = path.join(__dirname, 'output');
    if (!fs.existsSync(testDir)) {
      fs.mkdirSync(testDir, { recursive: true });
    }
    
    const outputPath = path.join(testDir, 'test-output.xlsx');
    fs.writeFileSync(outputPath, Buffer.from(buffer));
    console.log('💾 檔案已儲存到:', outputPath);
    
    return true;
  } catch (error) {
    console.error('❌ 測試失敗:', error);
    return false;
  }
}

// 測試多工作表
async function testMultipleSheets() {
  console.log('\n🧪 測試多工作表...');
  
  try {
    const { Workbook } = await import('../dist/index.esm.js');
    
    const wb = new Workbook();
    
    // 建立多個工作表
    const ws1 = wb.getWorksheet("工作表1");
    const ws2 = wb.getWorksheet("工作表2");
    
    ws1.setCell("A1", "工作表1的資料");
    ws2.setCell("A1", "工作表2的資料");
    
    // 測試索引存取
    const wsByIndex = wb.getWorksheet(1);
    console.log('📋 工作表1名稱:', wsByIndex.name);
    
    const buffer = await wb.writeBuffer();
    console.log('✅ 多工作表測試成功，檔案大小:', buffer.byteLength, 'bytes');
    
    return true;
  } catch (error) {
    console.error('❌ 多工作表測試失敗:', error);
    return false;
  }
}

// 測試便利方法
async function testConvenienceMethods() {
  console.log('\n🧪 測試便利方法...');
  
  try {
    const { Workbook } = await import('../dist/index.esm.js');
    
    const wb = new Workbook();
    
    // 使用便利方法
    wb.setCell("測試工作表", "A1", "便利方法測試");
    const cell = wb.getCell("測試工作表", "A1");
    
    console.log('✅ 便利方法測試成功:', cell.value);
    
    return true;
  } catch (error) {
    console.error('❌ 便利方法測試失敗:', error);
    return false;
  }
}

// 主測試函數
async function runTests() {
  console.log('🚀 開始執行 xml-xlsx-lite 測試...\n');
  
  const results = [];
  
  results.push(await testBasic());
  results.push(await testMultipleSheets());
  results.push(await testConvenienceMethods());
  
  const passed = results.filter(r => r).length;
  const total = results.length;
  
  console.log(`\n📊 測試結果: ${passed}/${total} 通過`);
  
  if (passed === total) {
    console.log('🎉 所有測試都通過了！');
    process.exit(0);
  } else {
    console.log('❌ 部分測試失敗');
    process.exit(1);
  }
}

// 執行測試
runTests().catch(error => {
  console.error('💥 測試執行錯誤:', error);
  process.exit(1);
});
