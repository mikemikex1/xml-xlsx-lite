const XLSX = require('xlsx');
const fs = require('fs');

function checkFormulas() {
  console.log('🔍 檢查公式支援問題');
  console.log('='.repeat(50));
  
  try {
    if (!fs.existsSync('test-all-features.xlsx')) {
      console.log('❌ 檔案不存在: test-all-features.xlsx');
      return;
    }
    
    const workbook = XLSX.readFile('test-all-features.xlsx');
    const basicSheet = workbook.Sheets['基本功能'];
    
    if (!basicSheet) {
      console.log('❌ 基本功能工作表不存在');
      return;
    }
    
    console.log('✅ 基本功能工作表讀取成功');
    
    // 檢查特定儲存格
    const a1 = basicSheet['A1'];
    const b2 = basicSheet['B2'];
    const c2 = basicSheet['C2'];
    const d2 = basicSheet['D2'];
    
    console.log('\n📊 儲存格檢查:');
    console.log(`A1: ${a1 ? a1.v : 'undefined'}`);
    console.log(`B2: ${b2 ? b2.v : 'undefined'}`);
    console.log(`C2: ${c2 ? c2.v : 'undefined'}`);
    console.log(`D2: ${d2 ? d2.v : 'undefined'}`);
    
    // 檢查公式
    if (d2 && d2.f) {
      console.log(`D2 公式: ${d2.f}`);
    } else {
      console.log('D2 沒有公式');
    }
    
    // 檢查所有儲存格
    console.log('\n🔍 所有儲存格檢查:');
    const cellRefs = Object.keys(basicSheet);
    for (const ref of cellRefs) {
      if (ref !== '!ref' && ref !== '!margins' && ref !== '!cols' && ref !== '!rows') {
        const cell = basicSheet[ref];
        if (cell.f) {
          console.log(`${ref}: 公式 = ${cell.f}`);
        }
      }
    }
    
    console.log('\n🎉 公式檢查完成！');
    
  } catch (error) {
    console.error('❌ 檢查失敗:', error);
  }
}

checkFormulas();
