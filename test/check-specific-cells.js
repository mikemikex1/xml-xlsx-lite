const JSZip = require('jszip');
const fs = require('fs');

async function checkSpecificCells() {
  console.log('🔍 檢查特定儲存格的 XML 內容');
  console.log('='.repeat(50));

  try {
    if (fs.existsSync('test-all-features.xlsx')) {
      const buffer = fs.readFileSync('test-all-features.xlsx');
      const zip = await JSZip.loadAsync(buffer);
      
      console.log('✅ test-all-features.xlsx 讀取成功');
      
      // 檢查基本功能工作表的 D2 儲存格（應該有公式）
      const sheet1Xml = await zip.file('xl/worksheets/sheet1.xml').async('text');
      console.log('\n📊 基本功能工作表 XML (尋找 D2 儲存格):');
      
      // 尋找包含 D2 的行
      const lines = sheet1Xml.split('\n');
      for (const line of lines) {
        if (line.includes('r="D2"')) {
          console.log('找到 D2 儲存格:', line.trim());
          break;
        }
      }
      
      // 檢查進階功能工作表的 C2 和 C3 儲存格（應該有公式）
      const sheet3Xml = await zip.file('xl/worksheets/sheet3.xml').async('text');
      console.log('\n📊 進階功能工作表 XML (尋找 C2 和 C3 儲存格):');
      
      for (const line of lines) {
        if (line.includes('r="C2"') || line.includes('r="C3"')) {
          console.log('找到儲存格:', line.trim());
        }
      }
      
      // 檢查 sharedStrings.xml 是否包含公式
      const sstXml = await zip.file('xl/sharedStrings.xml').async('text');
      console.log('\n📝 檢查 sharedStrings.xml 是否包含公式:');
      
      if (sstXml.includes('=B2*C2')) {
        console.log('❌ 發現公式字串 =B2*C2');
      } else {
        console.log('✅ 沒有發現公式字串 =B2*C2');
      }
      
      if (sstXml.includes('=B3*C3')) {
        console.log('❌ 發現公式字串 =B3*C3');
      } else {
        console.log('✅ 沒有發現公式字串 =B3*C3');
      }
      
      if (sstXml.includes('=B2*2')) {
        console.log('❌ 發現公式字串 =B2*2');
      } else {
        console.log('✅ 沒有發現公式字串 =B2*2');
      }
      
      if (sstXml.includes('=SUM(B2:B3)')) {
        console.log('❌ 發現公式字串 =SUM(B2:B3)');
      } else {
        console.log('✅ 沒有發現公式字串 =SUM(B2:B3)');
      }
      
    } else {
      console.log('❌ 檔案不存在: test-all-features.xlsx');
    }
    
    console.log('\n🎉 特定儲存格檢查完成！');
    
  } catch (error) {
    console.error('❌ 檢查失敗:', error);
  }
}

checkSpecificCells();
