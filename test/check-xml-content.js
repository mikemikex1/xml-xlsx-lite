const JSZip = require('jszip');
const fs = require('fs');

async function checkXmlContent() {
  console.log('🔍 檢查 Excel 檔案的內部 XML 內容');
  console.log('='.repeat(60));

  try {
    // 檢查 test-all-features.xlsx
    if (fs.existsSync('test-all-features.xlsx')) {
      console.log('\n📁 檢查 test-all-features.xlsx:');
      const buffer = fs.readFileSync('test-all-features.xlsx');
      const zip = await JSZip.loadAsync(buffer);
      
      console.log('📋 檔案列表:');
      const files = Object.keys(zip.files);
      files.forEach(file => {
        if (!file.endsWith('/')) {
          console.log(`  ${file}`);
        }
      });

      // 檢查工作表 XML
      console.log('\n📊 檢查工作表 XML:');
      for (let i = 1; i <= 2; i++) {
        const sheetFile = `xl/worksheets/sheet${i}.xml`;
        if (zip.file(sheetFile)) {
          const sheetXml = await zip.file(sheetFile).async('text');
          console.log(`\n${sheetFile}:`);
          console.log('前 500 字元:', sheetXml.substring(0, 500));
          if (sheetXml.length > 500) {
            console.log('... (檔案過長，只顯示前 500 字元)');
          }
        }
      }

      // 檢查 sharedStrings.xml
      console.log('\n📝 檢查 sharedStrings.xml:');
      if (zip.file('xl/sharedStrings.xml')) {
        const sstXml = await zip.file('xl/sharedStrings.xml').async('text');
        console.log('前 500 字元:', sstXml.substring(0, 500));
        if (sstXml.length > 500) {
          console.log('... (檔案過長，只顯示前 500 字元)');
        }
      }

      // 檢查 workbook.xml
      console.log('\n📚 檢查 workbook.xml:');
      if (zip.file('xl/workbook.xml')) {
        const workbookXml = await zip.file('xl/workbook.xml').async('text');
        console.log('前 500 字元:', workbookXml.substring(0, 500));
        if (workbookXml.length > 500) {
          console.log('... (檔案過長，只顯示前 500 字元)');
        }
      }
    }

    // 檢查 test-simple.xlsx 作為對比
    if (fs.existsSync('test-simple.xlsx')) {
      console.log('\n📁 檢查 test-simple.xlsx (對比):');
      const buffer = fs.readFileSync('test-simple.xlsx');
      const zip = await JSZip.loadAsync(buffer);
      
      console.log('📋 檔案列表:');
      const files = Object.keys(zip.files);
      files.forEach(file => {
        if (!file.endsWith('/')) {
          console.log(`  ${file}`);
        }
      });

      // 檢查工作表 XML
      console.log('\n📊 檢查工作表 XML:');
      const sheetFile = 'xl/worksheets/sheet1.xml';
      if (zip.file(sheetFile)) {
        const sheetXml = await zip.file(sheetFile).async('text');
        console.log(`\n${sheetFile}:`);
        console.log('前 500 字元:', sheetXml.substring(0, 500));
        if (sheetXml.length > 500) {
          console.log('... (檔案過長，只顯示前 500 字元)');
        }
      }
    }

    console.log('\n🎉 XML 內容檢查完成！');

  } catch (error) {
    console.error('❌ 檢查失敗:', error);
    console.error(error.stack);
  }
}

checkXmlContent();
