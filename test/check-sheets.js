const XLSX = require('xlsx');

try {
  const workbook = XLSX.readFile('test-all-features.xlsx');
  console.log('工作表名稱:', workbook.SheetNames);
  
  workbook.SheetNames.forEach(name => {
    try {
      const data = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
      console.log(`${name}: ${data.length} 行`);
      
      if (data.length > 0) {
        console.log(`  第一行:`, data[0]);
        if (data.length > 1) {
          console.log(`  第二行:`, data[1]);
        }
      }
    } catch (error) {
      console.log(`${name}: 讀取失敗 - ${error.message}`);
    }
  });
} catch (error) {
  console.error('讀取檔案失敗:', error.message);
}
