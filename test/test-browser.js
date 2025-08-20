const fs = require('fs');
const path = require('path');

// 建立瀏覽器測試用的 HTML 檔案
function createBrowserTestHTML() {
  const html = `<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>xml-xlsx-lite 瀏覽器測試</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        .test-section { margin: 20px 0; padding: 15px; border: 1px solid #ddd; border-radius: 5px; }
        .success { background-color: #d4edda; border-color: #c3e6cb; }
        .error { background-color: #f8d7da; border-color: #f5c6cb; }
        button { padding: 10px 20px; margin: 5px; background: #007bff; color: white; border: none; border-radius: 3px; cursor: pointer; }
        button:hover { background: #0056b3; }
        .log { background: #f8f9fa; padding: 10px; border-radius: 3px; margin: 10px 0; font-family: monospace; }
    </style>
</head>
<body>
    <h1>🧪 xml-xlsx-lite 瀏覽器測試</h1>
    
    <div class="test-section">
        <h3>基本功能測試</h3>
        <button onclick="testBasic()">執行基本測試</button>
        <div id="basic-result"></div>
    </div>
    
    <div class="test-section">
        <h3>多工作表測試</h3>
        <button onclick="testMultipleSheets()">執行多工作表測試</button>
        <div id="multiple-sheets-result"></div>
    </div>
    
    <div class="test-section">
        <h3>下載測試</h3>
        <button onclick="testDownload()">下載測試檔案</button>
        <div id="download-result"></div>
    </div>
    
    <div class="test-section">
        <h3>測試日誌</h3>
        <div id="log" class="log"></div>
    </div>

    <script type="module">
        import { Workbook } from '../dist/index.esm.js';
        
        window.Workbook = Workbook;
        
        // 日誌函數
        function log(message, type = 'info') {
            const logDiv = document.getElementById('log');
            const timestamp = new Date().toLocaleTimeString();
            const logEntry = document.createElement('div');
            logEntry.innerHTML = \`[\${timestamp}] \${message}\`;
            logDiv.appendChild(logEntry);
            console.log(message);
        }
        
        // 基本功能測試
        window.testBasic = async function() {
            const resultDiv = document.getElementById('basic-result');
            resultDiv.innerHTML = '<p>執行中...</p>';
            
            try {
                const wb = new Workbook();
                const ws = wb.getWorksheet("測試工作表");
                
                // 測試各種資料型別
                ws.setCell("A1", 123);
                ws.setCell("B2", "Hello World");
                ws.setCell("C3", true);
                ws.setCell("D4", new Date());
                ws.setCell("E5", "中文測試");
                
                log('✅ 儲存格設定成功');
                
                // 測試讀取
                const cellA1 = ws.getCell("A1");
                log(\`📊 A1 儲存格: \${cellA1.value}, 型別: \${cellA1.type}\`);
                
                // 生成檔案
                const buffer = await wb.writeBuffer();
                log(\`📁 檔案大小: \${buffer.byteLength} bytes\`);
                
                resultDiv.innerHTML = '<p class="success">✅ 基本測試成功！</p>';
                return true;
            } catch (error) {
                log(\`❌ 基本測試失敗: \${error.message}\`);
                resultDiv.innerHTML = \`<p class="error">❌ 測試失敗: \${error.message}</p>\`;
                return false;
            }
        };
        
        // 多工作表測試
        window.testMultipleSheets = async function() {
            const resultDiv = document.getElementById('multiple-sheets-result');
            resultDiv.innerHTML = '<p>執行中...</p>';
            
            try {
                const wb = new Workbook();
                
                // 建立多個工作表
                const ws1 = wb.getWorksheet("工作表1");
                const ws2 = wb.getWorksheet("工作表2");
                
                ws1.setCell("A1", "工作表1的資料");
                ws2.setCell("A1", "工作表2的資料");
                
                // 測試索引存取
                const wsByIndex = wb.getWorksheet(1);
                log(\`📋 工作表1名稱: \${wsByIndex.name}\`);
                
                const buffer = await wb.writeBuffer();
                log(\`✅ 多工作表測試成功，檔案大小: \${buffer.byteLength} bytes\`);
                
                resultDiv.innerHTML = '<p class="success">✅ 多工作表測試成功！</p>';
                return true;
            } catch (error) {
                log(\`❌ 多工作表測試失敗: \${error.message}\`);
                resultDiv.innerHTML = \`<p class="error">❌ 測試失敗: \${error.message}</p>\`;
                return false;
            }
        };
        
        // 下載測試
        window.testDownload = async function() {
            const resultDiv = document.getElementById('download-result');
            resultDiv.innerHTML = '<p>執行中...</p>';
            
            try {
                const wb = new Workbook();
                const ws = wb.getWorksheet("下載測試");
                
                // 建立一些測試資料
                ws.setCell("A1", "產品名稱");
                ws.setCell("B1", "價格");
                ws.setCell("C1", "數量");
                
                ws.setCell("A2", "蘋果");
                ws.setCell("B2", 25);
                ws.setCell("C2", 100);
                
                ws.setCell("A3", "香蕉");
                ws.setCell("B3", 15);
                ws.setCell("C3", 200);
                
                const buffer = await wb.writeBuffer();
                
                // 建立下載連結
                const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'browser-test.xlsx';
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                
                log('💾 檔案下載成功');
                resultDiv.innerHTML = '<p class="success">✅ 下載測試成功！</p>';
                return true;
            } catch (error) {
                log(\`❌ 下載測試失敗: \${error.message}\`);
                resultDiv.innerHTML = \`<p class="error">❌ 測試失敗: \${error.message}</p>\`;
                return false;
            }
        };
        
        log('🚀 瀏覽器測試頁面載入完成');
    </script>
</body>
</html>`;

  const testDir = path.join(__dirname, 'browser');
  if (!fs.existsSync(testDir)) {
    fs.mkdirSync(testDir, { recursive: true });
  }
  
  const htmlPath = path.join(testDir, 'test.html');
  fs.writeFileSync(htmlPath, html);
  
  console.log('🌐 瀏覽器測試 HTML 檔案已建立:', htmlPath);
  console.log('💡 請在瀏覽器中開啟此檔案進行測試');
}

// 執行
createBrowserTestHTML();
