import { Workbook } from '../dist/index.js';
import * as fs from 'fs';

interface SalesData {
  department: string;
  name: string;
  month: string;
  amount: number;
}

interface PivotResult {
  department: string;
  name: string;
  month1: number;
  month2: number;
  total: number;
}

async function fixedUsageExampleTS(): Promise<void> {
  console.log('🧪 修正版本的 xml-xlsx-lite TypeScript 使用範例');
  console.log('='.repeat(50));

  try {
    // 建立工作簿
    const wb = new Workbook();
    console.log('✅ 工作簿創建成功');
    
    // 建立數據表
    const ws = wb.getWorksheet('數據');
    console.log('✅ 數據工作表創建成功');
    
    // 測試數據 - 使用強型別
    const data: (string | number)[][] = [
      ['部門', '姓名', '月份', '銷售額'],
      ['A', '小明', '1月', 100],
      ['A', '小明', '2月', 120],
      ['A', '小華', '1月', 90],
      ['B', '小美', '1月', 200],
      ['B', '小美', '2月', 180],
      ['B', '小強', '1月', 150],
    ];
    
    // 寫入數據表 - 使用更安全的方式
    console.log('📝 寫入數據...');
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        const cellAddress = String.fromCharCode(65 + c) + (r + 1);
        const cellValue = data[r][c];
        
        // 為標題行添加樣式
        if (r === 0) {
          ws.setCell(cellAddress, cellValue, { 
            font: { bold: true },
            fill: { type: 'pattern', color: 'E0E0E0' }
          });
        } else {
          // 為數值欄位添加格式
          if (c === 3) { // 銷售額欄位
            ws.setCell(cellAddress, cellValue, { 
              numFmt: '#,##0',
              alignment: { horizontal: 'right' }
            });
          } else {
            ws.setCell(cellAddress, cellValue);
          }
        }
      }
    }
    
    // 設定欄寬
    ws.setColumnWidth('A', 12); // 部門
    ws.setColumnWidth('B', 12); // 姓名
    ws.setColumnWidth('C', 10); // 月份
    ws.setColumnWidth('D', 15); // 銷售額
    
    console.log('✅ 數據寫入完成');
    
    // 創建樞紐分析表結果工作表（手動方式，避免自動樞紐分析表問題）
    console.log('\n📊 創建樞紐分析表結果...');
    const pivotSheet = wb.getWorksheet('樞紐分析表');
    
    // 設定標題
    pivotSheet.setCell('A1', '銷售額樞紐分析表', {
      font: { bold: true, size: 16 },
      alignment: { horizontal: 'center' }
    });
    
    // 設定欄標題
    pivotSheet.setCell('A3', '部門', { font: { bold: true } });
    pivotSheet.setCell('B3', '姓名', { font: { bold: true } });
    pivotSheet.setCell('C3', '1月', { font: { bold: true } });
    pivotSheet.setCell('D3', '2月', { font: { bold: true } });
    pivotSheet.setCell('E3', '總計', { font: { bold: true } });
    
    // 手動計算樞紐分析表結果 - 使用強型別
    const pivotData: PivotResult[] = [
      { department: 'A', name: '小明', month1: 100, month2: 120, total: 220 },
      { department: 'A', name: '小華', month1: 90, month2: 0, total: 90 },
      { department: 'B', name: '小美', month1: 200, month2: 180, total: 380 },
      { department: 'B', name: '小強', month1: 150, month2: 0, total: 150 }
    ];
    
    // 填入樞紐分析表結果
    pivotData.forEach((row, index) => {
      const rowNum = index + 4;
      pivotSheet.setCell(`A${rowNum}`, row.department);
      pivotSheet.setCell(`B${rowNum}`, row.name);
      pivotSheet.setCell(`C${rowNum}`, row.month1, { 
        numFmt: '#,##0',
        alignment: { horizontal: 'right' }
      });
      pivotSheet.setCell(`D${rowNum}`, row.month2, { 
        numFmt: '#,##0',
        alignment: { horizontal: 'right' }
      });
      pivotSheet.setCell(`E${rowNum}`, row.total, { 
        numFmt: '#,##0',
        font: { bold: true },
        alignment: { horizontal: 'right' }
      });
    });
    
    // 設定欄寬
    pivotSheet.setColumnWidth('A', 12);
    pivotSheet.setColumnWidth('B', 12);
    pivotSheet.setColumnWidth('C', 12);
    pivotSheet.setColumnWidth('D', 12);
    pivotSheet.setColumnWidth('E', 15);
    
    console.log('✅ 樞紐分析表創建完成');
    
    // 創建摘要工作表
    console.log('\n📋 創建摘要工作表...');
    const summarySheet = wb.getWorksheet('摘要');
    
    // 計算摘要統計 - 使用強型別
    const salesData: SalesData[] = data.slice(1).map(row => ({
      department: row[0] as string,
      name: row[1] as string,
      month: row[2] as string,
      amount: row[3] as number
    }));
    
    const totalSales = salesData.reduce((sum, row) => sum + row.amount, 0);
    const deptA = salesData.filter(row => row.department === 'A').reduce((sum, row) => sum + row.amount, 0);
    const deptB = salesData.filter(row => row.department === 'B').reduce((sum, row) => sum + row.amount, 0);
    
    // 設定摘要內容
    summarySheet.setCell('A1', '銷售摘要報告', {
      font: { bold: true, size: 18 },
      alignment: { horizontal: 'center' }
    });
    
    summarySheet.setCell('A3', '總銷售額:', { font: { bold: true } });
    summarySheet.setCell('B3', totalSales, { 
      numFmt: '#,##0',
      font: { bold: true, size: 14 }
    });
    
    summarySheet.setCell('A4', 'A部門銷售額:', { font: { bold: true } });
    summarySheet.setCell('B4', deptA, { numFmt: '#,##0' });
    
    summarySheet.setCell('A5', 'B部門銷售額:', { font: { bold: true } });
    summarySheet.setCell('B5', deptB, { numFmt: '#,##0' });
    
    // 設定欄寬
    summarySheet.setColumnWidth('A', 20);
    summarySheet.setColumnWidth('B', 15);
    
    console.log('✅ 摘要工作表創建完成');
    
    // 使用 writeBuffer 方法輸出 Excel 檔案
    console.log('\n💾 輸出 Excel 檔案...');
    const buffer = await wb.writeBuffer();
    const filename = 'fixed-output-ts.xlsx';
    fs.writeFileSync(filename, new Uint8Array(buffer));
    console.log(`✅ Excel 檔案 ${filename} 已產生`);
    
    // 顯示檔案統計
    const stats = fs.statSync(filename);
    console.log(`📏 檔案大小: ${(stats.size / 1024).toFixed(2)} KB`);
    
    // 顯示工作表清單
    const worksheets = wb.getWorksheets();
    console.log(`📊 工作表數量: ${worksheets.length}`);
    console.log('\n📋 工作表清單:');
    worksheets.forEach((sheet, index) => {
      console.log(`  ${index + 1}. ${sheet.name}`);
    });
    
    console.log('\n🎉 修正版本 TypeScript 使用範例完成！');
    console.log('\n📝 解決的問題:');
    console.log('  1. ✅ 完全支援 TypeScript 型別檢查');
    console.log('  2. ✅ 使用正確的 writeBuffer 方法');
    console.log('  3. ✅ 手動創建樞紐分析表，避免自動樞紐分析表問題');
    console.log('  4. ✅ 添加了強型別介面定義');
    console.log('  5. ✅ 創建了多個工作表展示功能');
    
    console.log('\n🔍 樞紐分析表結果:');
    console.log(`  總銷售額: ${totalSales.toLocaleString()}`);
    console.log(`  A部門: ${deptA.toLocaleString()}`);
    console.log(`  B部門: ${deptB.toLocaleString()}`);
    
  } catch (error) {
    console.error('❌ 執行失敗:', error);
    if (error instanceof Error) {
      console.error('錯誤堆疊:', error.stack);
    }
    throw error;
  }
}

// 執行範例
fixedUsageExampleTS().catch(console.error);
