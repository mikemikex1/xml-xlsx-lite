/**
 * 驗證 xml-xlsx-lite 套件功能
 * 測試新的 API 和相容性方法
 */

const fs = require('fs');
const path = require('path');
const { Workbook, addPivotToWorkbookBuffer } = require('../dist/index.js');

(async () => {
  const outDir = path.resolve(__dirname, 'out');
  if (!fs.existsSync(outDir)) fs.mkdirSync(outDir);

  console.log('🧪 開始驗證 xml-xlsx-lite 套件功能...\n');

  try {
    // 1) 產生基本資料檔 base-workbook.xlsx（供動態 Pivot 使用）
    console.log('📊 1. 創建基礎工作簿...');
    const wb = new Workbook();
    const ws = wb.getWorksheet('Data');
    
    const rows = [
      ['Department', 'Month', 'Sales', 'Region'],
      ['A', 'Jan', 100, 'North'],
      ['A', 'Feb', 120, 'North'],
      ['B', 'Jan', 200, 'South'],
      ['B', 'Feb', 180, 'South'],
      ['B', 'Mar', 160, 'South'],
      ['C', 'Jan', 150, 'East'],
      ['C', 'Feb', 170, 'East'],
    ];
    
    rows.forEach((r, i) => r.forEach((v, j) => ws.setCell(String.fromCharCode(65+j) + (i+1), v)));
    
    // 創建樞紐分析表工作表
    const pivotWs = wb.getWorksheet('Pivot');
    pivotWs.setCell('A1', '樞紐分析表');
    
    const baseBuffer = await wb.writeBuffer();
    const basePath = path.join(outDir, 'base-workbook.xlsx');
    fs.writeFileSync(basePath, new Uint8Array(baseBuffer));
    console.log('✓ base-workbook.xlsx 已生成');

    // 2) 檢查是否有使用者期待的 API（一致性檢查）
    console.log('\n🔍 2. 檢查 API 一致性...');
    const lacksWriteFileWithPivot = !('writeFileWithPivotTables' in wb);
    console.log(lacksWriteFileWithPivot
      ? '✗ writeFileWithPivotTables 未找到（建議添加薄包裝）'
      : '✓ writeFileWithPivotTables 已找到');

    const lacksWriteFile = !('writeFile' in wb);
    console.log(lacksWriteFile
      ? '✗ writeFile 未找到（建議添加薄包裝）'
      : '✓ writeFile 已找到');

    const lacksCreateManualPivot = !('createManualPivotTable' in wb);
    console.log(lacksCreateManualPivot
      ? '✗ createManualPivotTable 未找到（建議添加手動樞紐 API）'
      : '✓ createManualPivotTable 已找到');

    // 3) 嘗試插入原生動態樞紐（依 README 的 addPivotToWorkbookBuffer）
    console.log('\n🔧 3. 測試動態樞紐分析表插入...');
    try {
      const pivotOptions = {
        sourceSheet: 'Data',
        sourceRange: 'A1:D100',
        targetSheet: 'Pivot',
        anchorCell: 'A3',
        layout: {
          rows: [{ name: 'Department' }],
          cols: [{ name: 'Month' }],
          values: [{ name: 'Sales', agg: 'sum', displayName: 'Total Sales' }]
        },
        refreshOnLoad: true,
        styleName: 'PivotStyleMedium9'
      };
      
      const enhanced = await addPivotToWorkbookBuffer(fs.readFileSync(basePath), pivotOptions);
      const dynPath = path.join(outDir, 'dynamic-pivot.xlsx');
      fs.writeFileSync(dynPath, enhanced);
      console.log('✓ dynamic-pivot.xlsx 已生成（在 Excel 中開啟並拖曳欄位驗證）');
    } catch (e) {
      console.error('✗ addPivotToWorkbookBuffer 失敗:', e?.message || e);
    }

    // 4) 測試手動樞紐分析表 API（如果存在）
    console.log('\n📈 4. 測試手動樞紐分析表 API...');
    if ('createManualPivotTable' in wb) {
      try {
        const data = [
          { Department: 'A', Month: 'Jan', Sales: 100, Region: 'North' },
          { Department: 'A', Month: 'Feb', Sales: 120, Region: 'North' },
          { Department: 'B', Month: 'Jan', Sales: 200, Region: 'South' },
          { Department: 'B', Month: 'Feb', Sales: 180, Region: 'South' },
          { Department: 'B', Month: 'Mar', Sales: 160, Region: 'South' },
          { Department: 'C', Month: 'Jan', Sales: 150, Region: 'East' },
          { Department: 'C', Month: 'Feb', Sales: 170, Region: 'East' },
        ];

        const result = wb.createManualPivotTable(data, {
          rowField: 'Department',
          columnField: 'Month',
          valueField: 'Sales',
          aggregation: 'sum',
          numberFormat: '#,##0',
          showRowTotals: true,
          showColumnTotals: true,
          sortBy: 'value',
          sortOrder: 'desc'
        });

        console.log('✓ 手動樞紐分析表已創建');
        console.log(`  統計: ${result.summary.totalRows} 行, ${result.summary.totalColumns} 列, 總值: ${result.summary.totalValue}`);

        // 保存手動樞紐分析表
        const manualBuffer = await wb.writeBuffer();
        const manualPath = path.join(outDir, 'manual-pivot.xlsx');
        fs.writeFileSync(manualPath, new Uint8Array(manualBuffer));
        console.log('✓ manual-pivot.xlsx 已生成');
      } catch (e) {
        console.error('✗ createManualPivotTable 失敗:', e?.message || e);
      }
    } else {
      // 手動創建樞紐分析表結果（純彙總）
      console.log('⚠️  使用手動方法創建樞紐分析表...');
      const wb2 = new Workbook();
      const dataSheet = wb2.getWorksheet('Data');
      rows.forEach((r, i) => r.forEach((v, j) => dataSheet.setCell(String.fromCharCode(65+j) + (i+1), v)));
      
      const manual = wb2.getWorksheet('Pivot Manual');
      manual.setCell('A1', 'Month'); 
      manual.setCell('B1', 'Dept A'); 
      manual.setCell('C1', 'Dept B'); 
      manual.setCell('D1', 'Dept C'); 
      manual.setCell('E1', 'Total');
      
      const calc = [
        ['Jan', 100, 200, 150, 450],
        ['Feb', 120, 180, 170, 470],
        ['Mar', 0, 160, 0, 160]
      ];
      
      calc.forEach((r, i) => {
        const row = i + 2;
        manual.setCell(`A${row}`, r[0]);
        manual.setCell(`B${row}`, r[1], { numFmt: '#,##0' });
        manual.setCell(`C${row}`, r[2], { numFmt: '#,##0' });
        manual.setCell(`D${row}`, r[3], { numFmt: '#,##0' });
        manual.setCell(`E${row}`, r[4], { numFmt: '#,##0', font: { bold: true } });
      });
      
      const manualPath = path.join(outDir, 'manual-pivot.xlsx');
      fs.writeFileSync(manualPath, new Uint8Array(await wb2.writeBuffer()));
      console.log('✓ manual-pivot.xlsx 已生成（手動方法）');
    }

    // 5) 測試新的 writeFile 方法（如果存在）
    console.log('\n💾 5. 測試新的 writeFile 方法...');
    if ('writeFile' in wb) {
      try {
        const testPath = path.join(outDir, 'test-writeFile.xlsx');
        await wb.writeFile(testPath);
        console.log('✓ writeFile 方法測試成功');
      } catch (e) {
        console.error('✗ writeFile 方法測試失敗:', e?.message || e);
      }
    }

    if ('writeFileWithPivotTables' in wb) {
      try {
        const testPath = path.join(outDir, 'test-writeFileWithPivot.xlsx');
        await wb.writeFileWithPivotTables(testPath, {
          sourceSheet: 'Data',
          sourceRange: 'A1:D100',
          targetSheet: 'Pivot',
          anchorCell: 'A3',
          layout: {
            rows: [{ name: 'Region' }],
            cols: [{ name: 'Month' }],
            values: [{ name: 'Sales', agg: 'sum' }]
          }
        });
        console.log('✓ writeFileWithPivotTables 方法測試成功');
      } catch (e) {
        console.error('✗ writeFileWithPivotTables 方法測試失敗:', e?.message || e);
      }
    }

    // 6) 生成完整範例（單檔含兩個原生 Pivot + 一個手動彙總）
    console.log('\n🚀 6. 生成完整範例...');
    try {
      const completeWb = new Workbook();
      const dataWs = completeWb.getWorksheet('Data');
      
      // 添加更多測試資料
      const completeData = [
        ['Product', 'Category', 'Region', 'Month', 'Sales', 'Quantity'],
        ['Laptop', 'Electronics', 'North', 'Jan', 1200, 5],
        ['Laptop', 'Electronics', 'North', 'Feb', 1400, 6],
        ['Laptop', 'Electronics', 'South', 'Jan', 1100, 4],
        ['Laptop', 'Electronics', 'South', 'Feb', 1300, 5],
        ['Phone', 'Electronics', 'North', 'Jan', 800, 8],
        ['Phone', 'Electronics', 'North', 'Feb', 900, 9],
        ['Phone', 'Electronics', 'South', 'Jan', 750, 7],
        ['Phone', 'Electronics', 'South', 'Feb', 850, 8],
        ['Book', 'Education', 'North', 'Jan', 200, 20],
        ['Book', 'Education', 'North', 'Feb', 250, 25],
        ['Book', 'Education', 'South', 'Jan', 180, 18],
        ['Book', 'Education', 'South', 'Feb', 220, 22],
      ];
      
      completeData.forEach((r, i) => r.forEach((v, j) => dataWs.setCell(String.fromCharCode(65+j) + (i+1), v)));
      
      // 創建手動樞紐分析表
      if ('createManualPivotTable' in completeWb) {
        const pivotData = completeData.slice(1).map(row => ({
          Product: row[0],
          Category: row[1],
          Region: row[2],
          Month: row[3],
          Sales: Number(row[4]),
          Quantity: Number(row[5])
        }));

        completeWb.createManualPivotTable(pivotData, {
          rowField: 'Category',
          columnField: 'Month',
          valueField: 'Sales',
          aggregation: 'sum',
          numberFormat: '#,##0'
        });
      }

      const completeBuffer = await completeWb.writeBuffer();
      const completePath = path.join(outDir, 'complete-example.xlsx');
      fs.writeFileSync(completePath, new Uint8Array(completeBuffer));
      console.log('✓ complete-example.xlsx 已生成');

    } catch (e) {
      console.error('✗ 完整範例生成失敗:', e?.message || e);
    }

    console.log('\n🎉 驗證完成！');
    console.log(`📁 輸出檔案位於: ${outDir}`);
    console.log('\n📋 驗收清單:');
    console.log('1. 在 Excel 中開啟 dynamic-pivot.xlsx，嘗試拖曳樞紐欄位');
    console.log('2. 檢查 manual-pivot.xlsx 的彙總結果是否正確');
    console.log('3. 驗證 complete-example.xlsx 的完整功能');

  } catch (error) {
    console.error('❌ 驗證過程發生錯誤:', error);
    console.error('錯誤詳情:', error.stack);
  }
})();
