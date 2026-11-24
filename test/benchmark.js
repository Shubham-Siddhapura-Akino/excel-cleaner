import path from 'path';
import fs from 'fs';
import ExcelJS from 'exceljs';
import { processWorkbook } from '../cleaner.js';

async function runBenchmark() {
  console.log('========================================');
  console.log('  EXCEL CLEANER - PERFORMANCE BENCHMARK ');
  console.log('========================================\n');

  const benchDir = path.join(process.cwd(), 'test', 'bench_data');
  if (!fs.existsSync(benchDir)) fs.mkdirSync(benchDir, { recursive: true });

  const inputPath = path.join(benchDir, 'large_input.xlsx');
  const outputPath = path.join(benchDir, 'large_output.xlsx');

  console.log('Generating large test workbook (10 Worksheets, 5,000 rows/sheet, 100,000+ cells)...');
  const wb = new ExcelJS.Workbook();

  const numSheets = 10;
  const rowsPerSheet = 5000;

  for (let s = 1; s <= numSheets; s++) {
    const ws = wb.addWorksheet(`DataSheet_${s}`);
    
    // Header
    ws.addRow(['ID', '  Full Name  ', '  Department\n  ', '  Base Salary ', '  Status  ']);
    ws.getRow(1).font = { bold: true };

    for (let r = 1; r <= rowsPerSheet; r++) {
      ws.addRow([
        r,
        `  User_${r}   Name_${s}  `,
        `  Department_${(r % 5) + 1}\t\n  `,
        1000 + r,
        `  Active_${r}  `
      ]);
    }
  }

  console.log('Writing benchmark workbook to disk...');
  await wb.xlsx.writeFile(inputPath);
  const fileSizeMb = (fs.statSync(inputPath).size / (1024 * 1024)).toFixed(2);
  console.log(`Generated workbook size: ${fileSizeMb} MB\n`);

  console.log('Running processWorkbook Benchmark...');
  const memBefore = process.memoryUsage().heapUsed / 1024 / 1024;
  
  const stats = await processWorkbook(inputPath, outputPath);

  const memAfter = process.memoryUsage().heapUsed / 1024 / 1024;
  const memDiff = (memAfter - memBefore).toFixed(2);

  console.log('\n--- Benchmark Results ---');
  console.log(`Sheets Processed: ${stats.sheetsProcessed}`);
  console.log(`Rows Processed:   ${stats.rowsProcessed}`);
  console.log(`Cells Visited:    ${stats.cellsVisited}`);
  console.log(`String Cells:     ${stats.stringCells}`);
  console.log(`Modified Cells:   ${stats.modifiedCells}`);
  console.log(`Execution Time:   ${stats.executionTime}`);
  console.log(`Heap Memory Delta: ${memDiff} MB`);
  console.log('--------------------------\n');

  // Clean up benchmark files
  fs.rmSync(benchDir, { recursive: true, force: true });

  console.log('========================================');
  console.log('  BENCHMARK COMPLETED SUCCESSFULLY!     ');
  console.log('========================================\n');
}

runBenchmark().catch((err) => {
  console.error('Benchmark error:', err);
  process.exit(1);
});
