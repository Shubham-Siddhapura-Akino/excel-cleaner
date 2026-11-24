import fs from 'fs';
import path from 'path';
import assert from 'assert';
import ExcelJS from 'exceljs';
import { cleanText } from '../utils/cleanText.js';
import { processWorkbook, processWorksheet } from '../cleaner.js';
import { resolveFilePaths } from '../index.js';

console.log('========================================');
console.log('  EXCEL CLEANER - TEST SUITE RUNNER     ');
console.log('========================================\n');

// --------------------------------------------------------------------
// 1. Unit Tests for cleanText Utility
// --------------------------------------------------------------------
console.log('[1/3] Running Unit Tests for cleanText()...');

const testCases = [
  { input: '  Shubham     Siddhapura ', expected: 'Shubham Siddhapura' },
  { input: 'John\nSmith', expected: 'John Smith' },
  { input: '    John      Smith', expected: 'John Smith' },
  { input: '\t\tHeader\r\nValue\u00A0\u200B ', expected: 'Header Value' },
  { input: '"Quoted Text"', expected: 'Quoted Text' },
  { input: '  "  Shubham     Siddhapura " ', expected: 'Shubham Siddhapura' },
  { input: 'NoChangeNeeded', expected: 'NoChangeNeeded' },
  { input: '', expected: '' },
  { input: 12345, expected: 12345 },
  { input: true, expected: true },
  { input: null, expected: null }
];

let unitTestsPassed = 0;
for (const tc of testCases) {
  const result = cleanText(tc.input);
  assert.strictEqual(
    result,
    tc.expected,
    `Failed cleanText test for input: ${JSON.stringify(tc.input)}. Got: ${JSON.stringify(result)}, Expected: ${JSON.stringify(tc.expected)}`
  );
  unitTestsPassed++;
}
console.log(`  ✓ Passed ${unitTestsPassed}/${testCases.length} unit test cases.\n`);

// --------------------------------------------------------------------
// 2. Unit Tests for CLI Path Resolution
// --------------------------------------------------------------------
console.log('[2/3] Running Unit Tests for resolveFilePaths()...');

const pathDefault = resolveFilePaths([]);
assert.ok(pathDefault.inputPath.endsWith('input/workbook.xlsx'));
assert.ok(pathDefault.outputPath.endsWith('output/workbook_cleaned.xlsx'));

const pathSingle = resolveFilePaths(['data/my_sheet.xlsx']);
assert.strictEqual(pathSingle.inputPath, 'data/my_sheet.xlsx');
assert.strictEqual(pathSingle.outputPath, path.join('data', 'my_sheet_cleaned.xlsx'));

const pathDouble = resolveFilePaths(['in.xlsx', 'out.xlsx']);
assert.strictEqual(pathDouble.inputPath, 'in.xlsx');
assert.strictEqual(pathDouble.outputPath, 'out.xlsx');

console.log('  ✓ Passed CLI argument path resolution tests.\n');

// --------------------------------------------------------------------
// 3. End-to-End Workbook Preservation & Cleaning Test
// --------------------------------------------------------------------
console.log('[3/3] Creating Sample Excel Workbook & Verifying Preservation...');

const inputDir = path.join(process.cwd(), 'input');
const outputDir = path.join(process.cwd(), 'output');
if (!fs.existsSync(inputDir)) fs.mkdirSync(inputDir, { recursive: true });
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

const sampleInputPath = path.join(inputDir, 'test_workbook.xlsx');
const sampleOutputPath = path.join(outputDir, 'test_workbook_cleaned.xlsx');

async function createSampleWorkbook() {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Cleaner Test';

  // --- Sheet 1: Sheet1 (Formatting, String Cleaning, Merged Cells) ---
  const ws1 = wb.addWorksheet('Sheet1');

  // Views / Freeze panes
  ws1.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];

  // Column widths
  ws1.getColumn(1).width = 25;
  ws1.getColumn(2).width = 30;

  // Header Row with Styling
  const row1 = ws1.getRow(1);
  row1.height = 30;

  const cellA1 = ws1.getCell('A1');
  cellA1.value = '  Full Name  ';
  cellA1.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  cellA1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };
  cellA1.alignment = { vertical: 'middle', horizontal: 'center' };
  cellA1.border = {
    top: { style: 'thin', color: { argb: 'FF000000' } },
    bottom: { style: 'medium', color: { argb: 'FF000000' } }
  };

  const cellB1 = ws1.getCell('B1');
  cellB1.value = '  Department\n  ';
  cellB1.font = { name: 'Arial', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
  cellB1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E79' } };

  // Data Rows
  ws1.getCell('A2').value = '  Shubham     Siddhapura ';
  ws1.getCell('B2').value = ' Engineering ';

  ws1.getCell('A3').value = 'John\nSmith';
  ws1.getCell('B3').value = ' Human\tResources ';

  ws1.getCell('A4').value = '    John      Smith';
  ws1.getCell('B4').value = ' Marketing ';

  // Numeric & Date cells
  ws1.getCell('A5').value = 9999;
  ws1.getCell('B5').value = new Date('2026-08-03');
  ws1.getCell('B5').numFmt = 'yyyy-mm-dd';

  // Merged Cells
  ws1.mergeCells('A10:B10');
  ws1.getCell('A10').value = '  Merged     Section Header ';
  ws1.getCell('A10').font = { bold: true, color: { argb: 'FF008000' } };

  // Hidden Row & Column
  ws1.getRow(6).hidden = true;
  ws1.getCell('A6').value = ' Hidden Row Text ';
  ws1.getColumn(4).hidden = true;

  // AutoFilter
  ws1.autoFilter = 'A1:B5';

  // --- Sheet 2: Employees (Formulas, Comments, Hyperlinks) ---
  const ws2 = wb.addWorksheet('Employees');

  ws2.getCell('A1').value = '  Employee ID ';
  ws2.getCell('B1').value = ' Base Salary ';

  ws2.getCell('A2').value = ' EMP-001 ';
  ws2.getCell('B2').value = 50000;

  ws2.getCell('A3').value = ' EMP-002 ';
  ws2.getCell('B3').value = 60000;

  // Formula Cell
  ws2.getCell('A4').value = ' Total Salary ';
  ws2.getCell('B4').value = { formula: 'SUM(B2:B3)', result: 110000 };
  ws2.getCell('B4').numFmt = '$#,##0.00';

  // Cell Comment / Note
  ws2.getCell('A2').note = 'Unique employee identification string';

  // Hyperlink Cell (Preserve)
  ws2.getCell('A5').value = { text: ' Portal Link ', hyperlink: 'https://company.internal' };

  await wb.xlsx.writeFile(sampleInputPath);
  console.log(`  ✓ Sample input workbook written to: ${sampleInputPath}`);
}

async function verifyOutputWorkbook() {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(sampleOutputPath);

  const ws1 = wb.getWorksheet('Sheet1');

  // Verify String Cleaning
  assert.strictEqual(ws1.getCell('A1').value, 'Full Name');
  assert.strictEqual(ws1.getCell('B1').value, 'Department');
  assert.strictEqual(ws1.getCell('A2').value, 'Shubham Siddhapura');
  assert.strictEqual(ws1.getCell('B2').value, 'Engineering');
  assert.strictEqual(ws1.getCell('A3').value, 'John Smith');
  assert.strictEqual(ws1.getCell('B3').value, 'Human Resources');
  assert.strictEqual(ws1.getCell('A4').value, 'John Smith');
  assert.strictEqual(ws1.getCell('A10').value, 'Merged Section Header');
  assert.strictEqual(ws1.getCell('A6').value, 'Hidden Row Text');

  // Verify Non-String Preservation
  assert.strictEqual(ws1.getCell('A5').value, 9999);
  assert.ok(ws1.getCell('B5').value instanceof Date);
  assert.strictEqual(ws1.getCell('B5').numFmt, 'yyyy-mm-dd');

  // Verify Formatting & Styling Preservation
  assert.strictEqual(ws1.getCell('A1').font.bold, true);
  assert.strictEqual(ws1.getCell('A1').font.name, 'Arial');
  assert.strictEqual(ws1.getCell('A1').font.size, 12);
  assert.strictEqual(ws1.getCell('A1').font.color.argb, 'FFFFFFFF');
  assert.strictEqual(ws1.getCell('A1').fill.fgColor.argb, 'FF1F4E79');
  assert.strictEqual(ws1.getCell('A1').alignment.horizontal, 'center');
  assert.strictEqual(ws1.getColumn(1).width, 25);
  assert.strictEqual(ws1.getRow(1).height, 30);
  assert.strictEqual(ws1.getRow(6).hidden, true);
  assert.strictEqual(ws1.getColumn(4).hidden, true);
  assert.ok(ws1.views && ws1.views.length > 0 && ws1.views[0].state === 'frozen');

  // Verify Sheet 2 (Formulas, Comments)
  const ws2 = wb.getWorksheet('Employees');
  assert.strictEqual(ws2.getCell('A1').value, 'Employee ID');
  assert.strictEqual(ws2.getCell('A4').value, 'Total Salary');
  assert.strictEqual(ws2.getCell('B4').value.formula, 'SUM(B2:B3)');
  assert.strictEqual(ws2.getCell('B4').value.result, 110000);
  assert.strictEqual(ws2.getCell('B4').numFmt, '$#,##0.00');
  assert.strictEqual(ws2.getCell('A2').note, 'Unique employee identification string');
  assert.strictEqual(ws2.getCell('A5').value.hyperlink, 'https://company.internal');

  console.log('  ✓ Passed all end-to-end workbook preservation & string cleaning assertions!\n');
}

async function runRunner() {
  await createSampleWorkbook();
  const stats = await processWorkbook(sampleInputPath, sampleOutputPath);
  console.log('\n--- Execution Statistics ---');
  console.log(JSON.stringify(stats, null, 2));
  console.log('-----------------------------\n');
  await verifyOutputWorkbook();
  console.log('========================================');
  console.log('  ALL TESTS PASSED SUCCESSFULLY!        ');
  console.log('========================================\n');
}

runRunner().catch((err) => {
  console.error('\n❌ Test Suite Failed:', err);
  process.exit(1);
});
