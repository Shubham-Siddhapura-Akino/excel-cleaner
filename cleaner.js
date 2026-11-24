import fs from 'fs';
import path from 'path';
import ExcelJS from 'exceljs';
import { cleanText } from './utils/cleanText.js';

/**
 * Processes a single worksheet, cleaning all simple string cells while preserving formatting.
 *
 * @param {ExcelJS.Worksheet} worksheet - The ExcelJS worksheet to process
 * @param {Object} [options={}] - Configurable cleaning options
 * @param {Array} [sampleModifications=null] - Optional array to collect sample before/after diffs
 * @returns {Object} Statistics for this worksheet
 */
export function processWorksheet(worksheet, options = {}, sampleModifications = null) {
  let rowsProcessed = 0;
  let cellsVisited = 0;
  let stringCells = 0;
  let modifiedCells = 0;
  let maxNonEmptyRow = 0;

  worksheet.eachRow({ includeEmpty: false }, (row) => {
    rowsProcessed++;
    if (row.number > maxNonEmptyRow) {
      maxNonEmptyRow = row.number;
    }

    row.eachCell({ includeEmpty: false }, (cell) => {
      cellsVisited++;

      // Preserve non-string cells: Formula, RichText, Hyperlink objects, Numbers, Dates, Booleans, Empty cells
      if (
        cell.type === ExcelJS.ValueType.Formula ||
        cell.type === ExcelJS.ValueType.RichText
      ) {
        return;
      }

      const rawValue = cell.value;

      // Only clean primitive string values
      if (typeof rawValue === 'string') {
        stringCells++;
        const cleaned = cleanText(rawValue, options);
        if (cleaned !== rawValue) {
          if (sampleModifications && sampleModifications.length < 50) {
            sampleModifications.push({
              sheet: worksheet.name,
              cell: cell.address,
              before: rawValue,
              after: cleaned
            });
          }
          cell.value = cleaned;
          modifiedCells++;
        }
      }
    });
  });

  // Prune trailing empty rows to prevent ExcelJS from writing millions of empty row tags
  if (options.pruneEmptyRows !== false) {
    if (worksheet._rows && worksheet._rows.length > maxNonEmptyRow + 1) {
      worksheet._rows.splice(maxNonEmptyRow + 1);
    }
  }

  return {
    rowsProcessed,
    cellsVisited,
    stringCells,
    modifiedCells
  };
}

/**
 * Processes an entire Excel workbook, reading from inputPath and writing to outputPath.
 *
 * @param {string} inputPath - Absolute or relative path to the input Excel workbook
 * @param {string} outputPath - Absolute or relative path for the cleaned Excel workbook
 * @param {Object} [options={}] - Configurable cleaning options
 * @returns {Promise<Object>} Aggregated execution statistics
 */
export async function processWorkbook(inputPath, outputPath, options = {}) {
  const startTime = Date.now();

  const resolvedInputPath = path.resolve(inputPath);
  const resolvedOutputPath = path.resolve(outputPath);

  // 1. Input Validation & File Checks
  if (!fs.existsSync(resolvedInputPath)) {
    const error = new Error(`Input file does not exist: "${resolvedInputPath}"`);
    error.code = 'ENOENT';
    throw error;
  }

  // Check read permission
  try {
    fs.accessSync(resolvedInputPath, fs.constants.R_OK);
  } catch (err) {
    const error = new Error(`Permission denied reading input file: "${resolvedInputPath}"`);
    error.code = 'EACCES';
    throw error;
  }

  // 2. Load Workbook
  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.readFile(resolvedInputPath);
  } catch (err) {
    const error = new Error(`Failed to read Excel workbook "${resolvedInputPath}": ${err.message}`);
    error.code = 'INVALID_WORKBOOK';
    throw error;
  }

  console.log('Workbook Loaded');

  let totalSheetsProcessed = 0;
  let totalRowsProcessed = 0;
  let totalCellsVisited = 0;
  let totalStringCells = 0;
  let totalModifiedCells = 0;

  // 3. Process Worksheets
  for (const worksheet of workbook.worksheets) {
    console.log(`Processing Sheet: ${worksheet.name}`);

    const stats = processWorksheet(worksheet, options);

    totalSheetsProcessed++;
    totalRowsProcessed += stats.rowsProcessed;
    totalCellsVisited += stats.cellsVisited;
    totalStringCells += stats.stringCells;
    totalModifiedCells += stats.modifiedCells;
  }

  console.log(`\nSaving cleaned workbook to "${resolvedOutputPath}"...`);

  // 4. Ensure output directory exists & write workbook
  const outputDir = path.dirname(resolvedOutputPath);
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }

  try {
    await workbook.xlsx.writeFile(resolvedOutputPath);
  } catch (err) {
    const error = new Error(`Failed to save cleaned workbook to "${resolvedOutputPath}": ${err.message}`);
    error.code = 'EWRITE';
    throw error;
  }

  console.log('Workbook saved successfully.');
  console.log('\nCleaning Complete');

  const executionTimeMs = Date.now() - startTime;
  const executionTimeStr = executionTimeMs >= 1000
    ? `${(executionTimeMs / 1000).toFixed(2)}s`
    : `${executionTimeMs}ms`;

  const finalStats = {
    sheetsProcessed: totalSheetsProcessed,
    rowsProcessed: totalRowsProcessed,
    cellsVisited: totalCellsVisited,
    stringCells: totalStringCells,
    modifiedCells: totalModifiedCells,
    executionTime: executionTimeStr,
    executionTimeMs
  };

  return finalStats;
}

/**
 * Processes an in-memory Excel workbook buffer, returning cleaned buffer and rich statistics.
 *
 * @param {Buffer} buffer - Binary buffer of the input Excel file
 * @param {Object} [options={}] - Configurable cleaning options
 * @returns {Promise<{ buffer: Buffer, stats: Object, sheetDetails: Array, sampleModifications: Array }>}
 */
export async function processWorkbookBuffer(buffer, options = {}) {
  const startTime = Date.now();

  const workbook = new ExcelJS.Workbook();
  try {
    await workbook.xlsx.load(buffer);
  } catch (err) {
    const error = new Error(`Failed to parse Excel workbook: ${err.message}`);
    error.code = 'INVALID_WORKBOOK';
    throw error;
  }

  let totalSheetsProcessed = 0;
  let totalRowsProcessed = 0;
  let totalCellsVisited = 0;
  let totalStringCells = 0;
  let totalModifiedCells = 0;
  const sheetDetails = [];
  const sampleModifications = [];

  for (const worksheet of workbook.worksheets) {
    const stats = processWorksheet(worksheet, options, sampleModifications);

    sheetDetails.push({
      name: worksheet.name,
      rows: stats.rowsProcessed,
      cells: stats.cellsVisited,
      stringCells: stats.stringCells,
      modifiedCells: stats.modifiedCells
    });

    totalSheetsProcessed++;
    totalRowsProcessed += stats.rowsProcessed;
    totalCellsVisited += stats.cellsVisited;
    totalStringCells += stats.stringCells;
    totalModifiedCells += stats.modifiedCells;
  }

  const outBuffer = await workbook.xlsx.writeBuffer();
  const executionTimeMs = Date.now() - startTime;
  const executionTimeStr = executionTimeMs >= 1000
    ? `${(executionTimeMs / 1000).toFixed(2)}s`
    : `${executionTimeMs}ms`;

  return {
    buffer: Buffer.from(outBuffer),
    stats: {
      sheetsProcessed: totalSheetsProcessed,
      rowsProcessed: totalRowsProcessed,
      cellsVisited: totalCellsVisited,
      stringCells: totalStringCells,
      modifiedCells: totalModifiedCells,
      executionTime: executionTimeStr,
      executionTimeMs
    },
    sheetDetails,
    sampleModifications
  };
}

export default {
  processWorksheet,
  processWorkbook,
  processWorkbookBuffer
};
