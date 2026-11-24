import path from 'path';
import { processWorkbook } from './cleaner.js';

/**
 * Derives default input and output file paths based on CLI arguments.
 *
 * @param {string[]} args - Process arguments (process.argv.slice(2))
 * @returns {{ inputPath: string, outputPath: string }} Resolved paths
 */
export function resolveFilePaths(args) {
  const defaultInput = path.join(process.cwd(), 'input', 'workbook.xlsx');
  const defaultOutput = path.join(process.cwd(), 'output', 'workbook_cleaned.xlsx');

  if (!args || args.length === 0) {
    return {
      inputPath: defaultInput,
      outputPath: defaultOutput
    };
  }

  const inputPath = args[0];

  if (args.length >= 2) {
    return {
      inputPath,
      outputPath: args[1]
    };
  }

  // If output is omitted, create <input_name>_cleaned.<ext>
  const ext = path.extname(inputPath);
  const dir = path.dirname(inputPath);
  const baseName = path.basename(inputPath, ext);
  const outputPath = path.join(dir, `${baseName}_cleaned${ext || '.xlsx'}`);

  return {
    inputPath,
    outputPath
  };
}

/**
 * Displays statistics formatted cleanly according to specifications.
 *
 * @param {Object} stats - Statistics object returned by processWorkbook
 */
export function displayStatistics(stats) {
  console.log('\nStatistics\n');
  console.log(`Sheets Processed: ${stats.sheetsProcessed}`);
  console.log(`Rows Processed:   ${stats.rowsProcessed}`);
  console.log(`Cells Visited:    ${stats.cellsVisited}`);
  console.log(`String Cells:     ${stats.stringCells}`);
  console.log(`Modified Cells:   ${stats.modifiedCells}`);
  console.log(`Execution Time:   ${stats.executionTime}`);
}

/**
 * Main application entry point.
 */
async function main() {
  const args = process.argv.slice(2);
  const { inputPath, outputPath } = resolveFilePaths(args);

  try {
    const stats = await processWorkbook(inputPath, outputPath);
    displayStatistics(stats);
    process.exit(0);
  } catch (error) {
    console.error('\nError processing workbook:');
    console.error(`  ${error.message}`);
    process.exit(1);
  }
}

// Execute if run directly from CLI
if (import.meta.url === `file://${process.argv[1]}` || process.argv[1]?.endsWith('index.js')) {
  main();
}
