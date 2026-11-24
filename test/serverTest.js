import http from 'http';
import fs from 'fs';
import path from 'path';
import assert from 'assert';
import ExcelJS from 'exceljs';
import { buildFastify } from '../server/app.js';

const TEST_PORT = 38472;

function makeRequest(options, bodyBuffer = null) {
  return new Promise((resolve, reject) => {
    const req = http.request(options, (res) => {
      const chunks = [];
      res.on('data', (c) => chunks.push(c));
      res.on('end', () => {
        resolve({
          statusCode: res.statusCode,
          headers: res.headers,
          body: Buffer.concat(chunks)
        });
      });
    });
    req.on('error', reject);
    if (bodyBuffer) {
      req.write(bodyBuffer);
    }
    req.end();
  });
}

async function runServerTests() {
  console.log('--- Starting Fastify Backend Integration Tests ---');

  const app = await buildFastify({ logger: false });
  await app.listen({ port: TEST_PORT, host: 'localhost' });
  console.log(`Fastify test server listening on port ${TEST_PORT}`);

  try {
    // 1. Test GET /api/health
    console.log('[Test 1] Testing GET /api/health...');
    const healthRes = await makeRequest({
      hostname: 'localhost',
      port: TEST_PORT,
      path: '/api/health',
      method: 'GET'
    });
    assert.strictEqual(healthRes.statusCode, 200);
    const healthData = JSON.parse(healthRes.body.toString());
    assert.strictEqual(healthData.status, 'ok');
    assert.strictEqual(healthData.engine, 'Fastify');
    console.log('  ✓ GET /api/health returned 200 OK with Fastify engine');

    // 2. Test POST /api/clean with sample Excel workbook
    console.log('[Test 2] Testing POST /api/clean (Uploading Excel Workbook)...');
    const sampleFilePath = path.join(process.cwd(), 'input', 'test_workbook.xlsx');
    assert.ok(fs.existsSync(sampleFilePath), 'Sample file test_workbook.xlsx must exist');
    const fileBuffer = fs.readFileSync(sampleFilePath);

    const cleanRes = await makeRequest(
      {
        hostname: 'localhost',
        port: TEST_PORT,
        path: '/api/clean',
        method: 'POST',
        headers: {
          'Content-Type': 'application/octet-stream',
          'X-Filename': encodeURIComponent('test_workbook.xlsx'),
          'X-Options': encodeURIComponent(JSON.stringify({ removeQuotes: true, normalizeWhitespace: true, trimSpaces: true }))
        }
      },
      fileBuffer
    );

    assert.strictEqual(cleanRes.statusCode, 200);
    const cleanData = JSON.parse(cleanRes.body.toString());
    assert.strictEqual(cleanData.success, true);
    assert.ok(cleanData.downloadId);
    assert.strictEqual(cleanData.stats.sheetsProcessed, 2);
    assert.strictEqual(cleanData.stats.modifiedCells, 15);
    console.log(`  ✓ POST /api/clean succeeded. downloadId: ${cleanData.downloadId}, modified: ${cleanData.stats.modifiedCells}`);

    // 3. Test GET /api/download/:downloadId
    console.log('[Test 3] Testing GET /api/download/:downloadId...');
    const downloadRes = await makeRequest({
      hostname: 'localhost',
      port: TEST_PORT,
      path: `/api/download/${cleanData.downloadId}`,
      method: 'GET'
    });
    assert.strictEqual(downloadRes.statusCode, 200);
    assert.ok(downloadRes.headers['content-disposition'].includes('test_workbook_cleaned.xlsx'));

    // Verify downloaded buffer integrity with ExcelJS
    const downloadedWb = new ExcelJS.Workbook();
    await downloadedWb.xlsx.load(downloadRes.body);
    const ws1 = downloadedWb.getWorksheet('Sheet1');
    assert.strictEqual(ws1.getCell('A2').value, 'Shubham Siddhapura');
    assert.strictEqual(ws1.getCell('B2').value, 'Engineering');
    assert.strictEqual(ws1.getCell('A1').font.bold, true);
    console.log('  ✓ GET /api/download verified. Excel file is valid and cleaned properly');

    // 4. Test GET /api/history
    console.log('[Test 4] Testing GET /api/history...');
    const historyRes = await makeRequest({
      hostname: 'localhost',
      port: TEST_PORT,
      path: '/api/history',
      method: 'GET'
    });
    assert.strictEqual(historyRes.statusCode, 200);
    const historyData = JSON.parse(historyRes.body.toString());
    assert.strictEqual(historyData.success, true);
    assert.ok(Array.isArray(historyData.files));
    console.log(`  ✓ GET /api/history returned ${historyData.files.length} recent output file(s)`);

    // 5. Test Batch /api/download-zip
    console.log('[Test 5] Testing Batch ZIP Download (/api/download-zip)...');
    const cleanRes2 = await makeRequest(
      {
        hostname: 'localhost',
        port: TEST_PORT,
        path: '/api/clean',
        method: 'POST',
        headers: {
          'Content-Type': 'application/octet-stream',
          'X-Filename': encodeURIComponent('second_sheet.xlsx'),
          'X-Options': encodeURIComponent(JSON.stringify({ removeQuotes: true, normalizeWhitespace: true }))
        }
      },
      fileBuffer
    );
    const cleanData2 = JSON.parse(cleanRes2.body.toString());
    assert.strictEqual(cleanData2.success, true);

    const zipRes = await makeRequest({
      hostname: 'localhost',
      port: TEST_PORT,
      path: `/api/download-zip?ids=${cleanData.downloadId},${cleanData2.downloadId}`,
      method: 'GET'
    });

    assert.strictEqual(zipRes.statusCode, 200);
    assert.strictEqual(zipRes.headers['content-type'], 'application/zip');
    assert.ok(zipRes.headers['content-disposition'].includes('.zip'));

    const JSZip = (await import('jszip')).default;
    const loadedZip = await JSZip.loadAsync(zipRes.body);
    const zipFilenames = Object.keys(loadedZip.files);
    assert.ok(zipFilenames.length >= 2, `Expected at least 2 files in zip, got: ${zipFilenames.join(', ')}`);
    console.log(`  ✓ GET /api/download-zip verified. ZIP contains ${zipFilenames.length} files: ${zipFilenames.join(', ')}`);

    console.log('\n========================================');
    console.log('  ALL FASTIFY TESTS PASSED SUCCESSFULLY! ');
    console.log('========================================');
  } finally {
    await app.close();
  }
}

runServerTests().catch((err) => {
  console.error('\n❌ Fastify Test Failed:', err);
  process.exit(1);
});
