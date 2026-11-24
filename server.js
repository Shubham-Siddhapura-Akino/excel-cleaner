import http from 'http';
import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { fileURLToPath } from 'url';
import JSZip from 'jszip';
import { processWorkbookBuffer } from './cleaner.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = parseInt(process.env.PORT || '3000', 10);
const PUBLIC_DIR = path.join(__dirname, 'public');
const OUTPUT_DIR = path.join(__dirname, 'output');

// Ensure output directory exists
if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

// In-memory cache for cleaned workbooks (cleaned after 2 hours)
// Map<string, { buffer: Buffer, filename: string, createdAt: number, stats: Object }>
const cleanedFilesCache = new Map();

const CACHE_TTL_MS = 2 * 60 * 60 * 1000; // 2 hours
const cacheCleanupInterval = setInterval(() => {
  const now = Date.now();
  for (const [id, item] of cleanedFilesCache.entries()) {
    if (now - item.createdAt > CACHE_TTL_MS) {
      cleanedFilesCache.delete(id);
    }
  }
}, 15 * 60 * 1000);
cacheCleanupInterval.unref();

/**
 * Parses multipart boundary from Content-Type header.
 */
function getBoundary(contentType) {
  const match = contentType?.match(/boundary=(?:"([^"]+)"|([^;]+))/i);
  return match ? (match[1] || match[2]).trim() : null;
}

/**
 * Extracts file buffer and filename from multipart/form-data payload.
 */
function parseMultipartFile(bodyBuffer, boundary) {
  const boundaryBuffer = Buffer.from(`--${boundary}`);
  let start = bodyBuffer.indexOf(boundaryBuffer);

  while (start !== -1) {
    const nextStart = bodyBuffer.indexOf(boundaryBuffer, start + boundaryBuffer.length);
    const part = bodyBuffer.subarray(start + boundaryBuffer.length, nextStart === -1 ? undefined : nextStart);

    const headerEnd = part.indexOf('\r\n\r\n');
    if (headerEnd !== -1) {
      const headerStr = part.subarray(0, headerEnd).toString('utf8');
      const content = part.subarray(headerEnd + 4, part.length - 2); // strip trailing \r\n

      const filenameMatch = headerStr.match(/filename="([^"]+)"/i);
      if (filenameMatch) {
        return {
          filename: filenameMatch[1],
          buffer: content
        };
      }
    }

    start = nextStart;
  }
  return null;
}

/**
 * Serves static files from public directory.
 */
function serveStaticFile(res, filePath, contentType = 'text/html') {
  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Not found' }));
      return;
    }
    res.writeHead(200, {
      'Content-Type': contentType,
      'Cache-Control': 'no-cache'
    });
    res.end(data);
  });
}

/**
 * HTTP Request Handler.
 */
const server = http.createServer(async (req, res) => {
  const parsedUrl = new URL(req.url, `http://${req.headers.host || 'localhost'}`);
  const pathname = parsedUrl.pathname;
  const method = req.method.toUpperCase();

  // CORS headers for local developer flexibility
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, X-Filename, X-Options');

  if (method === 'OPTIONS') {
    res.writeHead(204);
    res.end();
    return;
  }

  // Route: Static UI
  if (method === 'GET' && (pathname === '/' || pathname === '/index.html')) {
    const indexPath = path.join(PUBLIC_DIR, 'index.html');
    serveStaticFile(res, indexPath, 'text/html; charset=utf-8');
    return;
  }

  // Route: Health check
  if (method === 'GET' && pathname === '/api/health') {
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ status: 'ok', uptime: process.uptime() }));
    return;
  }

  // Route: Clean Excel File
  if (method === 'POST' && pathname === '/api/clean') {
    const chunks = [];
    let totalLength = 0;
    const MAX_FILE_SIZE = 100 * 1024 * 1024; // 100MB limit

    req.on('data', (chunk) => {
      totalLength += chunk.length;
      if (totalLength > MAX_FILE_SIZE) {
        res.writeHead(413, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'File size exceeds 100MB limit' }));
        req.destroy();
        return;
      }
      chunks.push(chunk);
    });

    req.on('end', async () => {
      if (res.writableEnded) return;

      try {
        const fullBody = Buffer.concat(chunks);
        const contentType = req.headers['content-type'] || '';

        let fileBuffer = null;
        let originalFilename = 'workbook.xlsx';
        let options = {};

        // Parse optional options header
        const rawOptionsHeader = req.headers['x-options'];
        if (rawOptionsHeader) {
          try {
            options = JSON.parse(decodeURIComponent(rawOptionsHeader));
          } catch (e) {
            console.warn('Could not parse X-Options header:', e.message);
          }
        }

        // Check if multipart or raw stream
        const boundary = getBoundary(contentType);
        if (boundary) {
          const parsed = parseMultipartFile(fullBody, boundary);
          if (!parsed || !parsed.buffer || parsed.buffer.length === 0) {
            res.writeHead(400, { 'Content-Type': 'application/json' });
            res.end(JSON.stringify({ error: 'No Excel file found in multipart upload' }));
            return;
          }
          fileBuffer = parsed.buffer;
          if (parsed.filename) {
            originalFilename = parsed.filename;
          }
        } else {
          // Direct binary stream
          fileBuffer = fullBody;
          const headerName = req.headers['x-filename'];
          if (headerName) {
            try {
              originalFilename = decodeURIComponent(headerName);
            } catch {
              originalFilename = headerName;
            }
          }
        }

        if (!fileBuffer || fileBuffer.length === 0) {
          res.writeHead(400, { 'Content-Type': 'application/json' });
          res.end(JSON.stringify({ error: 'Received empty file payload' }));
          return;
        }

        // Derive cleaned output filename
        const ext = path.extname(originalFilename) || '.xlsx';
        const base = path.basename(originalFilename, ext);
        const cleanedFilename = `${base}_cleaned${ext}`;

        // Process Workbook
        console.log(`[API] Cleaning workbook: "${originalFilename}" (${fileBuffer.length} bytes)...`);
        const result = await processWorkbookBuffer(fileBuffer, options);

        // Generate download token
        const downloadId = crypto.randomUUID();

        // Save in memory cache
        cleanedFilesCache.set(downloadId, {
          buffer: result.buffer,
          filename: cleanedFilename,
          createdAt: Date.now(),
          stats: result.stats
        });

        // Also save to output folder for convenience and persistence
        const persistentPath = path.join(OUTPUT_DIR, cleanedFilename);
        try {
          fs.writeFileSync(persistentPath, result.buffer);
          console.log(`[API] Saved copy to: ${persistentPath}`);
        } catch (saveErr) {
          console.warn('[API] Could not save copy to output dir:', saveErr.message);
        }

        res.writeHead(200, { 'Content-Type': 'application/json' });
        res.end(
          JSON.stringify({
            success: true,
            downloadId,
            downloadUrl: `/api/download/${downloadId}`,
            originalFilename,
            cleanedFilename,
            originalSize: fileBuffer.length,
            cleanedSize: result.buffer.length,
            stats: result.stats,
            sheetDetails: result.sheetDetails,
            sampleModifications: result.sampleModifications
          })
        );
      } catch (err) {
        console.error('[API] Error cleaning workbook:', err);
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(
          JSON.stringify({
            success: false,
            error: err.message || 'Failed to process Excel workbook'
          })
        );
      }
    });

    req.on('error', (err) => {
      console.error('[API] Request stream error:', err);
      if (!res.writableEnded) {
        res.writeHead(500, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'Upload stream failed' }));
      }
    });

    return;
  }

  // Route: Download Cleaned File
  if (method === 'GET' && pathname.startsWith('/api/download/')) {
    const downloadId = pathname.replace('/api/download/', '').trim();
    const cached = cleanedFilesCache.get(downloadId);

    if (cached) {
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(cached.filename)}"`,
        'Content-Length': cached.buffer.length,
        'Cache-Control': 'no-cache'
      });
      res.end(cached.buffer);
      return;
    }

    // Fallback: If not in cache, check if file exists in output/
    const safeFilename = path.basename(downloadId);
    const fallbackPath = path.join(OUTPUT_DIR, safeFilename);
    if (fs.existsSync(fallbackPath)) {
      const fileData = fs.readFileSync(fallbackPath);
      res.writeHead(200, {
        'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename="${encodeURIComponent(safeFilename)}"`,
        'Content-Length': fileData.length
      });
      res.end(fileData);
      return;
    }

    res.writeHead(404, { 'Content-Type': 'application/json' });
    res.end(JSON.stringify({ error: 'File not found or download link expired. Please clean the file again.' }));
    return;
  }

  // Route: Download All as ZIP archive
  if ((method === 'GET' || method === 'POST') && pathname === '/api/download-zip') {
    let ids = [];
    if (method === 'GET') {
      const idsParam = parsedUrl.searchParams.get('ids');
      if (idsParam) {
        ids = idsParam.split(',').map((s) => s.trim()).filter(Boolean);
      }
    } else {
      // POST body
      const chunks = [];
      for await (const chunk of req) chunks.push(chunk);
      try {
        const body = JSON.parse(Buffer.concat(chunks).toString('utf8'));
        if (Array.isArray(body.ids)) ids = body.ids;
      } catch (e) {
        console.warn('Could not parse JSON body in download-zip:', e.message);
      }
    }

    if (ids.length === 0) {
      res.writeHead(400, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'No download IDs provided' }));
      return;
    }

    try {
      const zip = new JSZip();
      let filesCount = 0;

      for (const id of ids) {
        let fileBuffer = null;
        let filename = null;

        const cached = cleanedFilesCache.get(id);
        if (cached) {
          fileBuffer = cached.buffer;
          filename = cached.filename;
        } else {
          // Check disk in output/
          const safeName = path.basename(id);
          const localPath = path.join(OUTPUT_DIR, safeName);
          if (fs.existsSync(localPath)) {
            fileBuffer = fs.readFileSync(localPath);
            filename = safeName;
          }
        }

        if (fileBuffer && filename) {
          let zipName = filename;
          let counter = 1;
          while (zip.file(zipName)) {
            const ext = path.extname(filename);
            const base = path.basename(filename, ext);
            zipName = `${base} (${counter})${ext}`;
            counter++;
          }
          zip.file(zipName, fileBuffer);
          filesCount++;
        }
      }

      if (filesCount === 0) {
        res.writeHead(404, { 'Content-Type': 'application/json' });
        res.end(JSON.stringify({ error: 'None of the requested files were found' }));
        return;
      }

      const zipBuffer = await zip.generateAsync({
        type: 'nodebuffer',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
      });

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const zipFilename = `cleaned_excel_workbooks_${timestamp}.zip`;

      res.writeHead(200, {
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename="${zipFilename}"`,
        'Content-Length': zipBuffer.length,
        'Cache-Control': 'no-cache'
      });
      res.end(zipBuffer);
    } catch (zipErr) {
      console.error('[API] Error creating zip archive:', zipErr);
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: 'Failed to create zip archive' }));
    }
    return;
  }

  // Route: History of output files
  if (method === 'GET' && pathname === '/api/history') {
    try {
      const files = fs.readdirSync(OUTPUT_DIR)
        .filter((f) => f.endsWith('.xlsx') || f.endsWith('.xlsm') || f.endsWith('.xls'))
        .map((filename) => {
          const filePath = path.join(OUTPUT_DIR, filename);
          const stat = fs.statSync(filePath);
          return {
            filename,
            size: stat.size,
            mtime: stat.mtime
          };
        })
        .sort((a, b) => b.mtime - a.mtime)
        .slice(0, 10);

      res.writeHead(200, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ success: true, files }));
    } catch (err) {
      res.writeHead(500, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({ error: err.message }));
    }
    return;
  }

  // 404 for any other route
  res.writeHead(404, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify({ error: 'Endpoint not found' }));
});

/**
 * Starts the server on an available port.
 */
function startServer(portToTry) {
  server.listen(portToTry, () => {
    console.log('\n======================================================');
    console.log(`  🚀 Excel Cleaner UI is running!`);
    console.log(`  🌐 Open in your browser: http://localhost:${portToTry}`);
    console.log('======================================================\n');
  });

  server.on('error', (err) => {
    if (err.code === 'EADDRINUSE') {
      console.warn(`Port ${portToTry} is in use, trying port ${portToTry + 1}...`);
      startServer(portToTry + 1);
    } else {
      console.error('Server error:', err);
    }
  });
}

// Start server if run directly
if (import.meta.url === `file://${process.argv[1]}` || process.argv[1]?.endsWith('server.js')) {
  startServer(PORT);
}

export { server, startServer };
