import Fastify from 'fastify';
import cors from '@fastify/cors';
import multipart from '@fastify/multipart';
import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import JSZip from 'jszip';
import { processWorkbookBuffer } from '../cleaner.js';

// Safe output directory for local and Vercel serverless environments
const OUTPUT_DIR = process.env.VERCEL
  ? path.join('/tmp', 'output')
  : path.join(process.cwd(), 'output');

try {
  if (!fs.existsSync(OUTPUT_DIR)) {
    fs.mkdirSync(OUTPUT_DIR, { recursive: true });
  }
} catch (e) {
  console.warn('Could not ensure OUTPUT_DIR:', e.message);
}

// In-memory cache for cleaned workbooks (2 hour TTL)
const cleanedFilesCache = new Map();
const CACHE_TTL_MS = 2 * 60 * 60 * 1000;

const cacheCleanupInterval = setInterval(() => {
  const now = Date.now();
  for (const [id, item] of cleanedFilesCache.entries()) {
    if (now - item.createdAt > CACHE_TTL_MS) {
      cleanedFilesCache.delete(id);
    }
  }
}, 15 * 60 * 1000);
cacheCleanupInterval.unref();

export async function buildFastify(options = {}) {
  const app = Fastify({
    logger: process.env.NODE_ENV === 'development' ? { level: 'info' } : false,
    bodyLimit: 100 * 1024 * 1024, // 100MB
    ...options
  });

  // Enable CORS
  await app.register(cors, {
    origin: true,
    methods: ['GET', 'POST', 'OPTIONS'],
    allowedHeaders: ['Content-Type', 'X-Filename', 'X-Options']
  });

  // Enable Multipart uploads
  await app.register(multipart, {
    limits: {
      fileSize: 100 * 1024 * 1024
    }
  });

  // Accept raw binary stream as buffer for direct uploads
  app.addContentTypeParser('application/octet-stream', { parseAs: 'buffer' }, (req, body, done) => {
    done(null, body);
  });
  app.addContentTypeParser('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', { parseAs: 'buffer' }, (req, body, done) => {
    done(null, body);
  });

  // Route: Health check
  app.get('/api/health', async (req, reply) => {
    return {
      status: 'ok',
      engine: 'Fastify',
      environment: process.env.VERCEL ? 'vercel-serverless' : 'standalone',
      uptime: process.uptime()
    };
  });

  // Route: Clean Excel File
  app.post('/api/clean', async (req, reply) => {
    let fileBuffer = null;
    let originalFilename = 'workbook.xlsx';
    let options = {};

    // 1. Parse optional options from header
    const rawOptionsHeader = req.headers['x-options'];
    if (rawOptionsHeader) {
      try {
        options = JSON.parse(decodeURIComponent(rawOptionsHeader));
      } catch (e) {
        console.warn('Could not parse X-Options header:', e.message);
      }
    }

    // 2. Extract file from multipart or raw octet-stream
    if (req.isMultipart()) {
      const part = await req.file();
      if (!part) {
        reply.code(400);
        return { success: false, error: 'No file part found in multipart request' };
      }
      fileBuffer = await part.toBuffer();
      if (part.filename) {
        originalFilename = part.filename;
      }
    } else if (Buffer.isBuffer(req.body)) {
      fileBuffer = req.body;
      const headerName = req.headers['x-filename'];
      if (headerName) {
        try {
          originalFilename = decodeURIComponent(headerName);
        } catch {
          originalFilename = headerName;
        }
      }
    } else {
      reply.code(400);
      return { success: false, error: 'Invalid payload. Send raw binary buffer or multipart/form-data.' };
    }

    if (!fileBuffer || fileBuffer.length === 0) {
      reply.code(400);
      return { success: false, error: 'Uploaded file is empty' };
    }

    const ext = path.extname(originalFilename) || '.xlsx';
    const base = path.basename(originalFilename, ext);
    const cleanedFilename = `${base}_cleaned${ext}`;

    try {
      const result = await processWorkbookBuffer(fileBuffer, options);
      const downloadId = crypto.randomUUID();

      cleanedFilesCache.set(downloadId, {
        buffer: result.buffer,
        filename: cleanedFilename,
        createdAt: Date.now(),
        stats: result.stats
      });

      // Attempt to save persistent copy
      try {
        if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });
        fs.writeFileSync(path.join(OUTPUT_DIR, cleanedFilename), result.buffer);
      } catch (err) {
        // Safe to ignore in read-only environments
      }

      return {
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
      };
    } catch (err) {
      reply.code(500);
      return {
        success: false,
        error: err.message || 'Failed to process Excel workbook'
      };
    }
  });

  // Route: Download Single Cleaned File
  app.get('/api/download/:id', async (req, reply) => {
    const { id } = req.params;
    const cached = cleanedFilesCache.get(id);

    if (cached) {
      reply.header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      reply.header('Content-Disposition', `attachment; filename="${encodeURIComponent(cached.filename)}"`);
      reply.header('Content-Length', cached.buffer.length);
      return reply.send(cached.buffer);
    }

    // Disk fallback
    const safeName = path.basename(id);
    const fallbackPath = path.join(OUTPUT_DIR, safeName);
    if (fs.existsSync(fallbackPath)) {
      const fileData = fs.readFileSync(fallbackPath);
      reply.header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      reply.header('Content-Disposition', `attachment; filename="${encodeURIComponent(safeName)}"`);
      reply.header('Content-Length', fileData.length);
      return reply.send(fileData);
    }

    reply.code(404);
    return { error: 'File not found or expired. Please clean the workbook again.' };
  });

  // Route: Download All as ZIP archive
  const handleZipDownload = async (req, reply) => {
    let ids = [];
    if (req.query && req.query.ids) {
      ids = req.query.ids.split(',').map((s) => s.trim()).filter(Boolean);
    } else if (req.body && Array.isArray(req.body.ids)) {
      ids = req.body.ids;
    }

    if (ids.length === 0) {
      reply.code(400);
      return { error: 'No download IDs provided' };
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
        reply.code(404);
        return { error: 'None of the requested files were found' };
      }

      const zipBuffer = await zip.generateAsync({
        type: 'nodebuffer',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
      });

      const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const zipFilename = `cleaned_excel_workbooks_${timestamp}.zip`;

      reply.header('Content-Type', 'application/zip');
      reply.header('Content-Disposition', `attachment; filename="${zipFilename}"`);
      reply.header('Content-Length', zipBuffer.length);
      return reply.send(zipBuffer);
    } catch (zipErr) {
      reply.code(500);
      return { error: 'Failed to create zip archive: ' + zipErr.message };
    }
  };

  app.get('/api/download-zip', handleZipDownload);
  app.post('/api/download-zip', handleZipDownload);

  // Route: History of output files
  app.get('/api/history', async (req, reply) => {
    try {
      if (!fs.existsSync(OUTPUT_DIR)) {
        return { success: true, files: [] };
      }
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
        .slice(0, 15);

      return { success: true, files };
    } catch (err) {
      return { success: true, files: [] };
    }
  });

  return app;
}

export default buildFastify;
