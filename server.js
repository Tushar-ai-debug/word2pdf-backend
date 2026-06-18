require('dotenv').config();
const express = require('express');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');
const { exec } = require('child_process');

const app = express();
const PORT = process.env.PORT || 5000;

// ── Directories ──────────────────────────────────────────────────────────────
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const CONVERTED_DIR = path.join(__dirname, 'converted');
[UPLOADS_DIR, CONVERTED_DIR].forEach(dir => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// ── Middleware ───────────────────────────────────────────────────────────────
app.use(cors({
  origin: process.env.FRONTEND_URL || '*',
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type'],
}));
app.use(express.json());

const limiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 20,
  message: { error: 'Too many requests, please try again later.' },
});
app.use('/api/', limiter);

// ── Multer ───────────────────────────────────────────────────────────────────
const ALLOWED_TYPES = {
  'docx-to-pdf':  { exts: ['.doc', '.docx'], mime: ['application/vnd.openxmlformats-officedocument.wordprocessingml.document','application/msword'] },
  'pptx-to-pdf':  { exts: ['.ppt', '.pptx'], mime: ['application/vnd.ms-powerpoint','application/vnd.openxmlformats-officedocument.presentationml.presentation'] },
  'xlsx-to-pdf':  { exts: ['.xls', '.xlsx'], mime: ['application/vnd.ms-excel','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'] },
  'jpg-to-pdf':   { exts: ['.jpg', '.jpeg', '.png', '.webp'], mime: ['image/jpeg','image/png','image/webp'] },
  'html-to-pdf':  { exts: ['.html', '.htm'], mime: ['text/html'] },
  'pdf-to-jpg':   { exts: ['.pdf'], mime: ['application/pdf'] },
  'pdf-to-word':  { exts: ['.pdf'], mime: ['application/pdf'] },
  'pdf-to-pptx':  { exts: ['.pdf'], mime: ['application/pdf'] },
  'pdf-to-xlsx':  { exts: ['.pdf'], mime: ['application/pdf'] },
  'pdf-to-pdfa':  { exts: ['.pdf'], mime: ['application/pdf'] },
};

function makeUpload(conversionType) {
  const allowed = ALLOWED_TYPES[conversionType];
  const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, UPLOADS_DIR),
    filename: (req, file, cb) => {
      const id = uuidv4();
      const ext = path.extname(file.originalname).toLowerCase();
      cb(null, `${id}${ext}`);
    },
  });
  const fileFilter = (req, file, cb) => {
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowed.exts.includes(ext) || allowed.mime.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error(`Invalid file type for this conversion. Allowed: ${allowed.exts.join(', ')}`), false);
    }
  };
  return multer({ storage, fileFilter, limits: { fileSize: 20 * 1024 * 1024 } });
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function safeDelete(filePath) {
  try { if (fs.existsSync(filePath)) fs.unlinkSync(filePath); } catch (e) {}
}
function scheduleCleanup(filePath, delayMs = 10 * 60 * 1000) {
  setTimeout(() => safeDelete(filePath), delayMs);
}

// ── LibreOffice converter ─────────────────────────────────────────────────────
const LO_BINS = [
  'libreoffice', 'soffice',
  '"C:\\Program Files\\LibreOffice\\program\\soffice.exe"',
  '"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"',
];

function libreOfficeConvert(inputPath, outputDir, targetFormat = 'pdf') {
  return new Promise((resolve, reject) => {
    let attempted = 0;
    function tryNext() {
      if (attempted >= LO_BINS.length) return reject(new Error('LibreOffice not found. Please install it on the server.'));
      const bin = LO_BINS[attempted++];
      const cmd = `${bin} --headless --convert-to ${targetFormat} --outdir "${outputDir}" "${inputPath}"`;
      exec(cmd, { timeout: 60000 }, (err) => {
        if (err) return tryNext();
        const baseName = path.basename(inputPath, path.extname(inputPath));
        const outPath = path.join(outputDir, `${baseName}.${targetFormat}`);
        if (fs.existsSync(outPath)) resolve(outPath);
        else tryNext();
      });
    }
    tryNext();
  });
}

// ── Generic conversion handler ────────────────────────────────────────────────
function conversionRoute(conversionType, handler) {
  const upload = makeUpload(conversionType);
  return [
    upload.single('file'),
    async (req, res) => {
      if (!req.file) return res.status(400).json({ error: 'No file uploaded or invalid file type.' });
      const inputPath = req.file.path;
      const fileId = path.basename(req.file.filename, path.extname(req.file.filename));
      try {
        const { outputPath, outputExt } = await handler(inputPath, fileId);
        safeDelete(inputPath);
        scheduleCleanup(outputPath);
        const stats = fs.statSync(outputPath);
        const origBase = path.basename(req.file.originalname, path.extname(req.file.originalname));
        res.json({
          success: true,
          fileId,
          originalName: req.file.originalname,
          outputName: `${origBase}.${outputExt}`,
          size: stats.size,
          downloadUrl: `/api/download/${fileId}?ext=${outputExt}`,
        });
      } catch (err) {
        safeDelete(inputPath);
        console.error(`[${conversionType}] Error:`, err.message);
        res.status(500).json({ error: err.message || 'Conversion failed.' });
      }
    }
  ];
}

// ── Routes ───────────────────────────────────────────────────────────────────

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ── CONVERT TO PDF ────────────────────────────────────────────────────────────

// DOCX → PDF (original route kept for backwards compat)
app.post('/api/upload', makeUpload('docx-to-pdf').single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded or invalid file type.' });
  const inputPath = req.file.path;
  const fileId = path.basename(req.file.filename, path.extname(req.file.filename));
  const outputPath = path.join(CONVERTED_DIR, `${fileId}.pdf`);
  try {
    const result = await libreOfficeConvert(inputPath, CONVERTED_DIR, 'pdf');
    if (result !== outputPath) {
      try { fs.renameSync(result, outputPath); } catch { fs.copyFileSync(result, outputPath); safeDelete(result); }
    }
    safeDelete(inputPath);
    scheduleCleanup(outputPath);
    const stats = fs.statSync(outputPath);
    const origBase = path.basename(req.file.originalname, path.extname(req.file.originalname));
    res.json({ success: true, fileId, originalName: req.file.originalname, pdfName: `${origBase}.pdf`, size: stats.size, downloadUrl: `/api/download/${fileId}` });
  } catch (err) {
    safeDelete(inputPath); safeDelete(outputPath);
    res.status(500).json({ error: err.message });
  }
});

// WORD → PDF (new route alias)
app.post('/api/convert/word-to-pdf', ...conversionRoute('docx-to-pdf', async (inputPath, fileId) => {
  const outDir = CONVERTED_DIR;
  const result = await libreOfficeConvert(inputPath, outDir, 'pdf');
  const outputPath = path.join(outDir, `${fileId}.pdf`);
  if (result !== outputPath) { try { fs.renameSync(result, outputPath); } catch { fs.copyFileSync(result, outputPath); safeDelete(result); } }
  return { outputPath, outputExt: 'pdf' };
}));

// PPTX → PDF
app.post('/api/convert/pptx-to-pdf', ...conversionRoute('pptx-to-pdf', async (inputPath, fileId) => {
  const outDir = CONVERTED_DIR;
  const result = await libreOfficeConvert(inputPath, outDir, 'pdf');
  const outputPath = path.join(outDir, `${fileId}.pdf`);
  if (result !== outputPath) { try { fs.renameSync(result, outputPath); } catch { fs.copyFileSync(result, outputPath); safeDelete(result); } }
  return { outputPath, outputExt: 'pdf' };
}));

// XLSX → PDF
app.post('/api/convert/xlsx-to-pdf', ...conversionRoute('xlsx-to-pdf', async (inputPath, fileId) => {
  const outDir = CONVERTED_DIR;
  const result = await libreOfficeConvert(inputPath, outDir, 'pdf');
  const outputPath = path.join(outDir, `${fileId}.pdf`);
  if (result !== outputPath) { try { fs.renameSync(result, outputPath); } catch { fs.copyFileSync(result, outputPath); safeDelete(result); } }
  return { outputPath, outputExt: 'pdf' };
}));

// JPG/PNG → PDF  (uses sharp + pdf-lib: npm install sharp pdf-lib)
app.post('/api/convert/jpg-to-pdf', ...conversionRoute('jpg-to-pdf', async (inputPath, fileId) => {
  const sharp = require('sharp');
  const { PDFDocument } = require('pdf-lib');
  const outputPath = path.join(CONVERTED_DIR, `${fileId}.pdf`);
  const { width, height, format } = await sharp(inputPath).metadata();
  const imgBuffer = await sharp(inputPath).toFormat('png').toBuffer();
  const pdfDoc = await PDFDocument.create();
  const img = await pdfDoc.embedPng(imgBuffer);
  const page = pdfDoc.addPage([width, height]);
  page.drawImage(img, { x: 0, y: 0, width, height });
  const pdfBytes = await pdfDoc.save();
  fs.writeFileSync(outputPath, pdfBytes);
  return { outputPath, outputExt: 'pdf' };
}));

// HTML → PDF  (uses puppeteer: npm install puppeteer)
app.post('/api/convert/html-to-pdf', ...conversionRoute('html-to-pdf', async (inputPath, fileId) => {
  const puppeteer = require('puppeteer');
  const outputPath = path.join(CONVERTED_DIR, `${fileId}.pdf`);
  const browser = await puppeteer.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] });
  const page = await browser.newPage();
  const htmlContent = fs.readFileSync(inputPath, 'utf8');
  await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
  await page.pdf({ path: outputPath, format: 'A4', printBackground: true });
  await browser.close();
  return { outputPath, outputExt: 'pdf' };
}));

// ── CONVERT FROM PDF ──────────────────────────────────────────────────────────

// PDF → JPG  (uses pdftoppm via poppler-utils: apt install poppler-utils)
app.post('/api/convert/pdf-to-jpg', ...conversionRoute('pdf-to-jpg', async (inputPath, fileId) => {
  const outputBase = path.join(CONVERTED_DIR, fileId);
  const outputPath = `${outputBase}.jpg`;
  await new Promise((resolve, reject) => {
    exec(`pdftoppm -jpeg -r 150 -singlefile "${inputPath}" "${outputBase}"`, { timeout: 30000 }, (err) => {
      if (err) reject(new Error('PDF to JPG failed. Make sure poppler-utils is installed (apt install poppler-utils).'));
      else resolve();
    });
  });
  // pdftoppm outputs as <base>.ppm or <base>.jpg depending on flag
  const candidates = [`${outputBase}.jpg`, `${outputBase}-1.jpg`];
  const found = candidates.find(f => fs.existsSync(f));
  if (!found) throw new Error('JPG output not found after conversion.');
  if (found !== outputPath) fs.renameSync(found, outputPath);
  return { outputPath, outputExt: 'jpg' };
}));

// PDF → WORD
app.post('/api/convert/pdf-to-word', ...conversionRoute('pdf-to-word', async (inputPath, fileId) => {
  const outDir = CONVERTED_DIR;
  const result = await libreOfficeConvert(inputPath, outDir, 'docx');
  const outputPath = path.join(outDir, `${fileId}.docx`);
  if (result !== outputPath) { try { fs.renameSync(result, outputPath); } catch { fs.copyFileSync(result, outputPath); safeDelete(result); } }
  return { outputPath, outputExt: 'docx' };
}));

// PDF → POWERPOINT
app.post('/api/convert/pdf-to-pptx', ...conversionRoute('pdf-to-pptx', async (inputPath, fileId) => {
  const outDir = CONVERTED_DIR;
  const result = await libreOfficeConvert(inputPath, outDir, 'pptx');
  const outputPath = path.join(outDir, `${fileId}.pptx`);
  if (result !== outputPath) { try { fs.renameSync(result, outputPath); } catch { fs.copyFileSync(result, outputPath); safeDelete(result); } }
  return { outputPath, outputExt: 'pptx' };
}));

// PDF → EXCEL
app.post('/api/convert/pdf-to-xlsx', ...conversionRoute('pdf-to-xlsx', async (inputPath, fileId) => {
  const outDir = CONVERTED_DIR;
  const result = await libreOfficeConvert(inputPath, outDir, 'xlsx');
  const outputPath = path.join(outDir, `${fileId}.xlsx`);
  if (result !== outputPath) { try { fs.renameSync(result, outputPath); } catch { fs.copyFileSync(result, outputPath); safeDelete(result); } }
  return { outputPath, outputExt: 'xlsx' };
}));

// PDF → PDF/A
app.post('/api/convert/pdf-to-pdfa', ...conversionRoute('pdf-to-pdfa', async (inputPath, fileId) => {
  const outDir = CONVERTED_DIR;
  const result = await libreOfficeConvert(inputPath, outDir, 'pdf:writer_pdf_Export');
  const outputPath = path.join(outDir, `${fileId}.pdf`);
  if (result !== outputPath) { try { fs.renameSync(result, outputPath); } catch { fs.copyFileSync(result, outputPath); safeDelete(result); } }
  return { outputPath, outputExt: 'pdf' };
}));

// ── Download ──────────────────────────────────────────────────────────────────
app.get('/api/download/:id', (req, res) => {
  const { id } = req.params;
  const { name, ext = 'pdf' } = req.query;
  if (!/^[a-f0-9-]{36}$/.test(id)) return res.status(400).json({ error: 'Invalid file ID.' });
  const filePath = path.join(CONVERTED_DIR, `${id}.${ext}`);
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: 'File not found or already expired.' });
  const safeName = (name ? String(name).replace(/[\/\\]/g, '').trim() : `converted.${ext}`);
  const mimeMap = { pdf: 'application/pdf', docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', jpg: 'image/jpeg' };
  res.setHeader('Content-Type', mimeMap[ext] || 'application/octet-stream');
  res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(safeName)}"`);
  res.sendFile(filePath, (err) => { if (err) console.error('Download error:', err.message); });
});

// ── Error handling ────────────────────────────────────────────────────────────
app.use((err, req, res, next) => {
  if (err.code === 'LIMIT_FILE_SIZE') return res.status(400).json({ error: 'File too large. Maximum size is 20MB.' });
  if (err.message?.includes('Invalid file type')) return res.status(400).json({ error: err.message });
  console.error('Unhandled error:', err);
  res.status(500).json({ error: 'Internal server error.' });
});

// ── Serve frontend in production ──────────────────────────────────────────────
if (process.env.NODE_ENV === 'production') {
  const frontendDistPath = path.join(__dirname, '../frontend/dist');
  app.use(express.static(frontendDistPath));
  app.get('*', (req, res, next) => {
    if (req.originalUrl.startsWith('/api/')) return next();
    res.sendFile(path.join(frontendDistPath, 'index.html'));
  });
}

app.listen(PORT, () => {
  console.log(`\n🚀 Doccify backend running on http://localhost:${PORT}`);
  console.log(`📁 Uploads: ${UPLOADS_DIR}`);
  console.log(`📄 Converted: ${CONVERTED_DIR}\n`);
  console.log('Supported conversions:');
  console.log('  TO PDF:   DOCX, PPTX, XLSX, JPG/PNG, HTML');
  console.log('  FROM PDF: JPG, WORD, PPTX, XLSX, PDF/A\n');
});
