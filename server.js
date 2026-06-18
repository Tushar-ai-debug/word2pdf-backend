require('dotenv').config();
const express = require('express');
const cors = require('cors');
const rateLimit = require('express-rate-limit');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const { v4: uuidv4 } = require('uuid');

const app = express();
const PORT = process.env.PORT || 5000;

// ── Directories ─────────────────────────────────────────────────────────────
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const CONVERTED_DIR = path.join(__dirname, 'converted');
[UPLOADS_DIR, CONVERTED_DIR].forEach(dir => {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
});

// ── Middleware ───────────────────────────────────────────────────────────────
app.use(cors({
  origin: process.env.FRONTEND_URL || 'http://localhost:5173',
  methods: ['GET', 'POST'],
  allowedHeaders: ['Content-Type'],
}));
app.use(express.json());

// Rate limiting
const limiter = rateLimit({
  windowMs: 15 * 60 * 1000, // 15 minutes
  max: 20,
  message: { error: 'Too many requests, please try again later.' },
});
app.use('/api/', limiter);

// ── Multer config ────────────────────────────────────────────────────────────
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOADS_DIR),
  filename: (req, file, cb) => {
    const id = uuidv4();
    const ext = path.extname(file.originalname);
    cb(null, `${id}${ext}`);
  },
});

const fileFilter = (req, file, cb) => {
  const allowed = [
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    'application/msword',
  ];
  const ext = path.extname(file.originalname).toLowerCase();
  if (allowed.includes(file.mimetype) || ext === '.docx' || ext === '.doc') {
    cb(null, true);
  } else {
    cb(new Error('Only .doc and .docx files are allowed'), false);
  }
};

const upload = multer({
  storage,
  fileFilter,
  limits: { fileSize: 10 * 1024 * 1024 }, // 10 MB
});

// ── Helper: delete file safely ────────────────────────────────────────────────
function safeDelete(filePath) {
  try {
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
  } catch (e) {
    console.error('Failed to delete file:', filePath, e.message);
  }
}

// ── Helper: schedule cleanup ──────────────────────────────────────────────────
function scheduleCleanup(filePath, delayMs = 10 * 60 * 1000) {
  setTimeout(() => safeDelete(filePath), delayMs);
}

// ── Conversion function ───────────────────────────────────────────────────────
const { exec } = require('child_process');

async function convertDocxToPdf(inputPath, outputPath) {
  // Try LibreOffice CLI first (best quality, most reliable)
  const libreOfficeSuccess = await tryLibreOfficeCLI(inputPath, outputPath);
  if (libreOfficeSuccess) return outputPath;

  // Fallback: docx-pdf npm package
  return new Promise((resolve, reject) => {
    try {
      const docxPdf = require('docx-pdf');
      docxPdf(inputPath, outputPath, (err) => {
        if (err) reject(new Error('Conversion failed: ' + err.message));
        else resolve(outputPath);
      });
    } catch (e) {
      reject(new Error(
        'Conversion failed. Please make sure LibreOffice is installed.\n' +
        'Download from: https://www.libreoffice.org/download'
      ));
    }
  });
}

// Try converting using LibreOffice command-line (works on Windows, Mac, Linux)
function tryLibreOfficeCLI(inputPath, outputPath) {
  return new Promise((resolve) => {
    const outputDir = path.dirname(outputPath);

    // LibreOffice executable paths for different OS
    const librePaths = [
      'libreoffice',                                              // Linux/Mac (in PATH)
      'soffice',                                                  // Linux/Mac alternative
      '"C:\\Program Files\\LibreOffice\\program\\soffice.exe"',  // Windows default
      '"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"', // Windows 32-bit
    ];

    const inputEscaped = `"${inputPath}"`;
    let attempted = 0;

    function tryNext() {
      if (attempted >= librePaths.length) {
        resolve(false); // All paths failed
        return;
      }

      const bin = librePaths[attempted++];
      const cmd = `${bin} --headless --convert-to pdf --outdir "${outputDir}" ${inputEscaped}`;

      exec(cmd, { timeout: 30000 }, (err, stdout, stderr) => {
        if (err) {
          tryNext(); // Try next path
          return;
        }

        // LibreOffice outputs file as <inputname>.pdf in outputDir
        const generatedName = path.basename(inputPath, path.extname(inputPath)) + '.pdf';
        const generatedPath = path.join(outputDir, generatedName);

        // Rename to our expected outputPath if different
        if (generatedPath !== outputPath && fs.existsSync(generatedPath)) {
          try {
            fs.renameSync(generatedPath, outputPath);
          } catch (e) {
            fs.copyFileSync(generatedPath, outputPath);
            safeDelete(generatedPath);
          }
        }

        if (fs.existsSync(outputPath)) {
          resolve(true);
        } else {
          tryNext();
        }
      });
    }

    tryNext();
  });
}

// ── Routes ───────────────────────────────────────────────────────────────────

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Upload & convert
app.post('/api/upload', upload.single('file'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded or invalid file type.' });
  }

  const inputPath = req.file.path;
  const fileId = path.basename(req.file.filename, path.extname(req.file.filename));
  const outputPath = path.join(CONVERTED_DIR, `${fileId}.pdf`);

  try {
    await convertDocxToPdf(inputPath, outputPath);

    // Clean up source file
    safeDelete(inputPath);

    // Auto-delete PDF after 10 minutes
    scheduleCleanup(outputPath);

    const stats = fs.statSync(outputPath);

    res.json({
      success: true,
      fileId,
      originalName: req.file.originalname,
      pdfName: `${path.basename(req.file.originalname, path.extname(req.file.originalname))}.pdf`,
      size: stats.size,
      downloadUrl: `/api/download/${fileId}`,
    });
  } catch (err) {
    safeDelete(inputPath);
    safeDelete(outputPath);
    console.error('Conversion error:', err.message);
    res.status(500).json({ error: err.message || 'Conversion failed. Please try again.' });
  }
});

// Download converted PDF
app.get('/api/download/:id', (req, res) => {
  const { id } = req.params;
  const { name } = req.query;

  // Sanitize id to prevent path traversal
  if (!/^[a-f0-9-]{36}$/.test(id)) {
    return res.status(400).json({ error: 'Invalid file ID.' });
  }

  const filePath = path.join(CONVERTED_DIR, `${id}.pdf`);
  if (!fs.existsSync(filePath)) {
    return res.status(404).json({ error: 'File not found or already expired.' });
  }

  // Set the dynamic filename, safely encoding it to prevent header issues
  const fallbackName = 'converted.pdf';
  let safeName = fallbackName;
  if (name && typeof name === 'string') {
    safeName = name.replace(/[\/\\]/g, '').trim();
    if (!safeName.endsWith('.pdf')) {
      safeName += '.pdf';
    }
  }

  res.setHeader('Content-Type', 'application/pdf');
  res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(safeName)}"`);
  res.sendFile(filePath, (err) => {
    if (err) console.error('Download error:', err.message);
  });
});

// Serve static assets from the React frontend in production
if (process.env.NODE_ENV === 'production') {
  const frontendDistPath = path.join(__dirname, '../frontend/dist');
  app.use(express.static(frontendDistPath));

  app.get('*', (req, res, next) => {
    if (req.originalUrl.startsWith('/api/')) {
      return next();
    }
    res.sendFile(path.join(frontendDistPath, 'index.html'));
  });
}

// ── Error handling ────────────────────────────────────────────────────────────
app.use((err, req, res, next) => {
  if (err.code === 'LIMIT_FILE_SIZE') {
    return res.status(400).json({ error: 'File too large. Maximum size is 10MB.' });
  }
  if (err.message && err.message.includes('Only .doc')) {
    return res.status(400).json({ error: err.message });
  }
  console.error('Unhandled error:', err);
  res.status(500).json({ error: 'Internal server error.' });
});

// ── Start ─────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n🚀 Word2PDF backend running on http://localhost:${PORT}`);
  console.log(`📁 Uploads: ${UPLOADS_DIR}`);
  console.log(`📄 Converted: ${CONVERTED_DIR}\n`);
});
