const express = require('express');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');
const PDFDocument = require('pdfkit');
const QRCode = require('qrcode');
const crypto = require('crypto');
const sharp = require('sharp');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static('public'));

const DATA_FILE   = path.join(__dirname, 'data', 'estudiantes.json');
const CONFIG_FILE = path.join(__dirname, 'data', 'config.json');
const UPLOADS_DIR = path.join(__dirname, 'uploads');
const TEMP_DIR    = path.join(__dirname, 'temp');

[path.join(__dirname, 'data'), UPLOADS_DIR, TEMP_DIR].forEach(d => {
  if (!fs.existsSync(d)) fs.mkdirSync(d, { recursive: true });
});

if (!fs.existsSync(DATA_FILE))
  fs.writeFileSync(DATA_FILE, '[]');

if (!fs.existsSync(CONFIG_FILE))
  fs.writeFileSync(CONFIG_FILE, JSON.stringify({
    empresa: process.env.EMPRESA || 'SECAP - Servicio de Capacitaciones',
    sence:   process.env.SENCE   || '',
    certhia: process.env.CERTHIA || '',
    firmante:process.env.FIRMANTE|| '',
    cargo:   process.env.CARGO   || '',
    baseUrl: process.env.BASE_URL || `http://localhost:${PORT}`,
    adminPassword: process.env.ADMIN_PASSWORD || 'admin123'
  }, null, 2));

// ---------- helpers ----------

const loadStudents = () => JSON.parse(fs.readFileSync(DATA_FILE, 'utf-8'));
const loadConfig   = () => {
  const cfg = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf-8'));
  // Las variables de entorno siempre tienen prioridad (persisten entre redeploys)
  if (process.env.BASE_URL)       cfg.baseUrl       = process.env.BASE_URL;
  if (process.env.ADMIN_PASSWORD) cfg.adminPassword = process.env.ADMIN_PASSWORD;
  if (process.env.EMPRESA)        cfg.empresa       = process.env.EMPRESA;
  if (process.env.SENCE)          cfg.sence         = process.env.SENCE;
  if (process.env.CERTHIA)        cfg.certhia       = process.env.CERTHIA;
  if (process.env.FIRMANTE)       cfg.firmante      = process.env.FIRMANTE;
  if (process.env.CARGO)          cfg.cargo         = process.env.CARGO;
  return cfg;
};
const saveStudents = d  => fs.writeFileSync(DATA_FILE,   JSON.stringify(d, null, 2));
const saveConfig   = c  => fs.writeFileSync(CONFIG_FILE, JSON.stringify(c, null, 2));

function normalizeRut(rut) {
  return String(rut).replace(/[.\-\s]/g, '').toUpperCase().trim();
}

function generateCode(rut, curso) {
  return crypto.createHash('md5')
    .update(`${normalizeRut(rut)}-${curso}-${Date.now()}-${Math.random()}`)
    .digest('hex').substring(0, 12).toUpperCase();
}

function requireAdmin(req, res, next) {
  const config = loadConfig();
  if (req.headers.password !== config.adminPassword)
    return res.status(401).json({ error: 'Contraseña incorrecta' });
  next();
}

function findImage(name) {
  for (const ext of ['png', 'jpg', 'jpeg', 'webp']) {
    const p = path.join(UPLOADS_DIR, `${name}.${ext}`);
    if (fs.existsSync(p)) return p;
  }
  return null;
}

async function imageBuffer(imgPath) {
  if (imgPath.endsWith('.webp')) {
    return await sharp(imgPath).png().toBuffer();
  }
  return fs.readFileSync(imgPath);
}

// ---------- public routes ----------

app.get('/api/info', (req, res) => {
  const { empresa } = loadConfig();
  res.json({ empresa });
});

app.post('/api/buscar', (req, res) => {
  const { rut } = req.body;
  if (!rut || !rut.trim())
    return res.status(400).json({ error: 'Ingresa tu RUT' });

  const rutNorm = normalizeRut(rut);
  if (rutNorm.length < 7)
    return res.status(400).json({ error: 'RUT inválido' });

  const found = loadStudents().filter(s => normalizeRut(s.rut) === rutNorm);
  if (!found.length)
    return res.status(404).json({ error: 'No se encontró ningún certificado para este RUT.' });

  res.json(found.map(({ rut, nombre, curso, fecha, horas, codigo }) =>
    ({ rut, nombre, curso, fecha, horas, codigo })));
});

app.get('/certificado/:codigo', async (req, res) => {
  const estudiante = loadStudents().find(s => s.codigo === req.params.codigo);
  if (!estudiante) return res.status(404).send('Certificado no encontrado');

  try {
    const pdf = await generatePDF(estudiante, loadConfig());
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition',
      `attachment; filename="certificado-${normalizeRut(estudiante.rut)}.pdf"`);
    res.end(pdf);
  } catch (err) {
    console.error('PDF error:', err);
    res.status(500).send('Error al generar el certificado');
  }
});

app.get('/api/verificar/:codigo', (req, res) => {
  const e = loadStudents().find(s => s.codigo === req.params.codigo);
  if (!e) return res.json({ valido: false, mensaje: 'Certificado no encontrado en el sistema' });
  res.json({ valido: true, nombre: e.nombre, rut: e.rut, curso: e.curso, fecha: e.fecha, horas: e.horas });
});

app.get('/verificar/:codigo', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'verificar.html'));
});

app.get('/admin', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'admin.html'));
});

// ---------- admin routes ----------

app.post('/admin/login', (req, res) => {
  const config = loadConfig();
  if (req.body.password === config.adminPassword)
    res.json({ ok: true });
  else
    res.status(401).json({ error: 'Contraseña incorrecta' });
});

const uploadTemp = multer({ dest: TEMP_DIR });
const uploadImg  = multer({
  dest: UPLOADS_DIR,
  limits: { fileSize: 5 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (file.mimetype.startsWith('image/')) cb(null, true);
    else cb(new Error('Solo imágenes PNG o JPG'));
  }
});

app.post('/admin/subir', (req, res) => {
  const config = loadConfig();
  if (req.headers.password !== config.adminPassword)
    return res.status(401).json({ error: 'No autorizado' });

  uploadTemp.single('archivo')(req, res, (err) => {
    if (err) return res.status(400).json({ error: err.message });
    if (!req.file) return res.status(400).json({ error: 'Archivo requerido' });

    try {
      const wb   = XLSX.readFile(req.file.path, { cellDates: true });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, dateNF: 'dd/mm/yyyy' });
      fs.unlinkSync(req.file.path);

      if (rows.length < 2)
        return res.status(400).json({ error: 'El archivo está vacío o no tiene datos' });

      const dataRows = rows.slice(1).filter(r => r[0] && r[1] && r[2]);
      const students = loadStudents();
      let added = 0, updated = 0, errors = [];

      for (let i = 0; i < dataRows.length; i++) {
        const row    = dataRows[i];
        const rut    = String(row[0] || '').trim();
        const nombre = String(row[1] || '').trim();
        const curso  = String(row[2] || '').trim();
        const fecha  = String(row[3] || '').trim();
        const horas  = String(row[4] || '').trim();

        if (!rut || !nombre || !curso) {
          errors.push(`Fila ${i + 2}: datos incompletos (RUT, Nombre y Curso son obligatorios)`);
          continue;
        }

        const rutNorm = normalizeRut(rut);
        const idx = students.findIndex(s => normalizeRut(s.rut) === rutNorm && s.curso === curso);

        if (idx >= 0) {
          students[idx] = { ...students[idx], nombre, fecha, horas };
          updated++;
        } else {
          students.push({ rut, nombre, curso, fecha, horas, codigo: generateCode(rut, curso) });
          added++;
        }
      }

      saveStudents(students);
      res.json({ ok: true, added, updated, total: students.length, errors });

    } catch (e) {
      if (req.file && fs.existsSync(req.file.path)) fs.unlinkSync(req.file.path);
      res.status(500).json({ error: 'Error procesando archivo: ' + e.message });
    }
  });
});

app.get('/admin/estudiantes', requireAdmin, (req, res) => {
  res.json(loadStudents());
});

app.delete('/admin/estudiantes/:codigo', requireAdmin, (req, res) => {
  saveStudents(loadStudents().filter(s => s.codigo !== req.params.codigo));
  res.json({ ok: true });
});

app.get('/admin/config', requireAdmin, (req, res) => {
  const { adminPassword, ...cfg } = loadConfig();
  res.json(cfg);
});

app.post('/admin/config', requireAdmin, (req, res) => {
  const current = loadConfig();
  const { adminPassword, ...updates } = req.body;
  saveConfig({ ...current, ...updates });
  res.json({ ok: true });
});

app.post('/admin/password', requireAdmin, (req, res) => {
  const { nuevaPassword } = req.body;
  if (!nuevaPassword || nuevaPassword.length < 6)
    return res.status(400).json({ error: 'La contraseña debe tener al menos 6 caracteres' });
  const config = loadConfig();
  config.adminPassword = nuevaPassword;
  saveConfig(config);
  res.json({ ok: true });
});

app.post('/admin/imagen/:tipo', (req, res) => {
  const config = loadConfig();
  if (req.headers.password !== config.adminPassword)
    return res.status(401).json({ error: 'No autorizado' });

  const { tipo } = req.params;
  if (!['logo', 'firma', 'logo_mandante', 'logo_secap'].includes(tipo))
    return res.status(400).json({ error: 'Tipo inválido' });

  uploadImg.single('imagen')(req, res, (err) => {
    if (err) return res.status(400).json({ error: err.message });
    if (!req.file) return res.status(400).json({ error: 'Imagen requerida' });

    const ext  = (path.extname(req.file.originalname).toLowerCase() || '.png');
    const dest = path.join(UPLOADS_DIR, `${tipo}${ext}`);

    ['png', 'jpg', 'jpeg', 'webp'].forEach(e => {
      const old = path.join(UPLOADS_DIR, `${tipo}.${e}`);
      if (fs.existsSync(old)) fs.unlinkSync(old);
    });

    fs.renameSync(req.file.path, dest);
    res.json({ ok: true, url: `/uploads/${tipo}${ext}?t=${Date.now()}` });
  });
});

app.use('/uploads', express.static(UPLOADS_DIR));

app.get('/admin/plantilla', requireAdmin, (req, res) => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([
    ['RUT', 'Nombre Completo', 'Nombre del Curso', 'Fecha (DD/MM/AAAA)', 'Horas'],
    ['12.345.678-9', 'Juan Pérez González', 'Seguridad en el Trabajo', '15/03/2025', '16'],
    ['98765432-1',   'María López Soto',    'Primeros Auxilios',       '20/04/2025', '8'],
  ]);
  ws['!cols'] = [{ wch: 16 }, { wch: 30 }, { wch: 35 }, { wch: 18 }, { wch: 8 }];
  XLSX.utils.book_append_sheet(wb, ws, 'Estudiantes');
  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', 'attachment; filename="plantilla-estudiantes.xlsx"');
  res.end(buf);
});

app.get('/admin/qr-acceso', requireAdmin, async (req, res) => {
  const { baseUrl } = loadConfig();
  const url = (baseUrl || `http://localhost:${PORT}`).replace(/\/$/, '');
  const png = await QRCode.toBuffer(url, { width: 300, margin: 2 });
  res.setHeader('Content-Type', 'image/png');
  res.end(png);
});

// ---------- PDF generation ----------

async function generatePDF(estudiante, config) {
  return new Promise(async (resolve, reject) => {
    try {
      const doc = new PDFDocument({ layout: 'landscape', size: 'A4',
        margins: { top: 0, bottom: 0, left: 0, right: 0 } });

      const chunks = [];
      doc.on('data', c => chunks.push(c));
      doc.on('end',  () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);

      const W      = doc.page.width;
      const H      = doc.page.height;
      const BLUE   = '#1A5DA5';   // SECAP azul oscuro
      const CELEST = '#3DAAD8';   // SECAP celeste
      const GOLD   = '#C4973C';   // dorado para bordes decorativos del certificado
      const BG     = '#F0F8FF';   // fondo celeste muy claro

      // Background
      doc.rect(0, 0, W, H).fill(BG);

      // Outer border azul
      doc.rect(12, 12, W - 24, H - 24).lineWidth(3).stroke(BLUE);
      // Inner border celeste
      doc.rect(19, 19, W - 38, H - 38).lineWidth(1).stroke(CELEST);
      // Thin gold inner
      doc.rect(23, 23, W - 46, H - 46).lineWidth(0.5).stroke(GOLD);

      // Header bar degradado simulado (dos rectángulos)
      doc.rect(12, 12, W - 24, 100).fill(BLUE);
      doc.rect(12, 80, W - 24, 32).fill('#1562B5');

      // Logo SECAP (izquierda del header)
      const logoPath = findImage('logo');
      if (logoPath) {
        const logoBuf = await imageBuffer(logoPath);
        doc.image(logoBuf, 28, 16, { fit: [170, 90] });
      }

      // Logo del organismo mandante (derecha del header, opcional)
      const logoMandantePath = findImage('logo_mandante');
      if (logoMandantePath) {
        const logoMandanteBuf = await imageBuffer(logoMandantePath);
        doc.image(logoMandanteBuf, W - 168, 14, { fit: [140, 78] });
      }

      // Company in header
      doc.fillColor('#FFFFFF').font('Helvetica-Bold').fontSize(17)
         .text(config.empresa || 'SECAP', 210, 26, { width: W - 420, align: 'center' });

      let hy = 52;
      if (config.sence) {
        doc.font('Helvetica').fontSize(9).fillColor('#A8D4EF')
           .text(`Código SENCE: ${config.sence}`, 210, hy, { width: W - 420, align: 'center' });
        hy += 14;
      }
      if (config.certhia) {
        doc.font('Helvetica').fontSize(9).fillColor('#A8D4EF')
           .text(`Resolución CERTHIA: ${config.certhia}`, 210, hy, { width: W - 420, align: 'center' });
      }

      // Title
      doc.fillColor(BLUE).font('Helvetica-Bold').fontSize(34)
         .text('C E R T I F I C A D O', 0, 128, { align: 'center' });
      doc.fillColor(CELEST).font('Helvetica').fontSize(12)
         .text('DE APROBACIÓN DE CURSO', 0, 168, { align: 'center' });

      // Dividers
      doc.moveTo(80, 188).lineTo(W - 80, 188).lineWidth(0.7).stroke(CELEST);

      // Body
      doc.fillColor('#555').font('Helvetica').fontSize(12)
         .text('Se certifica que el/la participante:', 0, 202, { align: 'center' });

      const nameFontSize = estudiante.nombre.length > 35 ? 20 : 26;
      doc.fillColor(BLUE).font('Helvetica-Bold').fontSize(nameFontSize)
         .text(estudiante.nombre.toUpperCase(), 60, 222, { width: W - 120, align: 'center' });

      doc.fillColor('#666').font('Helvetica').fontSize(11)
         .text(`R.U.T.: ${estudiante.rut}`, 0, 258, { align: 'center' });

      doc.fillColor('#555').fontSize(12)
         .text('ha aprobado satisfactoriamente el curso de capacitación:', 0, 278, { align: 'center' });

      const courseFontSize = estudiante.curso.length > 50 ? 13 : 16;
      doc.fillColor(BLUE).font('Helvetica-Bold').fontSize(courseFontSize)
         .text(`"${estudiante.curso}"`, 60, 298, { width: W - 120, align: 'center' });

      const parts = [];
      if (estudiante.horas) parts.push(`con una duración de ${estudiante.horas} horas cronológicas`);
      if (estudiante.fecha) parts.push(`fecha ${estudiante.fecha}`);
      if (parts.length)
        doc.fillColor('#666').font('Helvetica').fontSize(11)
           .text(parts.join(' · '), 60, 328, { width: W - 120, align: 'center' });

      // Bottom divider
      doc.moveTo(80, 356).lineTo(W - 80, 356).lineWidth(0.7).stroke(CELEST);

      // Signature block (left-center)
      const sigCX = W / 2 - 140;
      const firmaPath = findImage('firma');
      if (firmaPath) {
        const firmaBuf = await imageBuffer(firmaPath);
        doc.image(firmaBuf, sigCX - 70, 364, { fit: [140, 48] });
      }

      doc.moveTo(sigCX - 90, 418).lineTo(sigCX + 90, 418).lineWidth(0.7).stroke('#333');
      doc.fillColor('#222').font('Helvetica-Bold').fontSize(11)
         .text(config.firmante || 'Director Ejecutivo', sigCX - 90, 423, { width: 180, align: 'center' });
      doc.fillColor('#666').font('Helvetica').fontSize(9)
         .text(config.cargo || '', sigCX - 90, 438, { width: 180, align: 'center' });

      // QR code (bottom right)
      const baseUrl = (config.baseUrl || `http://localhost:${PORT}`).replace(/\/$/, '');
      const verifyUrl = `${baseUrl}/verificar/${estudiante.codigo}`;
      const qrBuf = await QRCode.toBuffer(verifyUrl, {
        width: 90, margin: 1, color: { dark: '#1A5DA5', light: '#F0F8FF' }
      });
      doc.image(qrBuf, W - 143, 362, { width: 78 });
      doc.fillColor('#999').fontSize(7)
         .text('Verifique la autenticidad', W - 146, 446, { width: 84, align: 'center' });
      doc.fillColor('#BBB').fontSize(6)
         .text(`Cód: ${estudiante.codigo}`, W - 146, 457, { width: 84, align: 'center' });

      // Logo secundario SECAP (centro inferior, ej: persona abrazando mascota)
      const logoSecapPath = findImage('logo_secap');
      if (logoSecapPath) {
        const logoSecapBuf = await imageBuffer(logoSecapPath);
        const lw = 110;
        doc.image(logoSecapBuf, (W - lw) / 2, 466, { fit: [lw, 65] });
      }

      doc.end();
    } catch (err) {
      reject(err);
    }
  });
}

app.listen(PORT, () => {
  console.log(`\n✓ Sistema de Certificados iniciado`);
  console.log(`  Estudiantes : http://localhost:${PORT}/`);
  console.log(`  Admin       : http://localhost:${PORT}/admin`);
  console.log(`  Contraseña  : admin123  (cámbiala en el panel)\n`);
});
