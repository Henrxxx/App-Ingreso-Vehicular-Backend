import express from 'express';
import cors from 'cors';
import multer from 'multer';
import dotenv from 'dotenv';
import fs from 'fs';
import path from 'path';
import { exec } from 'child_process';
import { promisify } from 'util';
import nodemailer from 'nodemailer';
import { ImapFlow } from 'imapflow';
import MailComposer from 'nodemailer/lib/mail-composer/index.js';

dotenv.config();

const execAsync = promisify(exec);
const app = express();

const transporter = nodemailer.createTransport({
  host: 'cargushn.com',
  port: 465,
  secure: true,
  auth: {
    user: 'aservicios@cargushn.com',
    pass: process.env.SMTP_PASS
  }
});

function normalizePedidoBase(raw = '') {
  const match = String(raw).match(/\d{4}/);
  return match ? match[0] : '';
}

function uniqueEmails(list = []) {
  const seen = new Set();
  const result = [];

  for (const item of list) {
    const email = String(item || '').trim().toLowerCase();
    if (!email) continue;
    if (seen.has(email)) continue;
    seen.add(email);
    result.push(email);
  }

  return result;
}

function addressListToEmails(list = []) {
  if (!Array.isArray(list)) return [];
  return uniqueEmails(
    list.map(item => item?.address || '').filter(Boolean)
  );
}

function extractHeaderValue(sourceText, headerName) {
  const regex = new RegExp(`^${headerName}:\\s*(.+)$`, 'im');
  const match = sourceText.match(regex);
  return match ? match[1].trim() : '';
}

function extractReferences(headerValue = '') {
  return String(headerValue)
    .match(/<[^>]+>/g) || [];
}

function extractThreadMetaFromSource(sourceBuffer) {
  const sourceText = Buffer.isBuffer(sourceBuffer)
    ? sourceBuffer.toString('utf8')
    : String(sourceBuffer || '');

  const messageId = extractHeaderValue(sourceText, 'Message-ID');
  const inReplyTo = extractHeaderValue(sourceText, 'In-Reply-To');
  const referencesRaw = extractHeaderValue(sourceText, 'References');

  return {
    messageId,
    inReplyTo,
    references: extractReferences(referencesRaw)
  };
}

function createImapClient() {
  const client = new ImapFlow({
    host: process.env.IMAP_HOST,
    port: Number(process.env.IMAP_PORT || 993),
    secure: String(process.env.IMAP_SECURE || 'true').toLowerCase() === 'true',
    auth: {
      user: process.env.IMAP_USER,
      pass: process.env.IMAP_PASS
    },

    // Evita que conexiones lentas tumben el backend tan rápido
    connectionTimeout: 20000,
    greetingTimeout: 20000,
    socketTimeout: 45000
  });

  // MUY IMPORTANTE: evita que un error de IMAP tumbe todo Node
  client.on('error', err => {
    console.error('[IMAP] Error controlado:', err?.message || err);
  });

  return client;
}

async function appendToSentMailbox(mailOptions) {
  const sentMailbox = process.env.IMAP_SENT_MAILBOX || 'Sent';
  const client = createImapClient();

  await client.connect();
  try {
    const rawMessage = await new MailComposer(mailOptions).compile().build();
    await client.append(sentMailbox, rawMessage, ['\\Seen']);
  } finally {
    await client.logout();
  }
}

function escapeHtml(value = '') {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function valueOrNoDef(value) {
  const clean = String(value || '').trim();
  return clean ? escapeHtml(clean) : 'no definido';
}

function pickFirstDefined(data, keys = []) {
  for (const key of keys) {
    const value = data?.[key];
    if (String(value || '').trim()) {
      return String(value).trim();
    }
  }
  return '';
}

function hasCheck(checkedSet, number) {
  return checkedSet.has(Number(number));
}

function buildDetallesObservacionesHtml(data = {}) {
  const checkedNumbers = Array.isArray(data.checkboxNumbers) ? data.checkboxNumbers : [];
  const checkedSet = new Set(checkedNumbers.map(Number));

  const fragments = [];

  // Vehículo nuevo/usado
  if (hasCheck(checkedSet, 1)) {
    fragments.push('vehículo nuevo');
  } else if (hasCheck(checkedSet, 2)) {
    fragments.push('vehículo usado');
  } else {
    fragments.push('tipo de condición no definida');
  }

  // Documentación
  if (hasCheck(checkedSet, 18)) {
    fragments.push('trae documentación');
  } else if (hasCheck(checkedSet, 20)) {
    fragments.push('no trae documentación');
  } else {
    fragments.push('documentación no definida');
  }

  // Sistema eléctrico
  if (hasCheck(checkedSet, 39)) {
    fragments.push('sistema eléctrico funcional');
  } else if (hasCheck(checkedSet, 40)) {
    fragments.push('sistema eléctrico con deficiencias');
  } else if (hasCheck(checkedSet, 41)) {
    fragments.push('sistema eléctrico no funcional');
  } else {
    fragments.push('estado del sistema eléctrico no definido');
  }

  // Freno de mano
  if (hasCheck(checkedSet, 30)) {
    fragments.push('freno de mano en buen estado');
  } else if (hasCheck(checkedSet, 31)) {
    fragments.push('freno de mano con deficiencias');
  } else if (hasCheck(checkedSet, 32)) {
    fragments.push('<strong><u>freno de mano no funcional</u></strong>');
  } else {
    fragments.push('estado del freno de mano no definido');
  }

  // Borner de batería
  if (hasCheck(checkedSet, 57)) {
    fragments.push('borner de la batería en buen estado');
  } else if (hasCheck(checkedSet, 58)) {
    fragments.push('borner de la batería con deficiencias');
  } else if (hasCheck(checkedSet, 59)) {
    fragments.push('borner de la batería en mal estado');
  } else {
    fragments.push('estado del borner de la batería no definido');
  }

  // Herramientas
  const herramientas = [];
  herramientas.push(hasCheck(checkedSet, 8) ? 'trae jack' : 'no trae jack');
  herramientas.push(hasCheck(checkedSet, 9) ? 'trae llave de ruedas' : 'no trae llave de ruedas');
  herramientas.push(hasCheck(checkedSet, 10) ? 'trae maneral' : 'no trae maneral');
  fragments.push(...herramientas);

  // Radio
  if (hasCheck(checkedSet, 78)) {
    fragments.push('radio funcional');
  } else if (hasCheck(checkedSet, 79)) {
    fragments.push('radio con deficiencias');
  } else if (hasCheck(checkedSet, 80)) {
    fragments.push('no trae radio / radio no funcional');
  } else {
    fragments.push('estado de la radio no definido');
  }

  // Parabrisas
  if (hasCheck(checkedSet, 45)) {
    fragments.push('parabrisas en buen estado');
  } else if (hasCheck(checkedSet, 46)) {
    fragments.push('parabrisas con deficiencias');
  } else if (hasCheck(checkedSet, 47)) {
    fragments.push('parabrisas en mal estado');
  } else {
    fragments.push('estado del parabrisas no definido');
  }

  return fragments
    .map(fragment => {
      // Si ya trae HTML (caso freno de mano peligroso), respetarlo
      if (fragment.includes('<strong>') || fragment.includes('<u>')) {
        return fragment;
      }
      return escapeHtml(fragment);
    })
    .join('; ') + '.';
}

function buildNotasHtml(data = {}) {
  const notas = Array.isArray(data.notasExtra) ? data.notasExtra : [];
  const cleanNotas = notas
    .map(v => String(v || '').trim())
    .filter(Boolean);

  if (!cleanNotas.length) return '';

  const primera = cleanNotas[0];
  const restantes = cleanNotas.slice(1);

  return `
    <div style="
      margin-top:14px;
      background:#fff76a;
      padding:10px 12px;
      border:1px solid #d6cc32;
      border-radius:4px;
      font-family: Arial, sans-serif;
      font-size:14px;
      line-height:1.6;
      display:inline-block;
    ">
      <div>
        <strong><u>Nota:</u></strong> ${escapeHtml(primera)}
      </div>
      ${restantes.map(nota => `
        <div style="padding-left:44px;">
          ${escapeHtml(nota)}
        </div>
      `).join('')}
    </div>
  `;
}

async function inspectMailboxForPedido(client, mailboxName, pedidoBase) {
  const lookbackDays = Number(process.env.IMAP_LOOKBACK_DAYS || 90);
  const results = [];

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - lookbackDays);

  const lock = await client.getMailboxLock(mailboxName);
  try {
    // Buscar en el servidor IMAP solo por asunto + fecha
    const uids = await client.search({
      since: cutoffDate,
      subject: pedidoBase
    }, { uid: true });

    if (!uids.length) {
      return results;
    }

    // Validación extra: que el número esté realmente aislado en el asunto
    const pedidoRegex = new RegExp(`(?:^|[^0-9])${pedidoBase}(?:[^0-9]|$)`, 'i');

    // Solo traemos metadatos livianos primero
    const messages = await client.fetchAll(uids, {
      uid: true,
      envelope: true,
      internalDate: true
    }, { uid: true });

    for (const message of messages) {
      const subject = String(message.envelope?.subject || '').trim();
      if (!pedidoRegex.test(subject)) continue;

      // Solo si ya coincidió el asunto, traemos source de ese mensaje
      const fullMessage = await client.fetchOne(message.uid, {
        uid: true,
        source: true
      }, { uid: true });

      const meta = extractThreadMetaFromSource(fullMessage?.source);

      results.push({
        mailbox: mailboxName,
        uid: message.uid,
        date: message.internalDate ? new Date(message.internalDate) : new Date(0),
        subject,
        messageId: meta.messageId,
        inReplyTo: meta.inReplyTo,
        references: meta.references,
        from: addressListToEmails(message.envelope?.from),
        replyTo: addressListToEmails(message.envelope?.replyTo),
        to: addressListToEmails(message.envelope?.to),
        cc: addressListToEmails(message.envelope?.cc)
      });
    }
  } finally {
    lock.release();
  }

  return results;
}

async function findLatestThreadByPedido(pedidoBase) {
  const client = createImapClient();
  const inboxMailbox = process.env.IMAP_INBOX_MAILBOX || 'INBOX';
  const sentMailbox = process.env.IMAP_SENT_MAILBOX || 'Sent';

  await client.connect();

  try {
    const inboxMatches = await inspectMailboxForPedido(client, inboxMailbox, pedidoBase);
    const sentMatches = await inspectMailboxForPedido(client, sentMailbox, pedidoBase);

    const allMatches = [...inboxMatches, ...sentMatches]
      .sort((a, b) => b.date - a.date);

    return allMatches[0] || null;
  } catch (error) {
    console.error('[IMAP] Error buscando pedido:', error?.message || error);

    const err = new Error('No se pudo completar la búsqueda de correos.');
    err.code = 'THREAD_SEARCH_TIMEOUT';
    throw err;
  } finally {
    try {
      await client.logout();
    } catch (_) {
      // ignorar
    }
  }
}

function buildReplyRecipients(thread) {
  const sender = String(process.env.OUTLOOK_SENDER || '').trim().toLowerCase();
  const sentMailbox = String(process.env.IMAP_SENT_MAILBOX || 'Sent').toLowerCase();

  let toList = [];
  let ccList = [];

  if (String(thread.mailbox || '').toLowerCase() === sentMailbox) {
    toList = [...thread.to];
    ccList = [...thread.cc];
  } else {
    toList = thread.replyTo.length ? [...thread.replyTo] : [...thread.from];
    ccList = [...thread.to, ...thread.cc];
  }

  toList = uniqueEmails(toList).filter(email => email !== sender);
  ccList = uniqueEmails(ccList).filter(email => email !== sender && !toList.includes(email));

  if (!toList.length) {
    toList = uniqueEmails(
      String(process.env.OUTLOOK_DEFAULT_TO || '')
        .split(',')
        .map(v => v.trim())
    );
  }

  return {
    to: toList.join(', '),
    cc: ccList.length ? ccList.join(', ') : undefined
  };
}

transporter.verify((error, success) => {
  if (error) {
    console.log("❌ Error SMTP:", error);
  } else {
    console.log("✅ SMTP listo");
  }
});

const TEMP_UPLOAD_DIR = process.env.TEMP_UPLOAD_DIR || './temp_uploads';
const MEGACMD_DIR = process.env.MEGACMD_DIR || '';

if (!fs.existsSync(TEMP_UPLOAD_DIR)) {
  fs.mkdirSync(TEMP_UPLOAD_DIR, { recursive: true });
}

const upload = multer({
  storage: multer.diskStorage({
    destination: function (req, file, cb) {
      cb(null, TEMP_UPLOAD_DIR);
    },
    filename: function (req, file, cb) {
      const safeName = `${Date.now()}-${file.originalname.replace(/[^\w.\- ]+/g, '_')}`;
      cb(null, safeName);
    }
  }),
  limits: {
    fileSize: 20 * 1024 * 1024,
    files: 30
  }
});

app.use(cors({
  origin: process.env.FRONTEND_ORIGIN?.split(',').map(v => v.trim()) || '*'
}));

app.use(express.json({ limit: '10mb' }));

/* =========================
   HELPERS GENERALES
========================= */
function sanitizeName(value = '') {
  return value
    .trim()
    .replace(/[\\/:*?"<>|#%&{}[\]]+/g, '')
    .replace(/\s+/g, ' ')
    .substring(0, 80);
}

function buildMegaClientFolderName(clientName = '', vinChasis = '') {
  const safeClient = sanitizeName(clientName);
  const safeVin = sanitizeName(vinChasis);

  if (safeClient && safeVin) {
    return `${safeClient} ${safeVin}`;
  }

  if (safeClient) return safeClient;
  if (safeVin) return safeVin;

  return 'Cliente sin nombre';
}

function getMonthFolderName(date = new Date()) {
  const meses = [
    'Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
    'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'
  ];
  return `${meses[date.getMonth()]} ${date.getFullYear()}`;
}

function normalizeMegaPath(...parts) {
  const cleaned = parts
    .map(p => String(p || '').trim())
    .filter(Boolean)
    .map(p => p.replace(/^\/+|\/+$/g, ''));
  return '/' + cleaned.join('/');
}

function logStep(step, extra = '') {
  console.log(`[MEGAcmd] ${step}${extra ? ` -> ${extra}` : ''}`);
}

function getMegaBat(commandName) {
  if (!MEGACMD_DIR) {
    throw new Error('MEGACMD_DIR no está definido en .env');
  }

  const fullPath = path.join(MEGACMD_DIR, `mega-${commandName}.bat`);

  if (!fs.existsSync(fullPath)) {
    throw new Error(`No se encontró el comando MEGAcmd: ${fullPath}`);
  }

  return fullPath;
}

function quoteCmdArg(value) {
  const str = String(value ?? '');
  return `"${str.replace(/"/g, '""')}"`;
}

function cleanMegaText(text = '') {
  return String(text)
    .replace(/\x1B\[[0-9;]*[A-Za-z]/g, '') // ANSI
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(line => {
      if (!line) return false;
      if (/^TRANSFERRING/i.test(line)) return false;
      if (/^#+/.test(line)) return false;
      if (/^\d+(\.\d+)?\s*%/.test(line)) return false;
      return true;
    })
    .join('\n')
    .trim();
}

async function runMegaCmd(commandName, args = []) {
  const batFile = getMegaBat(commandName);

  const quotedBat = `"${batFile}"`;
  const quotedArgs = args.map(arg => `"${String(arg).replace(/"/g, '""')}"`).join(' ');
  const fullCommand = `${quotedBat}${quotedArgs ? ' ' + quotedArgs : ''}`;

  logStep(`Ejecutando mega-${commandName}`, fullCommand);

  try {
    const { stdout, stderr } = await execAsync(fullCommand, {
      windowsHide: true,
      shell: 'cmd.exe'
    });

    const out = cleanMegaText(stdout || '');
    const err = cleanMegaText(stderr || '');

    if (err) {
      const lowered = err.toLowerCase();
      if (
        lowered.includes('err') ||
        lowered.includes('error') ||
        lowered.includes('failed') ||
        lowered.includes('access violation') ||
        lowered.includes('no se reconoce como un comando')
      ) {
        throw new Error(err);
      }
    }

    return out;
  } catch (error) {
    const stderr = cleanMegaText(error?.stderr || '');
    const stdout = cleanMegaText(error?.stdout || '');
    const rawMessage = error?.message || String(error);

    const message =
      stderr ||
      stdout ||
      rawMessage ||
      'Ocurrió un error al ejecutar MEGAcmd.';

    throw new Error(message);
  }
}

async function megaWhoAmI() {
  return await runMegaCmd('whoami');
}

async function megaEnsureLoggedIn() {
  const who = await megaWhoAmI();
  if (!who || !who.trim()) {
    throw new Error('MEGAcmd no tiene sesión activa. Inicia sesión manualmente primero.');
  }
  return who;
}

async function megaPathExists(megaPath) {
  try {
    await runMegaCmd('ls', [megaPath]);
    return true;
  } catch {
    return false;
  }
}

async function megaEnsureFolder(megaPath) {
  const exists = await megaPathExists(megaPath);
  if (exists) return;

  await runMegaCmd('mkdir', [megaPath]);
}

async function megaUploadFile(localFilePath, megaFolderPath) {
  await runMegaCmd('put', [localFilePath, megaFolderPath]);
}

async function megaExportFolder(megaFolderPath) {
  try {
    const output = await runMegaCmd('export', ['-a', megaFolderPath]);
    const match = output.match(/https:\/\/mega\.nz\/[^\s]+/i);

    if (!match) {
      throw new Error(`No se pudo extraer el link desde la salida de MEGAcmd: ${output}`);
    }

    return match[0];
  } catch (error) {
    const message = String(error.message || error);

    // Si ya estaba exportada, intentamos obtener el link existente
    if (message.toLowerCase().includes('already exported')) {
      const output = await runMegaCmd('export', [megaFolderPath]);
      const match = output.match(/https:\/\/mega\.nz\/[^\s]+/i);

      if (!match) {
        throw new Error(
          `La carpeta ya estaba exportada, pero no se pudo recuperar el link existente. Salida: ${output}`
        );
      }

      return match[0];
    }

    throw error;
  }
}

async function uploadFilesToMega(clientName, vinChasis, files) {
  const basePath = process.env.MEGA_BASE_PATH;
  if (!basePath) {
    throw new Error('MEGA_BASE_PATH no está definido en .env');
  }

  await megaEnsureLoggedIn();

  const monthFolderName = getMonthFolderName(new Date());
  const clientFolderName = buildMegaClientFolderName(clientName, vinChasis);

  const monthFolderPath = normalizeMegaPath(basePath, monthFolderName);
  const clientFolderPath = normalizeMegaPath(basePath, monthFolderName, clientFolderName);

  await megaEnsureFolder(basePath);
  await megaEnsureFolder(monthFolderPath);
  await megaEnsureFolder(clientFolderPath);

  for (const file of files) {
    await megaUploadFile(file.path, clientFolderPath);
  }

  const folderLink = await megaExportFolder(clientFolderPath);

  return {
    monthFolderName,
    clientFolderName,
    folderLink
  };
}

function cleanupUploadedFiles(files = []) {
  for (const file of files) {
    try {
      if (file?.path && fs.existsSync(file.path)) {
        fs.unlinkSync(file.path);
      }
    } catch (error) {
      console.warn('No se pudo borrar archivo temporal:', file?.path, error.message);
    }
  }
}



function buildTestEmailHtml(payload) {
  const { data = {}, megaFolderLink = '', pedidoBase = '' } = payload || {};

  const cliente = pickFirstDefined(data, ['nombreCliente']) || '';
  const pedido = String(data?.pedidoRaw || pedidoBase || '').trim();
  const vin = pickFirstDefined(data, ['vinChasis', 'vin']) || '';
  const placa = pickFirstDefined(data, ['placa']) || '';
  const marca = pickFirstDefined(data, ['marca']) || '';
  const modelo = pickFirstDefined(data, ['numeroVehiculo', 'modelo']) || '';
  const cono = pickFirstDefined(data, ['numeroColorCono', 'cono', 'tag']) || '';

  const detallesHtml = buildDetallesObservacionesHtml(data);
  const notasHtml = buildNotasHtml(data);

  return `
    <div style="font-family: Arial, sans-serif; color:#222; line-height:1.5; font-size:14px;">
      <p style="margin:0;">Buen día</p>

      <div style="height:22px;"></div>

      <p style="margin:0 0 16px 0;">
        Este correo es para informar que el día de hoy ingresó al plantel el vehículo del cliente:
        <strong>${escapeHtml(cliente || 'no definido')}</strong>
      </p>

      <table style="border-collapse:collapse; width:100%; font-size:14px; margin-top:10px;">
        <thead>
          <tr style="background:#b7d7a8;">
            <th style="border:1px solid #5b7f4a; padding:8px; text-align:left;">Cliente</th>
            <th style="border:1px solid #5b7f4a; padding:8px; text-align:left;">Pedido</th>
            <th style="border:1px solid #5b7f4a; padding:8px; text-align:left;">VIN / Placa</th>
            <th style="border:1px solid #5b7f4a; padding:8px; text-align:left;">Marca</th>
            <th style="border:1px solid #5b7f4a; padding:8px; text-align:left;">Modelo</th>
            <th style="border:1px solid #5b7f4a; padding:8px; text-align:left;"># y Color cono</th>
            <th style="border:1px solid #5b7f4a; padding:8px; text-align:left;">Detalles / Observaciones</th>
          </tr>
        </thead>
        <tbody>
          <tr>
            <td style="border:1px solid #7f7f7f; padding:8px; vertical-align:top;">${valueOrNoDef(cliente)}</td>
            <td style="border:1px solid #7f7f7f; padding:8px; vertical-align:top;">${valueOrNoDef(pedido)}</td>
            <td style="border:1px solid #7f7f7f; padding:8px; vertical-align:top;">
              ${valueOrNoDef(vin)} / ${valueOrNoDef(placa)}
            </td>
            <td style="border:1px solid #7f7f7f; padding:8px; vertical-align:top;">${valueOrNoDef(marca)}</td>
            <td style="border:1px solid #7f7f7f; padding:8px; vertical-align:top;">${valueOrNoDef(modelo)}</td>
            <td style="border:1px solid #7f7f7f; padding:8px; vertical-align:top;">${valueOrNoDef(cono)}</td>
            <td style="border:1px solid #7f7f7f; padding:8px; vertical-align:top;">${detallesHtml}</td>
          </tr>
        </tbody>
      </table>

      ${notasHtml}

      <div style="margin-top:14px;">
        <strong>Fotografías:</strong>
        <a href="${escapeHtml(megaFolderLink)}" target="_blank">${escapeHtml(megaFolderLink)}</a>
      </div>

      <div style="margin-top:18px; color:#666; font-size:12px;">
        Este correo fue generado automáticamente por el sistema de ingreso vehicular creado por Henry Carbajal.
      </div>
    </div>
  `;
}


/* =========================
   RUTAS
========================= */
app.get('/api/health', (req, res) => {
  res.json({ ok: true, message: 'Backend activo' });
});

app.get('/api/mega/test-link', async (req, res) => {
  try {
    await megaEnsureLoggedIn();

    const basePath = process.env.MEGA_BASE_PATH;
    await megaEnsureFolder(basePath);

    const testFolderPath = normalizeMegaPath(basePath, 'PRUEBA LINK');
    await megaEnsureFolder(testFolderPath);
    const folderLink = await megaExportFolder(testFolderPath);

    res.json({
      ok: true,
      folderPath: testFolderPath,
      folderLink
    });
  } catch (error) {
    console.error('Error en /api/mega/test-link:', error);
    res.status(500).json({
      ok: false,
      error: error.message || String(error)
    });
  }
});

app.post('/api/mega/upload', upload.array('photos', 30), async (req, res) => {
  const files = req.files || [];

  try {
    const clientName = sanitizeName(req.body.clientName || '');
    const vinChasis = sanitizeName(req.body.vinChasis || '');

    if (!clientName) {
      cleanupUploadedFiles(files);
      return res.status(400).json({ error: 'El nombre del cliente es obligatorio.' });
    }

    if (!files.length) {
      return res.status(400).json({ error: 'No se recibieron fotos.' });
    }

    const result = await uploadFilesToMega(clientName, vinChasis, files);

    cleanupUploadedFiles(files);

    res.json({
      ok: true,
      ...result
    });
  } catch (error) {
    cleanupUploadedFiles(files);
    console.error('Error en /api/mega/upload:', error);
    res.status(500).json({
      error: error.message || 'Error subiendo archivos a MEGA.'
    });
  }
});

app.post('/api/outlook/send', async (req, res) => {
  try {
    const {
      subject,
      megaFolderLink,
      data,
      pedido,
      forceSendWithoutThread = false
    } = req.body || {};

    const pedidoBase = normalizePedidoBase(pedido);

    if (!pedidoBase) {
      return res.status(400).json({
        error: 'Debes ingresar un número de pedido válido. Ejemplo: 6949-1'
      });
    }

    if (!megaFolderLink) {
      return res.status(400).json({
        error: 'No existe link de MEGA.'
      });
    }

    const cliente = data?.nombreCliente || 'Cliente no especificado';
    const vehiculo = [
      data?.marca || '',
      data?.numeroVehiculo || '',
      data?.placa || ''
    ].filter(Boolean).join(' / ') || 'No especificado';

    const html = buildTestEmailHtml({
      data,
      megaFolderLink,
      inspectionSummary: {
        totalMarcados: Array.isArray(data?.checkboxStates)
          ? data.checkboxStates.filter(Boolean).length
          : 0
      },
      pedidoBase
    });

    let thread = null;

    // SOLO buscar hilo si no se forzó enviar como nuevo
    if (!forceSendWithoutThread) {
      try {
        thread = await findLatestThreadByPedido(pedidoBase);

        if (!thread) {
          return res.status(404).json({
            code: 'THREAD_NOT_FOUND',
            error: `No se encontró correos con el pedido ${pedidoBase}.`,
            pedidoBase
          });
        }
      } catch (searchError) {
        return res.status(408).json({
          code: searchError.code || 'THREAD_SEARCH_TIMEOUT',
          error: searchError.message || 'No se pudo completar la búsqueda de correos.',
          pedidoBase
        });
      }
    }

    const asuntoNuevo = subject || (
      pedidoBase
        ? `Ingreso Vehicular ${pedidoBase}, ${cliente}`
        : `Ingreso Vehicular ${cliente}`
    );

    let mailOptions = {
      from: `"CARGUS - Ingreso Vehicular" <${process.env.OUTLOOK_SENDER}>`,
      to: process.env.OUTLOOK_DEFAULT_TO,
      subject: asuntoNuevo,
      html
    };

    if (thread) {
      const recipients = buildReplyRecipients(thread);
      const replySubject = thread.subject
        ? (/^re:/i.test(thread.subject) ? thread.subject : `RE: ${thread.subject}`)
        : (subject || `Ingreso Vehicular - Pedido ${pedidoBase} - ${cliente}`);

      mailOptions = {
        ...mailOptions,
        to: recipients.to || process.env.OUTLOOK_DEFAULT_TO,
        cc: recipients.cc,
        subject: replySubject,
        inReplyTo: thread.messageId || undefined,
        references: thread.references?.length
          ? [...thread.references, thread.messageId].filter(Boolean)
          : (thread.messageId ? [thread.messageId] : undefined)
      };
    }

    const info = await transporter.sendMail(mailOptions);
    await appendToSentMailbox(mailOptions);

    res.json({
      ok: true,
      mode: thread ? 'reply' : 'new',
      pedidoBase,
      messageId: info.messageId
    });
  } catch (error) {
    console.error('Error en /api/outlook/send:', error);
    res.status(500).json({
      error: error.message || 'Error enviando correo.'
    });
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Servidor en puerto ${PORT}`);
});

app.get('/', (req, res) => {
  res.send('Servidor de Ingreso Vehicular activo');
});