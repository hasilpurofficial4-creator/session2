/**
 * UNIT STOCK MANAGEMENT - WhatsApp Bot + API Server
 * Runs on Railway as a standalone Express + Baileys bot service.
 *
 * Features:
 *  - Express API: /send, /status, / (health)
 *  - Auth middleware (API_SECRET)
 *  - Outbox queue with auto-send every 30s
 *  - .{category} → Excel file | .{category}2 → text details
 *  - Rich emoji formatting on all messages
 */

const {
  default: makeWASocket,
  useMultiFileAuthState,
  DisconnectReason,
} = require('@whiskeysockets/baileys');
const pino = require('pino');
const fs = require('fs');
const path = require('path');
const express = require('express');
const ExcelJS = require('exceljs');
const axios = require('axios');

// ─── Config ──────────────────────────────────────────────────────────────────
const PORT = parseInt(process.env.PORT) || 3001;
const API_SECRET = process.env.API_SECRET || 'banu-saeed-secret-2024';
const ADMIN_NUMBER = process.env.ADMIN_NUMBER || '923291001302';
const SESSION_DIR = path.join(__dirname, 'session');
const OUTBOX_PATH = path.join(__dirname, 'data', 'outbox.json');
const GITHUB_RAW = 'https://raw.githubusercontent.com/hasilpurofficial4-creator/zaid2/main/data';
const BOT_NAME = 'UNIT STOCK MANAGEMENT';

// ─── State Tracking ─────────────────────────────────────────────────────────
let whatsappConnected = false;
let botRunning = false;
let sockRef = null;
let sentCount = 0;
let botLogs = [];
const startTime = Date.now();

function log(msg) {
  const ts = new Date().toISOString().slice(11, 19);
  const line = `[${ts}] ${msg}`;
  console.log(line);
  botLogs.push(line);
  if (botLogs.length > 50) botLogs = botLogs.slice(-50);
}

// ─── Outbox Helpers ──────────────────────────────────────────────────────────
function readOutbox() {
  try {
    if (!fs.existsSync(OUTBOX_PATH)) return [];
    return JSON.parse(fs.readFileSync(OUTBOX_PATH, 'utf8'));
  } catch { return []; }
}

function writeOutbox(data) {
  try {
    const dir = path.dirname(OUTBOX_PATH);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(OUTBOX_PATH, JSON.stringify(data, null, 2), 'utf8');
  } catch (err) {
    log('[OUTBOX] Write error: ' + err.message);
  }
}

// ─── Stock Data Fetch ────────────────────────────────────────────────────────
async function fetchStockData(section) {
  try {
    const res = await axios.get(`${GITHUB_RAW}/${section}.json`, { timeout: 10000 });
    return Array.isArray(res.data) ? res.data : [];
  } catch (err) {
    log('[STOCK] Fetch ' + section + ' error: ' + err.message);
    return [];
  }
}

// ─── Excel Generators ────────────────────────────────────────────────────────
function styleHeader(ws) {
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1B5E20' } };
  ws.getRow(1).alignment = { horizontal: 'center' };
}

async function generateItemsXlsx(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Items');
  ws.columns = [
    { header: '#', key: 'no', width: 5 },
    { header: 'Name', key: 'name', width: 25 },
    { header: 'Serial No', key: 'number', width: 18 },
    { header: 'Model', key: 'model', width: 15 },
    { header: 'Qty', key: 'quantity', width: 8 },
    { header: 'Person', key: 'person', width: 18 },
    { header: 'Status', key: 'status', width: 12 },
    { header: 'Date', key: 'timestamp', width: 20 },
  ];
  styleHeader(ws);
  data.forEach((e, i) => {
    ws.addRow({
      no: i + 1, name: e.name || '', number: e.number || '', model: e.model || '',
      quantity: e.quantity || 1, person: e.person || '', status: e.status || 'available',
      timestamp: e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB') : ''
    });
  });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function generateWalletXlsx(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Wallet');
  ws.columns = [
    { header: '#', key: 'no', width: 5 },
    { header: 'Type', key: 'type', width: 10 },
    { header: 'From / For', key: 'personOrPurpose', width: 25 },
    { header: 'Amount (Rs)', key: 'amount', width: 15 },
    { header: 'Date', key: 'timestamp', width: 20 },
  ];
  styleHeader(ws);
  let balance = 0;
  data.forEach((e, i) => {
    const amt = Number(e.amount) || 0;
    if (e.type === 'in') balance += amt; else balance -= amt;
    ws.addRow({
      no: i + 1, type: (e.type || '').toUpperCase(),
      personOrPurpose: e.personOrPurpose || '', amount: amt,
      timestamp: e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB') : ''
    });
  });
  ws.addRow({});
  const totalIn = data.filter(e => e.type === 'in').reduce((a, e) => a + Number(e.amount || 0), 0);
  const totalOut = data.filter(e => e.type === 'out').reduce((a, e) => a + Number(e.amount || 0), 0);
  const r1 = ws.addRow({ personOrPurpose: 'TOTAL RECEIVED', amount: totalIn }); r1.font = { bold: true, color: { argb: 'FF00AA00' } };
  const r2 = ws.addRow({ personOrPurpose: 'TOTAL SPENT', amount: totalOut }); r2.font = { bold: true, color: { argb: 'FFCC0000' } };
  const r3 = ws.addRow({ personOrPurpose: 'BALANCE', amount: balance }); r3.font = { bold: true, size: 14 };
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function generatePersonXlsx(data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Person');
  ws.columns = [
    { header: '#', key: 'no', width: 5 },
    { header: 'Name', key: 'personName', width: 22 },
    { header: 'Action', key: 'action', width: 10 },
    { header: 'Date & Time', key: 'timestamp', width: 22 },
  ];
  styleHeader(ws);
  data.forEach((e, i) => {
    ws.addRow({
      no: i + 1, personName: e.personName || '', action: e.action || '',
      timestamp: e.timestamp ? new Date(e.timestamp).toLocaleString('en-GB') : ''
    });
  });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

async function generateGenericXlsx(section, data) {
  if (!data.length) return null;
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(section.charAt(0).toUpperCase() + section.slice(1));
  const keys = Object.keys(data[0]).filter(k => k !== 'id');
  ws.columns = [
    { header: '#', key: 'no', width: 5 },
    ...keys.map(k => ({ header: k.charAt(0).toUpperCase() + k.slice(1), key: k, width: 20 }))
  ];
  styleHeader(ws);
  data.forEach((e, i) => {
    const row = { no: i + 1 };
    keys.forEach(k => { row[k] = e[k] != null ? String(e[k]) : ''; });
    ws.addRow(row);
  });
  return Buffer.from(await wb.xlsx.writeBuffer());
}

// ─── Text Detail Formatters (for .{category}2 commands) ──────────────────────
function fmtDate(ts) { return ts ? new Date(ts).toLocaleDateString('en-GB') : 'N/A'; }
function fmtDateTime(ts) { return ts ? new Date(ts).toLocaleString('en-GB') : 'N/A'; }
const LINE = '╔══════════════════════════════╗';
const DIV  = '╠══════════════════════════════╣';
const END  = '╚══════════════════════════════╝';

function formatItemsText(data) {
  let m = `📦 ✦ *𝗜𝗧𝗘𝗠𝗦 𝗜𝗡𝗩𝗘𝗡𝗧𝗢𝗥𝗬* ✦ 📦\n${LINE}\n`;
  m += `📊 *𝗧𝗼𝗧𝗮𝗹 𝗜𝗧𝗲𝗺𝘀:* ${data.length}\n${DIV}\n\n`;
  data.forEach((e, i) => {
    const st = e.status || 'available';
    const stIcon = st === 'available' ? '🟢' : st === 'in-use' ? '🔵' : '🔴';
    m += `${stIcon} *${i + 1}. ${e.name || 'Unnamed'}*\n`;
    m += `   🔢 Serial: ${e.number || 'N/A'}\n`;
    m += `   📐 Model: ${e.model || 'N/A'}\n`;
    m += `   📦 Qty: ${e.quantity || 1}\n`;
    m += `   👤 Person: ${e.person || 'N/A'}\n`;
    m += `   📅 ${fmtDate(e.timestamp)}\n`;
    if (i < data.length - 1) m += `   ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─\n`;
  });
  m += `\n${END}\n🏢 *${BOT_NAME}*\n🌐 https://zaidbwp.vercel.app`;
  return m;
}

function formatWalletText(data) {
  const totalIn = data.filter(e => e.type === 'in').reduce((a, e) => a + Number(e.amount || 0), 0);
  const totalOut = data.filter(e => e.type === 'out').reduce((a, e) => a + Number(e.amount || 0), 0);
  const bal = totalIn - totalOut;
  const balEmoji = bal >= 0 ? '✅' : '⚠️';
  let m = `💰 ✦ *𝗪𝗔𝗟𝗟𝗘𝗧 𝗥𝗘𝗣𝗢𝗥𝗧* ✦ 💰\n${LINE}\n`;
  m += `📥 *𝗧𝗼𝗧𝗮𝗹 𝗥𝗲𝗰𝗲𝗶𝘃𝗲𝗱:* Rs. ${totalIn.toLocaleString()}\n`;
  m += `📤 *𝗧𝗼𝗧𝗮𝗹 𝗦𝗽𝗲𝗻𝘁:* Rs. ${totalOut.toLocaleString()}\n`;
  m += `🏦 *𝗕𝗮𝗹𝗮𝗻𝗰𝗲:* ${balEmoji} Rs. ${bal.toLocaleString()}\n`;
  m += `📊 *𝗘𝗻𝘁𝗿𝗶𝗲𝘀:* ${data.length}\n${DIV}\n\n`;
  data.forEach((e, i) => {
    const amt = Number(e.amount) || 0;
    const icon = e.type === 'in' ? '📥' : '📤';
    m += `${icon} *${(e.type || '').toUpperCase()}* — Rs. ${amt.toLocaleString()}\n`;
    m += `  👤 ${e.personOrPurpose || 'N/A'}\n`;
    m += `  📅 ${fmtDate(e.timestamp)}\n`;
    if (i < data.length - 1) m += `  ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─\n`;
  });
  m += `\n${END}\n🏢 *${BOT_NAME}*\n🌐 https://zaidbwp.vercel.app`;
  return m;
}

function formatPersonText(data) {
  const workers = [...new Set(data.map(e => e.personName).filter(Boolean))].sort();
  let m = `👷 ✦ *𝗣𝗘𝗥𝗦𝗢𝗡 𝗔𝗧𝗧𝗘𝗡𝗗𝗔𝗡𝗖𝗘* ✦ 👷\n${LINE}\n`;
  m += `👥 *𝗪𝗼𝗿𝗸𝗲𝗿𝘀:* ${workers.length}\n`;
  m += `📋 *𝗧𝗼𝗧𝗮𝗹 𝗘𝗻𝘁𝗿𝗶𝗲𝘀:* ${data.length}\n${DIV}\n\n`;
  const sorted = [...data].sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
  sorted.forEach((e, i) => {
    const icon = e.action === 'enter' ? '🟢' : '🔴';
    const act = e.action === 'enter' ? 'CHECKED IN' : 'CHECKED OUT';
    m += `${icon} *${e.personName || 'Unknown'}* — ${act}\n`;
    m += `  📅 ${fmtDateTime(e.timestamp)}\n`;
    if (i < sorted.length - 1) m += `  ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─\n`;
  });
  m += `\n${END}\n🏢 *${BOT_NAME}*\n🌐 https://zaidbwp.vercel.app`;
  return m;
}

function formatMaintenanceText(data) {
  const open = data.filter(e => e.status !== 'solved').length;
  const solved = data.length - open;
  let m = `🔧 ✦ *𝗠𝗔𝗜𝗡𝗧𝗘𝗡𝗔𝗡𝗖𝗘 𝗥𝗘𝗣𝗢𝗥𝗧* ✦ 🔧\n${LINE}\n`;
  m += `📊 *𝗧𝗼𝗧𝗮𝗹:* ${data.length} | 🔴 *𝗢𝗽𝗲𝗻:* ${open} | ✅ *𝗦𝗼𝗹𝘃𝗲𝗱:* ${solved}\n${DIV}\n\n`;
  data.forEach((e, i) => {
    const stIcon = e.status === 'solved' ? '✅' : '🔴';
    m += `${stIcon} *#${i + 1} — ${e.category || 'Issue'}*\n`;
    m += `  📝 ${e.subject || 'N/A'}\n`;
    m += `  📄 ${e.description || 'N/A'}\n`;
    m += `  📅 ${fmtDate(e.timestamp)}\n`;
    if (e.solvedAt) m += `  ✅ Solved: ${fmtDate(e.solvedAt)}\n`;
    if (i < data.length - 1) m += `  ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─\n`;
  });
  m += `\n${END}\n🏢 *${BOT_NAME}*\n🌐 https://zaidbwp.vercel.app`;
  return m;
}

function formatSamplesText(data) {
  const inCount = data.filter(e => e.type === 'in').length;
  const outCount = data.length - inCount;
  let m = `🧪 ✦ *𝗦𝗔𝗠𝗣𝗟𝗘𝗦 𝗥𝗘𝗣𝗢𝗥𝗧* ✦ 🧪\n${LINE}\n`;
  m += `📥 *𝗜𝗻:* ${inCount} | 📤 *𝗢𝘂𝘁:* ${outCount} | 📊 *𝗧𝗼𝗧𝗮𝗹:* ${data.length}\n${DIV}\n\n`;
  data.forEach((e, i) => {
    const icon = e.type === 'in' ? '📥' : '📤';
    const label = e.type === 'in' ? 'RECEIVED' : 'SENT OUT';
    m += `${icon} *${label}* — ${e.personName || 'N/A'}\n`;
    m += `  📋 Program: ${e.program || 'N/A'}\n`;
    m += `  🔢 Pieces: ${e.pieces || 'N/A'}\n`;
    m += `  📅 ${fmtDate(e.timestamp)}\n`;
    if (i < data.length - 1) m += `  ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─\n`;
  });
  m += `\n${END}\n🏢 *${BOT_NAME}*\n🌐 https://zaidbwp.vercel.app`;
  return m;
}

function formatClippingText(data) {
  const inEntries = data.filter(e => e.type === 'in');
  let totalSize = 0;
  inEntries.forEach(e => { const n = parseFloat(e.size); if (!isNaN(n)) totalSize += n; });
  const totalPay = totalSize * 12;
  let m = `✂️ ✦ *𝗖𝗟𝗜𝗣𝗣𝗜𝗡𝗚 𝗥𝗘𝗣𝗢𝗥𝗧* ✂️\n${LINE}\n`;
  m += `📊 *𝗧𝗼𝗧𝗮𝗹:* ${data.length} entries\n`;
  m += `📏 *𝗧𝗼𝗧𝗮𝗹 𝗬𝗮𝗿𝗱𝘀:* ${totalSize}\n`;
  m += `💰 *𝗧𝗼𝗧𝗮𝗹 𝗣𝗮𝘆𝗺𝗲𝗻𝘁:* Rs. ${totalPay.toLocaleString()} (${totalSize}×12)\n${DIV}\n\n`;
  data.forEach((e, i) => {
    let icon, label;
    if (e.type === 'in') { icon = '📥'; label = 'CLIPPED IN'; }
    else if (e.type === 'transfer') { icon = '💸'; label = 'TRANSFER'; }
    else { icon = '📤'; label = 'OUT FOR CLIPPING'; }
    m += `${icon} *${label}* — ${e.clipperName || 'N/A'}\n`;
    m += `  📏 Size: ${e.size || 'N/A'} yards\n`;
    m += `  📅 ${fmtDate(e.timestamp)}\n`;
    if (i < data.length - 1) m += `  ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─ ─\n`;
  });
  m += `\n${END}\n🏢 *${BOT_NAME}*\n🌐 https://zaidbwp.vercel.app`;
  return m;
}

// ─── Send File Helper ─────────────────────────────────────────────────────────
async function sendXlsx(sock, chatId, data, genFn, fileName, caption) {
  if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No data found.' }); return; }
  const buf = await genFn(data);
  if (!buf) { await sock.sendMessage(chatId, { text: '❌ Could not generate file.' }); return; }
  const tmpDir = path.join(__dirname, 'tmp');
  if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });
  const tmpFile = path.join(tmpDir, fileName + '_' + Date.now() + '.xlsx');
  fs.writeFileSync(tmpFile, buf);
  await sock.sendMessage(chatId, {
    document: fs.readFileSync(tmpFile), fileName: fileName + '.xlsx',
    mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    caption
  });
  fs.unlinkSync(tmpFile);
}

// ─── Session Restore ──────────────────────────────────────────────────────────
let loggedOutOnce = false;

function restoreSession() {
  const sessionId = process.env.SESSION_ID;
  if (!sessionId) { log('[SESSION] No SESSION_ID. Will generate QR.'); return false; }
  try {
    if (!fs.existsSync(SESSION_DIR)) fs.mkdirSync(SESSION_DIR, { recursive: true });
    const credsFile = path.join(SESSION_DIR, 'creds.json');
    if (fs.existsSync(credsFile)) {
      const creds = JSON.parse(fs.readFileSync(credsFile, 'utf8'));
      if (creds.me) { log('[SESSION] Valid session found for ' + (creds.me.id || '?')); return true; }
    }
    const base64Data = sessionId.includes('~') ? sessionId.split('~').slice(1).join('~') : sessionId;
    const buffer = Buffer.from(base64Data, 'base64');
    try {
      const AdmZip = require('adm-zip');
      const zip = new AdmZip(buffer);
      if (zip.getEntries().length > 0) { zip.extractAllTo(SESSION_DIR, true); return true; }
    } catch (_) {}
    try {
      const json = JSON.parse(buffer.toString('utf8'));
      if (json.creds || json.me || json.account) {
        fs.writeFileSync(credsFile, JSON.stringify(json, null, 2), 'utf8'); return true;
      }
      fs.writeFileSync(credsFile, JSON.stringify(json, null, 2), 'utf8'); return true;
    } catch (_) {}
    return false;
  } catch (err) { log('[SESSION] Restore error: ' + err.message); return false; }
}

// ─── Outbox Processing ────────────────────────────────────────────────────────
async function processOutbox(sock) {
  const outbox = readOutbox();
  const pending = outbox.filter(m => !m.sent);
  if (!pending.length) return { processed: 0, failed: 0 };
  let processed = 0, failed = 0;
  for (const entry of pending) {
    try {
      const target = entry.to.replace(/[^0-9]/g, '') + '@s.whatsapp.net';
      await sock.sendMessage(target, { text: entry.message });
      entry.sent = true; entry.sentAt = new Date().toISOString(); entry.error = null;
      processed++; sentCount++;
      log('[OUTBOX] ✅ Sent to +' + entry.to);
    } catch (err) {
      entry.error = err.message; failed++;
      log('[OUTBOX] ❌ Failed +' + entry.to + ': ' + err.message);
    }
  }
  const cleaned = outbox.filter(m => !m.sent || (Date.now() - new Date(m.sentAt).getTime()) < 3600000);
  writeOutbox(cleaned.length > 200 ? cleaned.slice(-200) : cleaned);
  return { processed, failed };
}

// ─── Express API Server ───────────────────────────────────────────────────────
const app = express();
app.use(express.json({ limit: '1mb' }));

function authMiddleware(req, res, next) {
  const secret = req.body?.secret;
  if (secret !== API_SECRET) return res.status(401).json({ success: false, error: 'Unauthorized' });
  next();
}

// Health check
app.get('/', (req, res) => {
  res.json({ bot: BOT_NAME, online: true, whatsappConnected, uptime: Math.floor((Date.now() - startTime) / 1000) + 's' });
});

// Status endpoint (called by Vercel wa-status.js)
app.get('/status', (req, res) => {
  const pending = readOutbox().filter(m => !m.sent).length;
  res.json({
    online: true, botRunning, whatsappConnected,
    pendingMessages: pending, sentMessages: sentCount,
    adminNumber: ADMIN_NUMBER,
    uptime: Math.floor((Date.now() - startTime) / 1000) + 's',
    recentLogs: botLogs.slice(-10)
  });
});

// Send endpoint (called by Vercel _notify-helper.js & whatsapp-send.js)
app.post('/send', authMiddleware, (req, res) => {
  const { message, to } = req.body;
  if (!message) return res.status(400).json({ success: false, error: 'message required' });
  const target = (to || ADMIN_NUMBER).replace(/[^0-9]/g, '');

  const entry = {
    id: 'msg_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8),
    to: target, message, sent: false,
    createdAt: new Date().toISOString()
  };
  const outbox = readOutbox();
  outbox.push(entry);
  writeOutbox(outbox.length > 500 ? outbox.slice(-500) : outbox);
  log('[API] Queued message for +' + target);

  // Immediate send if WhatsApp connected
  if (whatsappConnected && sockRef) {
    processImmediateSend(sockRef, entry).catch(err => {
      log('[API] Immediate send failed: ' + err.message);
    });
  }
  res.json({ success: true, message: 'Message queued for +' + target, id: entry.id });
});

async function processImmediateSend(sock, entry) {
  if (entry.sent) return;
  const target = entry.to.replace(/[^0-9]/g, '') + '@s.whatsapp.net';
  await sock.sendMessage(target, { text: entry.message });
  entry.sent = true; entry.sentAt = new Date().toISOString();
  sentCount++;
  log('[IMMEDIATE] ✅ Sent to +' + entry.to);
  // Update outbox
  const outbox = readOutbox();
  const idx = outbox.findIndex(m => m.id === entry.id);
  if (idx >= 0) { outbox[idx] = entry; writeOutbox(outbox); }
}

// ─── Main Bot ─────────────────────────────────────────────────────────────────
async function startBot() {
  if (!loggedOutOnce) restoreSession();
  if (!fs.existsSync(SESSION_DIR)) fs.mkdirSync(SESSION_DIR, { recursive: true });

  const { state, saveCreds } = await useMultiFileAuthState(SESSION_DIR);

  let waVersion = [2, 2413, 1];
  try {
    const baileys = require('@whiskeysockets/baileys');
    if (typeof baileys.fetchLatestBaileysVersion === 'function') {
      const { version } = await baileys.fetchLatestBaileysVersion();
      waVersion = version;
    }
  } catch (_) {}

  const sock = makeWASocket({
    version: waVersion, auth: state,
    logger: pino({ level: 'silent' }),
    printQRInTerminal: true,
    browser: [BOT_NAME, 'Chrome', '2.0.0'],
    generateHighQualityLinkPreview: false,
  });

  sock.ev.on('creds.update', saveCreds);

  sock.ev.on('connection.update', async (update) => {
    const { connection, lastDisconnect, qr } = update;
    if (qr) log('[QR] New QR code generated — scan in WhatsApp > Linked Devices');

    if (connection === 'open') {
      whatsappConnected = true; sockRef = sock; botRunning = true;
      log('═══════════════════════════════════');
      log('  ✅ WhatsApp CONNECTED');
      log('  Account: ' + (sock.user?.id || 'unknown'));
      log('═══════════════════════════════════');

      const result = await processOutbox(sock);
      if (result.processed > 0) log('[OUTBOX] Initial: ' + result.processed + ' sent');

      setInterval(async () => {
        try {
          const r = await processOutbox(sock);
          if (r.processed > 0) log('[OUTBOX] Auto: ' + r.processed + ' sent');
        } catch (err) { log('[OUTBOX] Auto error: ' + err.message); }
      }, 30000);
    }

    if (connection === 'close') {
      const reason = lastDisconnect?.error?.output?.statusCode;
      log('[CONN] Closed. Reason: ' + reason);
      if (reason === DisconnectReason.loggedOut) {
        whatsappConnected = false; sockRef = null;
        try { fs.readdirSync(SESSION_DIR).forEach(f => fs.unlinkSync(path.join(SESSION_DIR, f))); } catch (_) {}
        if (!loggedOutOnce) {
          loggedOutOnce = true;
          log('[CONN] SESSION_ID invalid. Generating new QR...');
          setTimeout(startBot, 2000);
        } else {
          log('[CONN] Waiting 60s before retry...');
          setTimeout(startBot, 60000);
        }
      } else {
        log('[CONN] Reconnecting in 5s...');
        setTimeout(startBot, 5000);
      }
    }
  });

  // ─── Message Handler (.commands) ─────────────────────────────────────────
  sock.ev.on('messages.upsert', async ({ messages }) => {
    const msg = messages[0];
    if (!msg.message) return;
    const chatId = msg.key.remoteJid;
    const text = msg.message.conversation || msg.message.extendedTextMessage?.text || '';
    if (!text.startsWith('.')) return;
    const cmd = text.trim().split(/\s+/)[0].toLowerCase();

    try {
      // .ping
      if (cmd === '.ping') {
        const start = Date.now();
        await sock.sendMessage(chatId, { text: '⚡ Pong!' });
        await sock.sendMessage(chatId, {
          text: `🏢 *${BOT_NAME}*\n⚡ Latency: ${Date.now() - start}ms\n🟢 Status: Online\n📦 Version: 2.0`
        });
      }

      // .status
      if (cmd === '.status') {
        const pending = readOutbox().filter(m => !m.sent).length;
        await sock.sendMessage(chatId, {
          text: `🏢 *${BOT_NAME}*\n${LINE}\n🟢 Bot: Online\n📱 Account: ${sock.user?.id || 'unknown'}\n📨 Pending: ${pending}\n✅ Sent: ${sentCount}\n⏱️ Uptime: ${Math.floor((Date.now() - startTime) / 1000)}s\n${END}`
        });
      }

      // .sendoutbox
      if (cmd === '.sendoutbox') {
        const sub = (text.trim().split(/\s+/)[1] || '').toLowerCase();
        if (sub === 'status') {
          const outbox = readOutbox();
          await sock.sendMessage(chatId, {
            text: `📬 *QUEUE STATUS*\n⏳ Pending: ${outbox.filter(m => !m.sent).length}\n✅ Sent: ${outbox.filter(m => m.sent).length}\n🔄 Auto-send: every 30s`
          });
          return;
        }
        const result = await processOutbox(sock);
        const pending = readOutbox().filter(m => !m.sent).length;
        await sock.sendMessage(chatId, {
          text: `📬 *QUEUE PROCESSED*\n✅ Sent: ${result.processed}\n❌ Failed: ${result.failed}\n⏳ Remaining: ${pending}`
        });
      }

      // ═══ ITEMS ═══
      if (cmd === '.items') {
        await sock.sendMessage(chatId, { text: '📦 Fetching items...' });
        const data = await fetchStockData('items');
        await sendXlsx(sock, chatId, data, generateItemsXlsx, 'items',
          `📦 *ITEMS* (${data.length} entries)\n📅 ${new Date().toLocaleDateString('en-GB')}`);
      }
      if (cmd === '.items2') {
        await sock.sendMessage(chatId, { text: '📦 Loading items details...' });
        const data = await fetchStockData('items');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No items data.' }); return; }
        await sock.sendMessage(chatId, { text: formatItemsText(data) });
      }

      // ═══ WALLET ═══
      if (cmd === '.wallet') {
        await sock.sendMessage(chatId, { text: '💰 Fetching wallet...' });
        const data = await fetchStockData('wallet');
        const totalIn = data.filter(e => e.type === 'in').reduce((a, e) => a + Number(e.amount || 0), 0);
        const totalOut = data.filter(e => e.type === 'out').reduce((a, e) => a + Number(e.amount || 0), 0);
        await sendXlsx(sock, chatId, data, generateWalletXlsx, 'wallet',
          `💰 *WALLET*\n📥 Received: Rs. ${totalIn.toLocaleString()}\n📤 Spent: Rs. ${totalOut.toLocaleString()}\n🏦 Balance: Rs. ${(totalIn - totalOut).toLocaleString()}`);
      }
      if (cmd === '.wallet2' || cmd === '.wallettxt') {
        await sock.sendMessage(chatId, { text: '💰 Loading wallet details...' });
        const data = await fetchStockData('wallet');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No wallet data.' }); return; }
        await sock.sendMessage(chatId, { text: formatWalletText(data) });
      }

      // ═══ PERSON ═══
      if (cmd === '.person') {
        await sock.sendMessage(chatId, { text: '👷 Fetching attendance...' });
        const data = await fetchStockData('person');
        await sendXlsx(sock, chatId, data, generatePersonXlsx, 'person',
          `👷 *PERSON ATTENDANCE* (${data.length} entries)`);
      }
      if (cmd === '.person2') {
        await sock.sendMessage(chatId, { text: '👷 Loading attendance details...' });
        const data = await fetchStockData('person');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No person data.' }); return; }
        await sock.sendMessage(chatId, { text: formatPersonText(data) });
      }

      // ═══ MAINTENANCE ═══
      if (cmd === '.maintenance') {
        await sock.sendMessage(chatId, { text: '🔧 Fetching maintenance...' });
        const data = await fetchStockData('maintenance');
        const open = data.filter(e => e.status !== 'solved').length;
        await sendXlsx(sock, chatId, data, (d) => generateGenericXlsx('maintenance', d), 'maintenance',
          `🔧 *MAINTENANCE* (${data.length} issues | 🔴 ${open} open | ✅ ${data.length - open} solved)`);
      }
      if (cmd === '.maintenance2') {
        await sock.sendMessage(chatId, { text: '🔧 Loading maintenance details...' });
        const data = await fetchStockData('maintenance');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No maintenance data.' }); return; }
        await sock.sendMessage(chatId, { text: formatMaintenanceText(data) });
      }

      // ═══ SAMPLES ═══
      if (cmd === '.samples') {
        await sock.sendMessage(chatId, { text: '🧪 Fetching samples...' });
        const data = await fetchStockData('samples');
        await sendXlsx(sock, chatId, data, (d) => generateGenericXlsx('samples', d), 'samples',
          `🧪 *SAMPLES* (${data.length} entries)`);
      }
      if (cmd === '.samples2') {
        await sock.sendMessage(chatId, { text: '🧪 Loading samples details...' });
        const data = await fetchStockData('samples');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No samples data.' }); return; }
        await sock.sendMessage(chatId, { text: formatSamplesText(data) });
      }

      // ═══ CLIPPING ═══
      if (cmd === '.clipping') {
        await sock.sendMessage(chatId, { text: '✂️ Fetching clipping...' });
        const data = await fetchStockData('clipping');
        const inE = data.filter(e => e.type === 'in');
        let totalSize = 0;
        inE.forEach(e => { const n = parseFloat(e.size); if (!isNaN(n)) totalSize += n; });
        await sendXlsx(sock, chatId, data, (d) => generateGenericXlsx('clipping', d), 'clipping',
          `✂️ *CLIPPING* (${data.length} entries | ${totalSize} yards | 💰 Rs. ${(totalSize * 12).toLocaleString()})`);
      }
      if (cmd === '.clipping2') {
        await sock.sendMessage(chatId, { text: '✂️ Loading clipping details...' });
        const data = await fetchStockData('clipping');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No clipping data.' }); return; }
        await sock.sendMessage(chatId, { text: formatClippingText(data) });
      }

      // ═══ HELP ═══
      if (cmd === '.help' || cmd === '.stock') {
        const helpMsg = [
          `🏢 *${BOT_NAME}*`,
          LINE,
          '',
          '📦 *.items* — Items Excel file',
          '📦 *.items2* — Items text details',
          '',
          '💰 *.wallet* — Wallet Excel file',
          '💰 *.wallet2* — Wallet text details',
          '',
          '👷 *.person* — Person Excel file',
          '👷 *.person2* — Person text details',
          '',
          '🔧 *.maintenance* — Maint. Excel file',
          '🔧 *.maintenance2* — Maint. text details',
          '',
          '🧪 *.samples* — Samples Excel file',
          '🧪 *.samples2* — Samples text details',
          '',
          '✂️ *.clipping* — Clipping Excel file',
          '✂️ *.clipping2* — Clipping text details',
          '',
          DIV,
          '⚡ *.ping* — Check latency',
          '📊 *.status* — Bot status',
          '📬 *.sendoutbox* — Process queue',
          END,
          `🏢 *${BOT_NAME}*`,
          '🌐 https://zaidbwp.vercel.app'
        ].join('\n');
        await sock.sendMessage(chatId, { text: helpMsg });
      }

    } catch (err) {
      log('[CMD] Error ' + cmd + ': ' + err.message);
      try { await sock.sendMessage(chatId, { text: '❌ Error: ' + err.message }); } catch (_) {}
    }
  });
}

// ─── Start Express THEN Bot ──────────────────────────────────────────────────
app.listen(PORT, () => {
  log('[SERVER] ✅ API listening on port ' + PORT);
  startBot().catch(err => {
    log('[FATAL] Bot failed: ' + err.message);
  });
});

log('═══════════════════════════════════════');
log('  🏢 ' + BOT_NAME);
log('  Admin: +' + ADMIN_NUMBER);
log('═══════════════════════════════════════');

// Prevent crashes
process.on('uncaughtException', (err) => log('[BOT] Uncaught: ' + err.message));
process.on('unhandledRejection', (reason) => log('[BOT] Unhandled: ' + reason));
