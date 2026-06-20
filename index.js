/**
 * UNIT STOCK MANAGEMENT - WhatsApp Bot
 * Full Express API + Baileys WhatsApp bot for stock management.
 *
 * Features:
 * - Express API: /send (queue messages), /status (bot health)
 * - Multi-file auth state with SESSION_ID restoration
 * - Commands: .items, .items2, .wallet, .wallet2, .person, .person2, etc.
 * - Auto-send message queue every 30 seconds
 * - Immediate send when WhatsApp is connected
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
const ADMIN_NUMBER = process.env.ADMIN_NUMBER || '923299931199';
const BOT_NAME = 'UNIT STOCK MANAGEMENT';
const SITE_URL = 'https://zaidbwp.vercel.app';
const SESSION_DIR = path.join(__dirname, 'session');
const OUTBOX_PATH = path.join(__dirname, 'data', 'outbox.json');
const GITHUB_RAW = 'https://raw.githubusercontent.com/hasilpurofficial4-creator/zaid2/main/data';

let whatsappConnected = false;
let botRunning = false;
let sockRef = null;
let sentCount = 0;
let botLogs = [];
function log(msg) {
  const line = '[' + new Date().toISOString().slice(11, 19) + '] ' + msg;
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
    console.error('[OUTBOX] Write error:', err.message);
  }
}

// ─── Auth Middleware ──────────────────────────────────────────────────────

function authMiddleware(req, res, next) {
  const secret = req.body && req.body.secret;
  if (!secret || secret !== API_SECRET) {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  next();
}

// ─── Immediate Send ──────────────────────────────────────────────────────

async function processImmediateSend(sock, entry) {
  try {
    const target = entry.to.replace(/[^0-9]/g, '') + '@s.whatsapp.net';
    await sock.sendMessage(target, { text: entry.message });
    entry.sent = true;
    entry.sentAt = new Date().toISOString();
    sentCount++;
    log('[IMMEDIATE] ✅ Sent to +' + entry.to);
  } catch (err) {
    log('[IMMEDIATE] ❌ Failed to +' + entry.to + ': ' + err.message);
  }
}

// ─── Text Formatters ─────────────────────────────────────────────────────

const LINE = '══════════════════════════════';
const DIV  = '──────────────────────────────';
const FOOTER = '\n' + LINE + '\n🏢 *' + BOT_NAME + '*\n🌐 ' + SITE_URL + '\n📱 Admin: +' + ADMIN_NUMBER;

function formatItemsText(data) {
  let msg = '📦 ✦ *𝗜𝗧𝗘𝗠𝗦 𝗥𝗘𝗣𝗢𝗥𝗧* ✦ 📦\n' + LINE + '\n';
  msg += '📊 *Total Items:* ' + data.length + '\n' + LINE + '\n\n';
  data.forEach((e, i) => {
    const statusIcon = e.status === 'available' ? '🟢' : e.status === 'in-use' ? '🔵' : '🔴';
    msg += (i+1) + '. 📦 *' + (e.name || 'N/A') + '*\n';
    msg += '   🔢 Serial: ' + (e.number || 'N/A') + '\n';
    msg += '   👤 Person: ' + (e.person || 'N/A') + '\n';
    msg += '   📐 Model: ' + (e.model || 'N/A') + '\n';
    msg += '   📦 Qty: ' + (e.quantity || 1) + '\n';
    msg += '   ' + statusIcon + ' Status: ' + (e.status || 'available') + '\n';
    if (i < data.length - 1) msg += DIV + '\n';
  });
  return msg + FOOTER;
}

function formatWalletText(data) {
  const totalIn = data.filter(e => e.type === 'in').reduce((a, e) => a + Number(e.amount || 0), 0);
  const totalOut = data.filter(e => e.type === 'out').reduce((a, e) => a + Number(e.amount || 0), 0);
  const balance = totalIn - totalOut;
  const balEmoji = balance >= 0 ? '✅' : '⚠️';
  let msg = '💰 ✦ *𝗪𝗔𝗟𝗟𝗘𝗧 𝗥𝗘𝗣𝗢𝗥𝗧* ✦ 💰\n' + LINE + '\n';
  msg += '📥 *Total Received:* Rs. ' + totalIn.toLocaleString() + '\n';
  msg += '📤 *Total Spent:* Rs. ' + totalOut.toLocaleString() + '\n';
  msg += '🏦 *Balance:* ' + balEmoji + ' Rs. ' + balance.toLocaleString() + '\n';
  msg += '📊 *Entries:* ' + data.length + '\n' + LINE + '\n\n';
  data.forEach((e, i) => {
    const icon = e.type === 'in' ? '📥' : '📤';
    const date = e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB') : '';
    msg += icon + ' *' + (e.type || '').toUpperCase() + '* — Rs. ' + (Number(e.amount)||0).toLocaleString() + '\n';
    msg += '  👤 ' + (e.personOrPurpose || 'N/A') + '\n  📅 ' + date + '\n';
    if (i < data.length - 1) msg += DIV + '\n';
  });
  return msg + FOOTER;
}

function formatPersonText(data) {
  let msg = '👷 ✦ *𝗣𝗘𝗥𝗦𝗢𝗡 𝗔𝗧𝗧𝗘𝗡𝗗𝗔𝗡𝗖𝗘* ✦ 👷\n' + LINE + '\n';
  msg += '📊 *Total Entries:* ' + data.length + '\n' + LINE + '\n\n';
  data.forEach((e, i) => {
    const icon = e.action === 'enter' ? '🟢' : '🔴';
    const date = e.timestamp ? new Date(e.timestamp).toLocaleString('en-GB') : '';
    msg += (i+1) + '. 👤 *' + (e.personName || 'N/A') + '*\n';
    msg += '   ' + icon + ' ' + (e.action === 'enter' ? 'Checked In' : 'Checked Out') + '\n';
    msg += '   📅 ' + date + '\n';
    if (i < data.length - 1) msg += DIV + '\n';
  });
  return msg + FOOTER;
}

function formatMaintenanceText(data) {
  const open = data.filter(e => e.status !== 'solved').length;
  let msg = '🔧 ✦ *𝗠𝗔𝗜𝗡𝗧𝗘𝗡𝗔𝗡𝗖𝗘 𝗥𝗘𝗣𝗢𝗥𝗧* ✦ 🔧\n' + LINE + '\n';
  msg += '📊 *Total:* ' + data.length + ' | 🔴 *Open:* ' + open + ' | ✅ *Solved:* ' + (data.length - open) + '\n' + LINE + '\n\n';
  data.forEach((e, i) => {
    const statusIcon = e.status === 'solved' ? '✅' : '🔴';
    msg += (i+1) + '. ' + statusIcon + ' *' + (e.category || 'Issue') + '*\n';
    msg += '   📝 ' + (e.subject || 'N/A') + '\n';
    msg += '   📄 ' + (e.description || 'N/A') + '\n';
    if (i < data.length - 1) msg += DIV + '\n';
  });
  return msg + FOOTER;
}

function formatSamplesText(data) {
  let msg = '🧪 ✦ *𝗦𝗔𝗠𝗣𝗟𝗘𝗦 𝗥𝗘𝗣𝗢𝗥𝗧* ✦ 🧪\n' + LINE + '\n';
  msg += '📊 *Total:* ' + data.length + '\n' + LINE + '\n\n';
  data.forEach((e, i) => {
    const icon = e.type === 'in' ? '📥' : '📤';
    msg += (i+1) + '. ' + icon + ' *' + (e.type === 'in' ? 'Sample In' : 'Sample Out') + '*\n';
    msg += '   👤 ' + (e.personName || 'N/A') + '\n';
    msg += '   📋 ' + (e.program || 'N/A') + '\n';
    msg += '   🔢 ' + (e.pieces || 'N/A') + ' pieces\n';
    if (i < data.length - 1) msg += DIV + '\n';
  });
  return msg + FOOTER;
}

function formatClippingText(data) {
  const inEntries = data.filter(e => e.type === 'in');
  let totalSize = 0;
  inEntries.forEach(e => { const n = parseFloat(e.size); if (!isNaN(n)) totalSize += n; });
  let msg = '✂️ ✦ *𝗖𝗟𝗜𝗣𝗣𝗜𝗡𝗚 𝗥𝗘𝗣𝗢𝗥𝗧* ✂️\n' + LINE + '\n';
  msg += '📊 *Total:* ' + data.length + ' | 📏 *Total Size:* ' + totalSize + ' yards\n' + LINE + '\n\n';
  data.forEach((e, i) => {
    const icon = e.type === 'in' ? '📥' : e.type === 'transfer' ? '💸' : '📤';
    msg += (i+1) + '. ' + icon + ' *' + (e.clipperName || 'N/A') + '*\n';
    msg += '   📏 ' + (e.size || 'N/A') + ' yards\n';
    msg += '   🏷️ ' + (e.type || 'N/A') + '\n';
    if (i < data.length - 1) msg += DIV + '\n';
  });
  return msg + FOOTER;
}

// ─── Stock Manager Helpers ───────────────────────────────────────────────────

async function fetchStockData(section) {
  try {
    const res = await axios.get(`${GITHUB_RAW}/${section}.json`, { timeout: 10000 });
    return Array.isArray(res.data) ? res.data : [];
  } catch (err) {
    console.error('[STOCK] Fetch ' + section + ' error:', err.message);
    return [];
  }
}

async function addExcelFooter(ws) {
  ws.addRow({});
  const footerRow1 = ws.addRow({ no: '', name: BOT_NAME });
  footerRow1.font = { bold: true, size: 11, color: { argb: 'FF25D366' } };
  const footerRow2 = ws.addRow({ no: '', name: SITE_URL });
  footerRow2.font = { color: { argb: 'FF0066CC' }, underline: true };
  const footerRow3 = ws.addRow({ no: '', name: 'Admin: +' + ADMIN_NUMBER });
  footerRow3.font = { italic: true, color: { argb: 'FF666666' } };
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
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF25D366' } };
  data.forEach((e, i) => {
    ws.addRow({
      no: i + 1,
      name: e.name || '',
      number: e.number || '',
      model: e.model || '',
      quantity: e.quantity || 1,
      person: e.person || '',
      status: e.status || 'available',
      timestamp: e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB') : ''
    });
  });
  await addExcelFooter(ws);
  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
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
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF25D366' } };
  let balance = 0;
  data.forEach((e, i) => {
    const amt = Number(e.amount) || 0;
    if (e.type === 'in') balance += amt;
    else balance -= amt;
    ws.addRow({
      no: i + 1,
      type: (e.type || '').toUpperCase(),
      personOrPurpose: e.personOrPurpose || '',
      amount: amt,
      timestamp: e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB') : ''
    });
  });
  // Summary row
  const totalIn = data.filter(e => e.type === 'in').reduce((a, e) => a + Number(e.amount || 0), 0);
  const totalOut = data.filter(e => e.type === 'out').reduce((a, e) => a + Number(e.amount || 0), 0);
  ws.addRow({});
  const sumRow1 = ws.addRow({ type: '', personOrPurpose: 'TOTAL RECEIVED', amount: totalIn });
  sumRow1.font = { bold: true, color: { argb: 'FF00AA00' } };
  const sumRow2 = ws.addRow({ type: '', personOrPurpose: 'TOTAL SPENT', amount: totalOut });
  sumRow2.font = { bold: true, color: { argb: 'FFCC0000' } };
  const sumRow3 = ws.addRow({ type: '', personOrPurpose: 'BALANCE', amount: balance });
  sumRow3.font = { bold: true, size: 14 };
  await addExcelFooter(ws);
  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
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
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF25D366' } };
  data.forEach((e, i) => {
    ws.addRow({
      no: i + 1,
      personName: e.personName || '',
      action: e.action || '',
      timestamp: e.timestamp ? new Date(e.timestamp).toLocaleString('en-GB') : ''
    });
  });
  await addExcelFooter(ws);
  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

async function generateGenericXlsx(section, data) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(section.charAt(0).toUpperCase() + section.slice(1));
  if (!data.length) return null;
  // Get all keys from first entry
  const keys = Object.keys(data[0]).filter(k => k !== 'id');
  ws.columns = [
    { header: '#', key: 'no', width: 5 },
    ...keys.map(k => ({ header: k.charAt(0).toUpperCase() + k.slice(1), key: k, width: 20 }))
  ];
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF25D366' } };
  data.forEach((e, i) => {
    const row = { no: i + 1 };
    keys.forEach(k => { row[k] = e[k] != null ? String(e[k]) : ''; });
    ws.addRow(row);
  });
  await addExcelFooter(ws);
  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

// ─── Session Restore from SESSION_ID ─────────────────────────────────────────

function restoreSession() {
  const sessionId = process.env.SESSION_ID;
  if (!sessionId) {
    console.log('[SESSION] No SESSION_ID env var. Will generate QR for pairing.');
    return false;
  }

  try {
    if (!fs.existsSync(SESSION_DIR)) {
      fs.mkdirSync(SESSION_DIR, { recursive: true });
    }

    // Check if session already has valid creds
    const credsFile = path.join(SESSION_DIR, 'creds.json');
    if (fs.existsSync(credsFile)) {
      const creds = JSON.parse(fs.readFileSync(credsFile, 'utf8'));
      if (creds.me) {
        console.log('[SESSION] Existing valid session found for ' + (creds.me.id || 'unknown'));
        return true;
      }
    }

    // Try to decode SESSION_ID (format: PREFIX~base64data)
    console.log('[SESSION] Restoring session from SESSION_ID...');
    const base64Data = sessionId.includes('~') ? sessionId.split('~').slice(1).join('~') : sessionId;
    const buffer = Buffer.from(base64Data, 'base64');

    // Try as ZIP first (contains session folder structure)
    try {
      const AdmZip = require('adm-zip');
      const zip = new AdmZip(buffer);
      const entries = zip.getEntries();

      if (entries.length > 0) {
        zip.extractAllTo(SESSION_DIR, true);
        console.log('[SESSION] Extracted ZIP with ' + entries.length + ' files');
        return true;
      }
    } catch (zipErr) {
      // Not a ZIP, try as JSON
    }

    // Try as JSON (creds data)
    try {
      const jsonStr = buffer.toString('utf8');
      const json = JSON.parse(jsonStr);
      if (json.creds || json.me || json.account) {
        // Write as creds.json
        fs.writeFileSync(credsFile, JSON.stringify(json, null, 2), 'utf8');
        console.log('[SESSION] Restored creds from SESSION_ID JSON');
        return true;
      }
    } catch (jsonErr) {
      // Not JSON either
    }

    // Try writing raw decoded content as creds.json
    try {
      const jsonStr = buffer.toString('utf8');
      const json = JSON.parse(jsonStr);
      fs.writeFileSync(credsFile, JSON.stringify(json, null, 2), 'utf8');
      console.log('[SESSION] Restored raw SESSION_ID as creds.json');
      return true;
    } catch (_) {}

    console.log('[SESSION] Could not decode SESSION_ID format. Will generate QR.');
    return false;

  } catch (err) {
    console.error('[SESSION] Restore error:', err.message);
    return false;
  }
}

// ─── Process Message Queue ───────────────────────────────────────────────────

async function processOutbox(sock) {
  const outbox = readOutbox();
  const pending = outbox.filter(m => !m.sent);

  if (pending.length === 0) return { processed: 0, failed: 0 };

  let processed = 0;
  let failed = 0;

  for (const entry of pending) {
    try {
      const target = entry.to.replace(/[^0-9]/g, '') + '@s.whatsapp.net';
      await sock.sendMessage(target, { text: entry.message });
      entry.sent = true;
      entry.sentAt = new Date().toISOString();
      entry.error = null;
      processed++;
      console.log('[OUTBOX] Sent to +' + entry.to);
    } catch (err) {
      entry.error = err.message;
      failed++;
      console.error('[OUTBOX] Failed to +' + entry.to + ': ' + err.message);
    }
  }

  // Clean old sent messages (keep 1 hour)
  const cleaned = outbox.filter(m => {
    if (!m.sent) return true;
    return (Date.now() - new Date(m.sentAt).getTime()) < 3600000;
  });
  writeOutbox(cleaned.length > 200 ? cleaned.slice(-200) : cleaned);

  return { processed, failed };
}

let loggedOutOnce = false;

// ─── Main Bot ────────────────────────────────────────────────────────────────

async function startBot() {
  // Try to restore session from SESSION_ID (only if not previously logged out)
  if (!loggedOutOnce) {
    restoreSession();
  }

  // Ensure session dir exists
  if (!fs.existsSync(SESSION_DIR)) {
    fs.mkdirSync(SESSION_DIR, { recursive: true });
  }

  const { state, saveCreds } = await useMultiFileAuthState(SESSION_DIR);

  // Try to get latest version, fallback to default
  let waVersion = [2, 2413, 1];
  try {
    const baileys = require('@whiskeysockets/baileys');
    if (typeof baileys.fetchLatestBaileysVersion === 'function') {
      const { version } = await baileys.fetchLatestBaileysVersion();
      waVersion = version;
      console.log('[BOT] Using Baileys WA version: ' + waVersion.join('.'));
    }
  } catch (_) {}

  const sock = makeWASocket({
    version: waVersion,
    auth: state,
    logger: pino({ level: 'silent' }),
    printQRInTerminal: true,
    browser: [BOT_NAME, 'Chrome', '1.0.0'],
    generateHighQualityLinkPreview: false,
  });

  // Save creds on update
  sock.ev.on('creds.update', saveCreds);

  // Connection handler
  sock.ev.on('connection.update', async (update) => {
    const { connection, lastDisconnect, qr } = update;

    if (qr) {
      console.log('[QR] New QR code generated. Scan in WhatsApp > Linked Devices.');
    }

    if (connection === 'open') {
      whatsappConnected = true;
      botRunning = true;
      sockRef = sock;
      log('═══════════════════════════════════════════');
      log('  WhatsApp Bot CONNECTED');
      log('  Account: ' + (sock.user?.id || 'unknown'));
      log('═══════════════════════════════════════════');

      // Process outbox immediately on connect
      const result = await processOutbox(sock);
      if (result.processed > 0) {
        log('[OUTBOX] Initial send: ' + result.processed + ' sent, ' + result.failed + ' failed');
      }

      // Auto-process every 30 seconds
      setInterval(async () => {
        try {
          const r = await processOutbox(sock);
          if (r.processed > 0) {
            log('[OUTBOX] Auto-send: ' + r.processed + ' sent');
          }
        } catch (err) {
          log('[OUTBOX] Auto-send error: ' + err.message);
        }
      }, 30000);
    }

    if (connection === 'close') {
      whatsappConnected = false;
      sockRef = null;
      const reason = lastDisconnect?.error?.output?.statusCode;
      log('[CONN] Connection closed. Reason code: ' + reason);

      if (reason === DisconnectReason.loggedOut) {
        console.log('[CONN] Session LOGGED OUT (401).');
        // Clear session dir
        try {
          const files = fs.readdirSync(SESSION_DIR);
          for (const f of files) {
            fs.unlinkSync(path.join(SESSION_DIR, f));
          }
        } catch (_) {}

        if (!loggedOutOnce) {
          loggedOutOnce = true;
          console.log('[CONN] SESSION_ID is invalid. Generating new QR for pairing...');
          console.log('[CONN] Scan the QR code in WhatsApp > Linked Devices > Link a Device');
          // Restart once to generate fresh QR (without trying old SESSION_ID)
          setTimeout(startBot, 2000);
        } else {
          console.log('[CONN] Waiting 60s before retry (scan QR code from Railway logs)...');
          setTimeout(startBot, 60000);
        }
      } else {
        console.log('[CONN] Reconnecting in 5s...');
        setTimeout(startBot, 5000);
      }
    }
  });

  // Message handler
  sock.ev.on('messages.upsert', async ({ messages }) => {
    const msg = messages[0];
    if (!msg.message) return;

    const chatId = msg.key.remoteJid;
    const text = msg.message.conversation
      || msg.message.extendedTextMessage?.text
      || '';

    if (!text.startsWith('.')) return;

    const parts = text.trim().split(/\s+/);
    const cmd = parts[0].toLowerCase();

    try {
      // .ping
      if (cmd === '.ping') {
        const start = Date.now();
        await sock.sendMessage(chatId, { text: 'Pong!' });
        const latency = Date.now() - start;
        await sock.sendMessage(chatId, {
          text: '*' + BOT_NAME + ' Bot*\nLatency: ' + latency + 'ms\nStatus: Online'
        });
      }

      // .sendoutbox
      if (cmd === '.sendoutbox') {
        const sub = (parts[1] || '').toLowerCase();

        if (sub === 'status') {
          const outbox = readOutbox();
          const pending = outbox.filter(m => !m.sent).length;
          const sent = outbox.filter(m => m.sent).length;
          await sock.sendMessage(chatId, {
            text: '*QUEUE STATUS*\nPending: ' + pending + '\nSent: ' + sent + '\nAuto-send: every 30s'
          });
          return;
        }

        const result = await processOutbox(sock);
        const outbox = readOutbox();
        const pending = outbox.filter(m => !m.sent).length;
        await sock.sendMessage(chatId, {
          text: '*QUEUE PROCESSED*\nSent: ' + result.processed + '\nFailed: ' + result.failed + '\nRemaining: ' + pending
        });
      }

      // .status
      if (cmd === '.status') {
        const outbox = readOutbox();
        const pending = outbox.filter(m => !m.sent).length;
        await sock.sendMessage(chatId, {
          text: '*' + BOT_NAME + '*\nBot: Online\nAccount: ' + (sock.user?.id || 'unknown') + '\nPending messages: ' + pending + '\nUptime: ' + Math.floor(process.uptime()) + 's'
        });
      }

      // ─── Stock Manager Commands ────────────────────────────────────────────

      // .items - send items.xlsx
      if (cmd === '.items') {
        await sock.sendMessage(chatId, { text: '📦 Fetching items data...' });
        const data = await fetchStockData('items');
        if (!data.length) {
          await sock.sendMessage(chatId, { text: '❌ No items data found.' });
          return;
        }
        const buf = await generateItemsXlsx(data);
        const tmpFile = path.join(__dirname, 'tmp', 'items_' + Date.now() + '.xlsx');
        if (!fs.existsSync(path.join(__dirname, 'tmp'))) fs.mkdirSync(path.join(__dirname, 'tmp'), { recursive: true });
        fs.writeFileSync(tmpFile, buf);
        await sock.sendMessage(chatId, {
          document: fs.readFileSync(tmpFile),
          fileName: 'items.xlsx',
          mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          caption: '📦 *ITEMS* (' + data.length + ' entries)\n📅 ' + new Date().toLocaleDateString('en-GB') + '\n🏢 ' + BOT_NAME
        });
        fs.unlinkSync(tmpFile);
      }

      // .items2 - send items as formatted text
      if (cmd === '.items2') {
        await sock.sendMessage(chatId, { text: '📦 Fetching items text...' });
        const data = await fetchStockData('items');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No items data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatItemsText(data) });
      }

      // .wallet - send wallet.xlsx
      if (cmd === '.wallet') {
        await sock.sendMessage(chatId, { text: '💰 Fetching wallet data...' });
        const data = await fetchStockData('wallet');
        if (!data.length) {
          await sock.sendMessage(chatId, { text: '❌ No wallet data found.' });
          return;
        }
        const buf = await generateWalletXlsx(data);
        const tmpFile = path.join(__dirname, 'tmp', 'wallet_' + Date.now() + '.xlsx');
        if (!fs.existsSync(path.join(__dirname, 'tmp'))) fs.mkdirSync(path.join(__dirname, 'tmp'), { recursive: true });
        fs.writeFileSync(tmpFile, buf);
        const totalIn = data.filter(e => e.type === 'in').reduce((a, e) => a + Number(e.amount || 0), 0);
        const totalOut = data.filter(e => e.type === 'out').reduce((a, e) => a + Number(e.amount || 0), 0);
        const balance = totalIn - totalOut;
        await sock.sendMessage(chatId, {
          document: fs.readFileSync(tmpFile),
          fileName: 'wallet.xlsx',
          mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          caption: '💰 *WALLET*\n📥 Received: Rs. ' + totalIn.toLocaleString() + '\n📤 Spent: Rs. ' + totalOut.toLocaleString() + '\n🏦 Balance: Rs. ' + balance.toLocaleString() + '\n🏢 ' + BOT_NAME
        });
        fs.unlinkSync(tmpFile);
      }

      // .wallet2 / .wallettxt - send wallet as formatted text
      if (cmd === '.wallet2' || cmd === '.wallettxt') {
        const data = await fetchStockData('wallet');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No wallet data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatWalletText(data) });
      }

      // .person - send person.xlsx
      if (cmd === '.person') {
        await sock.sendMessage(chatId, { text: '👷 Fetching person data...' });
        const data = await fetchStockData('person');
        if (!data.length) {
          await sock.sendMessage(chatId, { text: '❌ No person data found.' });
          return;
        }
        const buf = await generatePersonXlsx(data);
        const tmpFile = path.join(__dirname, 'tmp', 'person_' + Date.now() + '.xlsx');
        if (!fs.existsSync(path.join(__dirname, 'tmp'))) fs.mkdirSync(path.join(__dirname, 'tmp'), { recursive: true });
        fs.writeFileSync(tmpFile, buf);
        await sock.sendMessage(chatId, {
          document: fs.readFileSync(tmpFile),
          fileName: 'person.xlsx',
          mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          caption: '👷 *PERSON ATTENDANCE* (' + data.length + ' entries)\n🏢 ' + BOT_NAME
        });
        fs.unlinkSync(tmpFile);
      }

      // .person2 - send person as formatted text
      if (cmd === '.person2') {
        await sock.sendMessage(chatId, { text: '👷 Fetching person text...' });
        const data = await fetchStockData('person');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No person data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatPersonText(data) });
      }

      // .maintenance - send maintenance.xlsx
      if (cmd === '.maintenance') {
        await sock.sendMessage(chatId, { text: '🔧 Fetching maintenance data...' });
        const data = await fetchStockData('maintenance');
        if (!data.length) {
          await sock.sendMessage(chatId, { text: '❌ No maintenance data found.' });
          return;
        }
        const buf = await generateGenericXlsx('maintenance', data);
        const tmpFile = path.join(__dirname, 'tmp', 'maintenance_' + Date.now() + '.xlsx');
        if (!fs.existsSync(path.join(__dirname, 'tmp'))) fs.mkdirSync(path.join(__dirname, 'tmp'), { recursive: true });
        fs.writeFileSync(tmpFile, buf);
        const open = data.filter(e => e.status !== 'solved').length;
        await sock.sendMessage(chatId, {
          document: fs.readFileSync(tmpFile),
          fileName: 'maintenance.xlsx',
          mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          caption: '🔧 *MAINTENANCE* (' + data.length + ' issues | 🔴 ' + open + ' open | ✅ ' + (data.length - open) + ' solved)\n🏢 ' + BOT_NAME
        });
        fs.unlinkSync(tmpFile);
      }

      // .maintenance2 - send maintenance as formatted text
      if (cmd === '.maintenance2') {
        await sock.sendMessage(chatId, { text: '🔧 Fetching maintenance text...' });
        const data = await fetchStockData('maintenance');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No maintenance data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatMaintenanceText(data) });
      }

      // .samples - send samples.xlsx
      if (cmd === '.samples') {
        await sock.sendMessage(chatId, { text: '🧪 Fetching samples data...' });
        const data = await fetchStockData('samples');
        if (!data.length) {
          await sock.sendMessage(chatId, { text: '❌ No samples data found.' });
          return;
        }
        const buf = await generateGenericXlsx('samples', data);
        const tmpFile = path.join(__dirname, 'tmp', 'samples_' + Date.now() + '.xlsx');
        if (!fs.existsSync(path.join(__dirname, 'tmp'))) fs.mkdirSync(path.join(__dirname, 'tmp'), { recursive: true });
        fs.writeFileSync(tmpFile, buf);
        await sock.sendMessage(chatId, {
          document: fs.readFileSync(tmpFile),
          fileName: 'samples.xlsx',
          mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          caption: '🧪 *SAMPLES* (' + data.length + ' entries)\n🏢 ' + BOT_NAME
        });
        fs.unlinkSync(tmpFile);
      }

      // .samples2 - send samples as formatted text
      if (cmd === '.samples2') {
        await sock.sendMessage(chatId, { text: '🧪 Fetching samples text...' });
        const data = await fetchStockData('samples');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No samples data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatSamplesText(data) });
      }

      // .clipping - send clipping.xlsx
      if (cmd === '.clipping') {
        await sock.sendMessage(chatId, { text: '✂️ Fetching clipping data...' });
        const data = await fetchStockData('clipping');
        if (!data.length) {
          await sock.sendMessage(chatId, { text: '❌ No clipping data found.' });
          return;
        }
        const buf = await generateGenericXlsx('clipping', data);
        const tmpFile = path.join(__dirname, 'tmp', 'clipping_' + Date.now() + '.xlsx');
        if (!fs.existsSync(path.join(__dirname, 'tmp'))) fs.mkdirSync(path.join(__dirname, 'tmp'), { recursive: true });
        fs.writeFileSync(tmpFile, buf);
        const inEntries = data.filter(e => e.type === 'in');
        let totalSize = 0;
        inEntries.forEach(e => { const n = parseFloat(e.size); if (!isNaN(n)) totalSize += n; });
        const totalPayment = totalSize * 12;
        await sock.sendMessage(chatId, {
          document: fs.readFileSync(tmpFile),
          fileName: 'clipping.xlsx',
          mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          caption: '✂️ *CLIPPING* (' + data.length + ' entries | ' + totalSize + ' yards | 💰 Rs. ' + totalPayment.toLocaleString() + ')\n🏢 ' + BOT_NAME
        });
        fs.unlinkSync(tmpFile);
      }

      // .clipping2 - send clipping as formatted text
      if (cmd === '.clipping2') {
        await sock.sendMessage(chatId, { text: '✂️ Fetching clipping text...' });
        const data = await fetchStockData('clipping');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No clipping data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatClippingText(data) });
      }

      // .stock / .help - show all commands
      if (cmd === '.stock' || cmd === '.help') {
        const helpMsg = [
          '🤖 *' + BOT_NAME + '*',
          LINE,
          '📦 *.items* — Items Excel file',
          '📝 *.items2* — Items text details',
          '💰 *.wallet* — Wallet Excel file',
          '📝 *.wallet2* — Wallet text details',
          '👷 *.person* — Attendance Excel',
          '📝 *.person2* — Attendance text',
          '🔧 *.maintenance* — Maintenance Excel',
          '📝 *.maintenance2* — Maintenance text',
          '🧪 *.samples* — Samples Excel',
          '📝 *.samples2* — Samples text',
          '✂️ *.clipping* — Clipping Excel',
          '📝 *.clipping2* — Clipping text',
          DIV,
          '📌 *.ping* — Check latency',
          '📌 *.status* — Bot status',
          '📌 *.sendoutbox* — Process queue',
          LINE,
          '🌐 ' + SITE_URL,
          '📱 Admin: +' + ADMIN_NUMBER
        ].join('\n');
        await sock.sendMessage(chatId, { text: helpMsg });
      }

    } catch (err) {
      console.error('[CMD] Error handling ' + cmd + ':', err.message);
      try {
        await sock.sendMessage(chatId, { text: 'Error: ' + err.message });
      } catch (_) {}
    }
  });
}

// ─── Express API (starts FIRST, then bot) ──────────────────────────────────

const app = express();
app.use(express.json());

app.get('/', (req, res) => {
  res.json({
    bot: BOT_NAME,
    online: true,
    whatsappConnected,
    uptime: Math.floor(process.uptime()) + 's'
  });
});

app.get('/status', (req, res) => {
  const outbox = readOutbox();
  res.json({
    bot: BOT_NAME,
    online: true,
    botRunning,
    whatsappConnected,
    pendingMessages: outbox.filter(m => !m.sent).length,
    sentMessages: sentCount,
    adminNumber: ADMIN_NUMBER,
    uptime: Math.floor(process.uptime()) + 's',
    recentLogs: botLogs.slice(-10)
  });
});

app.post('/send', authMiddleware, (req, res) => {
  const { message, to } = req.body;
  if (!message) return res.status(400).json({ error: 'message required' });
  const target = to || ADMIN_NUMBER;
  const entry = {
    id: 'msg_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8),
    to: target,
    message,
    queuedAt: new Date().toISOString(),
    sent: false
  };
  const outbox = readOutbox();
  outbox.push(entry);
  writeOutbox(outbox);
  log('[API] Queued message for +' + target);
  if (whatsappConnected && sockRef) {
    processImmediateSend(sockRef, entry).catch(err => {
      log('[API] Immediate send failed: ' + err.message);
    });
  }
  res.json({ success: true, message: 'Message queued for +' + target, id: entry.id });
});

const PORT_START = PORT;
app.listen(PORT_START, () => {
  log('[SERVER] ✅ API listening on port ' + PORT_START);
  startBot().catch(err => { log('[FATAL] Bot failed: ' + err.message); });
});

// ─── Prevent crashes ───────────────────────────────────────────────────────
process.on('uncaughtException', (err) => {
  log('[BOT] Uncaught: ' + err.message);
});
process.on('unhandledRejection', (reason) => {
  log('[BOT] Unhandled rejection: ' + reason);
});
