/**
 * BANU SAEED HOSPITAL - Clean WhatsApp Bot
 * Replaces the broken obfuscated bot with a clean Baileys-based implementation.
 * 
 * Features:
 * - Multi-file auth state (session/ folder)
 * - SESSION_ID restoration from env (base64 encoded)
 * - QR code pairing for new sessions
 * - .sendoutbox command to process message queue
 * - Auto-send queue every 30 seconds
 * - .ping / .status commands
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

const PORT = process.env.PORT || 3001;
const API_SECRET = process.env.API_SECRET || 'banu-saeed-secret-2024';
const ADMIN_NUMBER = process.env.ADMIN_NUMBER || '923299931199';
const SESSION_DIR = path.join(__dirname, 'session');
const OUTBOX_PATH = path.join(__dirname, 'data', 'outbox.json');
const GITHUB_RAW = 'https://raw.githubusercontent.com/hasilpurofficial4-creator/zaid2/main/data';

// ─── State ──────────────────────────────────────────────────────────────────

let whatsappConnected = false;
let sockRef = null;
let botLogs = [];

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
    browser: ['Banu Saeed Hospital', 'Chrome', '1.0.0'],
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
      sockRef = sock;
      console.log('═══════════════════════════════════════════');
      console.log('  WhatsApp Bot CONNECTED');
      console.log('  Account: ' + (sock.user?.id || 'unknown'));
      console.log('═══════════════════════════════════════════');

      // Process outbox immediately on connect
      const result = await processOutbox(sock);
      if (result.processed > 0) {
        console.log('[OUTBOX] Initial send: ' + result.processed + ' sent, ' + result.failed + ' failed');
      }

      // Auto-process every 30 seconds
      setInterval(async () => {
        try {
          const r = await processOutbox(sock);
          if (r.processed > 0) {
            console.log('[OUTBOX] Auto-send: ' + r.processed + ' sent');
            botLogs.push('[OUTBOX] Auto-send: ' + r.processed + ' sent');
          }
        } catch (err) {
          console.error('[OUTBOX] Auto-send error:', err.message);
        }
      }, 30000);
    }

    if (connection === 'close') {
      whatsappConnected = false;
      const reason = lastDisconnect?.error?.output?.statusCode;
      console.log('[CONN] Connection closed. Reason code: ' + reason);

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
          text: '*BANU SAEED HOSPITAL Bot*\nLatency: ' + latency + 'ms\nStatus: Online'
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
          text: '*BANU SAEED HOSPITAL*\nBot: Online\nAccount: ' + (sock.user?.id || 'unknown') + '\nPending messages: ' + pending + '\nUptime: ' + Math.floor(process.uptime()) + 's'
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
          caption: '📦 *ITEMS* (' + data.length + ' entries)\n📅 ' + new Date().toLocaleDateString('en-GB')
        });
        fs.unlinkSync(tmpFile);
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
          caption: '💰 *WALLET*\n📥 Received: Rs. ' + totalIn.toLocaleString() + '\n📤 Spent: Rs. ' + totalOut.toLocaleString() + '\n🏦 Balance: Rs. ' + balance.toLocaleString()
        });
        fs.unlinkSync(tmpFile);
      }

      // .wallettxt - send all wallet details as formatted text
      if (cmd === '.wallettxt') {
        const data = await fetchStockData('wallet');
        if (!data.length) {
          await sock.sendMessage(chatId, { text: '❌ No wallet data found.' });
          return;
        }
        const totalIn = data.filter(e => e.type === 'in').reduce((a, e) => a + Number(e.amount || 0), 0);
        const totalOut = data.filter(e => e.type === 'out').reduce((a, e) => a + Number(e.amount || 0), 0);
        const balance = totalIn - totalOut;
        const balEmoji = balance >= 0 ? '✅' : '⚠️';
        const line = '══════════════════════════════';
        const div = '──────────────────────────────';
        let msg = '💰 ✦ *𝗪𝗔𝗟𝗟𝗘𝗧 𝗥𝗘𝗣𝗢𝗥𝗧* ✦ 💰\n' + line + '\n';
        msg += '📥 *𝗧𝗼𝗧𝗮𝗹 𝗥𝗲𝗰𝗲𝗶𝘃𝗲𝗱:* Rs. ' + totalIn.toLocaleString() + '\n';
        msg += '📤 *𝗧𝗼𝗧𝗮𝗹 𝗦𝗽𝗲𝗻𝘁:* Rs. ' + totalOut.toLocaleString() + '\n';
        msg += '🏦 *𝗕𝗮𝗹𝗮𝗻𝗰𝗲:* ' + balEmoji + ' Rs. ' + balance.toLocaleString() + '\n';
        msg += '📊 *𝗘𝗻𝘁𝗿𝗶𝗲𝘀:* ' + data.length + '\n' + line + '\n\n';
        data.forEach((e, i) => {
          const amt = Number(e.amount) || 0;
          const icon = e.type === 'in' ? '📥' : '📤';
          const date = e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB') : '';
          msg += icon + ' *' + (e.type || '').toUpperCase() + '* — Rs. ' + amt.toLocaleString() + '\n';
          msg += '  👤 ' + (e.personOrPurpose || 'N/A') + '\n';
          msg += '  📅 ' + date + '\n';
          if (i < data.length - 1) msg += div + '\n';
        });
        msg += '\n' + line + '\n👨‍💻 *𝗭𝗔𝗜𝗗 𝗕𝗪𝗣 𝗗𝗘𝗩𝗘𝗟𝗢𝗣𝗘𝗥*\n🌐 https://zaidbwp.vercel.app';
        await sock.sendMessage(chatId, { text: msg });
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
          caption: '👷 *PERSON ATTENDANCE* (' + data.length + ' entries)'
        });
        fs.unlinkSync(tmpFile);
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
          caption: '🔧 *MAINTENANCE* (' + data.length + ' issues | 🔴 ' + open + ' open | ✅ ' + (data.length - open) + ' solved)'
        });
        fs.unlinkSync(tmpFile);
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
          caption: '🧪 *SAMPLES* (' + data.length + ' entries)'
        });
        fs.unlinkSync(tmpFile);
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
          caption: '✂️ *CLIPPING* (' + data.length + ' entries | ' + totalSize + ' yards | 💰 Rs. ' + totalPayment.toLocaleString() + ')'
        });
        fs.unlinkSync(tmpFile);
      }

      // .stock - show all stock commands help
      if (cmd === '.stock' || cmd === '.help') {
        const helpMsg = [
          '🤖 *𝗭𝗔𝗜𝗗 𝗕𝗪𝗣 𝗦𝗧𝗢𝗖𝗞 𝗕𝗢𝗧*',
          '══════════════════════════════',
          '📦 *.items* — Send items.xlsx',
          '💰 *.wallet* — Send wallet.xlsx',
          '📝 *.wallettxt* — Wallet details (text)',
          '👷 *.person* — Send attendance.xlsx',
          '🔧 *.maintenance* — Send maintenance.xlsx',
          '🧪 *.samples* — Send samples.xlsx',
          '✂️ *.clipping* — Send clipping.xlsx',
          '──────────────────────────────',
          '📌 *.ping* — Check latency',
          '📌 *.status* — Bot status',
          '📌 *.sendoutbox* — Process queue',
          '══════════════════════════════',
          '👨‍💻 _ZAID ASHIQ BWP_',
          '🌐 https://zaidbwp.vercel.app'
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

// ─── Express API Server ─────────────────────────────────────────────────────

const app = express();
app.use(express.json());

// Auth middleware
function authMiddleware(req, res, next) {
  const { secret } = req.body;
  if (secret !== API_SECRET) {
    return res.status(401).json({ success: false, error: 'Unauthorized' });
  }
  next();
}

// Health check
app.get('/', (req, res) => {
  const outbox = readOutbox();
  const pending = outbox.filter(m => !m.sent).length;
  res.json({
    status: 'online',
    service: 'Banu Saeed Hospital WhatsApp API',
    whatsappConnected,
    uptime: Math.floor(process.uptime()) + 's',
    pendingMessages: pending
  });
});

// Status endpoint
app.get('/status', (req, res) => {
  const outbox = readOutbox();
  const pending = outbox.filter(m => !m.sent).length;
  const sent = outbox.filter(m => m.sent).length;
  res.json({
    online: true,
    botRunning: true,
    whatsappConnected,
    pendingMessages: pending,
    sentMessages: sent,
    adminNumber: ADMIN_NUMBER,
    uptime: Math.floor(process.uptime()) + 's',
    recentLogs: botLogs.slice(-10)
  });
});

// Send message (queues for auto-send)
app.post('/send', authMiddleware, (req, res) => {
  try {
    const { message, to } = req.body;
    if (!message) return res.status(400).json({ success: false, error: 'message is required' });

    const target = (to || ADMIN_NUMBER).replace(/[^0-9]/g, '');
    if (target.length < 7) return res.status(400).json({ success: false, error: 'Invalid phone number' });

    const outbox = readOutbox();
    const entry = {
      id: Date.now().toString(36) + Math.random().toString(36).substr(2, 5),
      to: target,
      message: message,
      sent: false,
      queuedAt: new Date().toISOString(),
      sentAt: null,
      error: null
    };
    outbox.push(entry);

    const cleaned = outbox.filter(m => !m.sent || (Date.now() - new Date(m.sentAt).getTime() < 3600000));
    writeOutbox(cleaned.length > 200 ? cleaned.slice(-200) : cleaned);

    console.log('[QUEUE] Queued for +' + target + ' (' + entry.id + ')');

    // If WhatsApp is connected, send immediately
    if (whatsappConnected && sockRef) {
      processImmediateSend(sockRef, entry).catch(err => {
        console.error('[QUEUE] Immediate send failed:', err.message);
      });
    }

    res.json({ success: true, message: 'Message queued for +' + target, id: entry.id });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// Pair endpoint
app.post('/pair', authMiddleware, (req, res) => {
  res.json({
    success: true,
    message: 'Check Railway logs for QR code. Scan in WhatsApp > Linked Devices > Link a Device.'
  });
});

// Debug endpoint
app.get('/debug', (req, res) => {
  res.json({
    online: true,
    botRunning: true,
    whatsappConnected,
    uptime: Math.floor(process.uptime()) + 's',
    nodeVersion: process.version,
    env: {
      PORT: process.env.PORT || '3001 (default)',
      SESSION_ID: process.env.SESSION_ID ? 'SET (' + process.env.SESSION_ID.substring(0, 8) + '...)' : 'NOT SET',
      API_SECRET: process.env.API_SECRET ? 'SET' : 'NOT SET',
      ADMIN_NUMBER: process.env.ADMIN_NUMBER || '923299931199 (default)',
    },
    recentLogs: botLogs.slice(-20)
  });
});

// Immediate send helper
async function processImmediateSend(sock, entry) {
  try {
    const target = entry.to.replace(/[^0-9]/g, '') + '@s.whatsapp.net';
    await sock.sendMessage(target, { text: entry.message });
    entry.sent = true;
    entry.sentAt = new Date().toISOString();
    entry.error = null;
    console.log('[SEND] Immediate send to +' + entry.to);
    // Update outbox
    const outbox = readOutbox();
    const idx = outbox.findIndex(m => m.id === entry.id);
    if (idx >= 0) outbox[idx] = entry;
    writeOutbox(outbox);
  } catch (err) {
    console.error('[SEND] Immediate send error:', err.message);
    throw err;
  }
}

// ─── Start Express + Bot ────────────────────────────────────────────────────

console.log('═══════════════════════════════════════════════');
console.log('  BANU SAEED HOSPITAL - WhatsApp Bot');
console.log('  Admin: +' + ADMIN_NUMBER);
console.log('  API Port: ' + PORT);
console.log('═══════════════════════════════════════════════');

// Start Express server FIRST (Railway needs this)
app.listen(PORT, () => {
  console.log('[SERVER] API listening on port ' + PORT);
  // Then start WhatsApp bot
  startBot().catch(err => {
    console.error('[FATAL] Bot failed to start:', err.message);
  });
});

// Prevent crashes
process.on('uncaughtException', (err) => {
  console.error('[BOT] Uncaught:', err.message);
});
process.on('unhandledRejection', (reason) => {
  console.error('[BOT] Unhandled rejection:', reason);
});
