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

// Button support via kango-wa
let sendButtons;
try {
  sendButtons = require('kango-wa').sendButtons;
} catch (_) {
  sendButtons = null;
}

// ─── Config ──────────────────────────────────────────────────────────────────

const PORT = parseInt(process.env.PORT) || 3001;
const API_SECRET = process.env.API_SECRET || 'banu-saeed-secret-2024';
const ADMIN_NUMBERS_RAW = process.env.ADMIN_NUMBER || '923299931199';
const ADMIN_NUMBERS = ADMIN_NUMBERS_RAW.split(',').map(n => n.trim()).filter(Boolean);
const ADMIN_NUMBER = ADMIN_NUMBERS[0]; // Primary admin for display
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

const LINE = '╔══════════════════════════════════╗';
const DIV  = '╠══════════════════════════════════╣';
const END  = '╚══════════════════════════════════╝';
const THIN = '──────────────────────────────────';
const FOOTER = '\n' + END + '\n🏢 *' + BOT_NAME + '*\n🌐 ' + SITE_URL + '\n📱 Admin: +' + ADMIN_NUMBER;

// Unicode Mathematical Bold converter for cool headers
function toBold(str) {
  const map = { 'A':'\ud835\udc00','B':'\ud835\udc01','C':'\ud835\udc02','D':'\ud835\udc03','E':'\ud835\udc04','F':'\ud835\udc05','G':'\ud835\udc06','H':'\ud835\udc07','I':'\ud835\udc08','J':'\ud835\udc09','K':'\ud835\udc0a','L':'\ud835\udc0b','M':'\ud835\udc0c','N':'\ud835\udc0d','O':'\ud835\udc0e','P':'\ud835\udc0f','Q':'\ud835\udc10','R':'\ud835\udc11','S':'\ud835\udc12','T':'\ud835\udc13','U':'\ud835\udc14','V':'\ud835\udc15','W':'\ud835\udc16','X':'\ud835\udc17','Y':'\ud835\udc18','Z':'\ud835\udc19','a':'\ud835\udc1a','b':'\ud835\udc1b','c':'\ud835\udc1c','d':'\ud835\udc1d','e':'\ud835\udc1e','f':'\ud835\udc1f','g':'\ud835\udc20','h':'\ud835\udc21','i':'\ud835\udc22','j':'\ud835\udc23','k':'\ud835\udc24','l':'\ud835\udc25','m':'\ud835\udc26','n':'\ud835\udc27','o':'\ud835\udc28','p':'\ud835\udc29','q':'\ud835\udc2a','r':'\ud835\udc2b','s':'\ud835\udc2c','t':'\ud835\udc2d','u':'\ud835\udc2e','v':'\ud835\udc2f','w':'\ud835\udc30','x':'\ud835\udc31','y':'\ud835\udc32','z':'\ud835\udc33','0':'\ud835\udfce','1':'\ud835\udfcf','2':'\ud835\udfd0','3':'\ud835\udfd1','4':'\ud835\udfd2','5':'\ud835\udfd3','6':'\ud835\udfd4','7':'\ud835\udfd5','8':'\ud835\udfd6','9':'\ud835\udfd7' };
  return str.split('').map(c => map[c] || c).join('');
}

function formatItemsText(data) {
  let msg = '📦 ✦ *' + toBold('ITEMS REPORT') + '* ✦ 📦\n' + LINE + '\n';
  msg += '📊 *' + toBold('Total Items') + ':* ' + data.length + '\n' + DIV + '\n\n';
  data.forEach((e, i) => {
    const statusIcon = e.status === 'available' ? '🟢' : e.status === 'in-use' ? '🔵' : '🔴';
    msg += (i+1) + '. 📦 *' + (e.name || 'N/A') + '*\n';
    msg += '   🔢 Serial: ' + (e.number || 'N/A') + '\n';
    msg += '   👤 Person: ' + (e.person || 'N/A') + '\n';
    msg += '   📐 Model: ' + (e.model || 'N/A') + '\n';
    msg += '   📦 Qty: ' + (e.quantity || 1) + '\n';
    msg += '   ' + statusIcon + ' Status: ' + (e.status || 'available') + '\n';
    if (i < data.length - 1) msg += THIN + '\n';
  });
  return msg + FOOTER;
}

function formatWalletText(data) {
  const totalIn = data.filter(e => e.type === 'in').reduce((a, e) => a + Number(e.amount || 0), 0);
  const totalOut = data.filter(e => e.type === 'out').reduce((a, e) => a + Number(e.amount || 0), 0);
  const balance = totalIn - totalOut;
  const balEmoji = balance >= 0 ? '✅' : '⚠️';
  let msg = '💰 ✦ *' + toBold('WALLET REPORT') + '* ✦ 💰\n' + LINE + '\n';
  msg += '📥 *' + toBold('Total Received') + ':* Rs. ' + totalIn.toLocaleString() + '\n';
  msg += '📤 *' + toBold('Total Spent') + ':* Rs. ' + totalOut.toLocaleString() + '\n';
  msg += '🏦 *' + toBold('Balance') + ':* ' + balEmoji + ' Rs. ' + balance.toLocaleString() + '\n';
  msg += '📊 *Entries:* ' + data.length + '\n' + DIV + '\n\n';
  // Show IN entries first (green), then OUT entries (red)
  const inEntries = data.filter(e => e.type === 'in');
  const outEntries = data.filter(e => e.type !== 'in');
  if (inEntries.length) {
    msg += '📥 *' + toBold('MONEY IN') + '* 📥\n' + THIN + '\n';
    inEntries.forEach((e, i) => {
      const date = e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB') : '';
      msg += '🟢 Rs. ' + (Number(e.amount)||0).toLocaleString() + ' — ' + (e.personOrPurpose || 'N/A') + '\n   📅 ' + date + '\n';
      if (i < inEntries.length - 1) msg += THIN + '\n';
    });
  }
  if (outEntries.length) {
    msg += '\n📤 *' + toBold('MONEY OUT') + '* 📤\n' + THIN + '\n';
    outEntries.forEach((e, i) => {
      const date = e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB') : '';
      msg += '🔴 Rs. ' + (Number(e.amount)||0).toLocaleString() + ' — ' + (e.personOrPurpose || 'N/A') + '\n   📅 ' + date + '\n';
      if (i < outEntries.length - 1) msg += THIN + '\n';
    });
  }
  return msg + FOOTER;
}

function formatPersonText(data) {
  let msg = '👷 ✦ *' + toBold('PERSON ATTENDANCE') + '* ✦ 👷\n' + LINE + '\n';
  msg += '📊 *' + toBold('Total Entries') + ':* ' + data.length + '\n' + DIV + '\n\n';
  data.forEach((e, i) => {
    const icon = e.action === 'enter' ? '🟢' : '🔴';
    const date = e.timestamp ? new Date(e.timestamp).toLocaleString('en-GB') : '';
    msg += (i+1) + '. 👤 *' + (e.personName || 'N/A') + '*\n';
    msg += '   ' + icon + ' ' + (e.action === 'enter' ? 'Checked In' : 'Checked Out') + '\n';
    msg += '   📅 ' + date + '\n';
    if (i < data.length - 1) msg += THIN + '\n';
  });
  return msg + FOOTER;
}

function formatMaintenanceText(data) {
  const open = data.filter(e => e.status !== 'solved').length;
  let msg = '🔧 ✦ *' + toBold('MAINTENANCE REPORT') + '* ✦ 🔧\n' + LINE + '\n';
  msg += '📊 *Total:* ' + data.length + ' | 🔴 *Open:* ' + open + ' | ✅ *Solved:* ' + (data.length - open) + '\n' + DIV + '\n\n';
  data.forEach((e, i) => {
    const statusIcon = e.status === 'solved' ? '✅' : '🔴';
    msg += (i+1) + '. ' + statusIcon + ' *' + (e.category || 'Issue') + '*\n';
    msg += '   📝 ' + (e.subject || 'N/A') + '\n';
    msg += '   📄 ' + (e.description || 'N/A') + '\n';
    if (i < data.length - 1) msg += THIN + '\n';
  });
  return msg + FOOTER;
}

function formatSamplesText(data) {
  let msg = '🧪 ✦ *' + toBold('SAMPLES REPORT') + '* ✦ 🧪\n' + LINE + '\n';
  msg += '📊 *Total:* ' + data.length + '\n' + DIV + '\n\n';
  const inEntries = data.filter(e => e.type === 'in');
  const outEntries = data.filter(e => e.type !== 'in');
  if (inEntries.length) {
    msg += '📥 *' + toBold('SAMPLE IN') + '* 📥\n' + THIN + '\n';
    inEntries.forEach((e, i) => {
      msg += '🟢 ' + (e.personName || 'N/A') + ' | ' + (e.program || 'N/A') + ' | ' + (e.pieces || 'N/A') + ' pcs\n';
      if (i < inEntries.length - 1) msg += THIN + '\n';
    });
  }
  if (outEntries.length) {
    msg += '\n📤 *' + toBold('SAMPLE OUT') + '* 📤\n' + THIN + '\n';
    outEntries.forEach((e, i) => {
      msg += '🔴 ' + (e.personName || 'N/A') + ' | ' + (e.program || 'N/A') + ' | ' + (e.pieces || 'N/A') + ' pcs\n';
      if (i < outEntries.length - 1) msg += THIN + '\n';
    });
  }
  return msg + FOOTER;
}

function formatClippingText(data) {
  const inEntries = data.filter(e => e.type === 'in');
  const outEntries = data.filter(e => e.type !== 'in');
  let totalSize = 0;
  inEntries.forEach(e => { const n = parseFloat(e.size); if (!isNaN(n)) totalSize += n; });
  let msg = '✂️ ✦ *' + toBold('CLIPPING REPORT') + '* ✂️\n' + LINE + '\n';
  msg += '📊 *Total:* ' + data.length + ' | 📏 *Total Size:* ' + totalSize + ' yards\n' + DIV + '\n\n';
  if (inEntries.length) {
    msg += '📥 *' + toBold('CLIPPED IN') + '* 📥\n' + THIN + '\n';
    inEntries.forEach((e, i) => {
      msg += '🟢 ' + (e.clipperName || 'N/A') + ' | ' + (e.size || 'N/A') + ' yds\n';
      if (i < inEntries.length - 1) msg += THIN + '\n';
    });
  }
  if (outEntries.length) {
    msg += '\n📤 *' + toBold('CLIPPING OUT') + '* 📤\n' + THIN + '\n';
    outEntries.forEach((e, i) => {
      const icon = e.type === 'transfer' ? '💸' : '🔴';
      msg += icon + ' ' + (e.clipperName || 'N/A') + ' | ' + (e.size || 'N/A') + ' yds | ' + (e.type || 'N/A') + '\n';
      if (i < outEntries.length - 1) msg += THIN + '\n';
    });
  }
  return msg + FOOTER;
}

// ─── Bills Helpers ──────────────────────────────────────────────────────────

function formatBillsText(bills, date) {
  const total = bills.reduce((a, b) => a + (Number(b.totalAmount) || 0), 0);
  const totalItems = bills.reduce((a, b) => a + (Array.isArray(b.items) ? b.items.length : 0), 0);
  let msg = '🧾 ✦ *' + toBold('BILLS — ' + (date || 'ALL').toUpperCase()) + '* ✦ 🧾\n' + LINE + '\n';
  msg += '📦 *Bills:* ' + bills.length + ' | 📋 *Items:* ' + totalItems + ' | 💰 *Total:* Rs. ' + total.toLocaleString() + '\n' + DIV + '\n\n';
  bills.forEach((b, i) => {
    const items = Array.isArray(b.items) ? b.items : [];
    msg += (i + 1) + '. 🧾 *' + (b.personName || 'N/A') + '* — ' + (b.billPurpose || 'N/A') + '\n';
    msg += '   📅 ' + (b.date || 'N/A') + ' ' + (b.time || '') + '\n';
    items.forEach((it, j) => {
      msg += '   ' + (j + 1) + '. ' + (it.name || '') + ' × ' + (it.quantity || 1) + ' — Rs. ' + Number(it.price || 0).toLocaleString() + '\n';
    });
    msg += '   💰 *Total: Rs. ' + Number(b.totalAmount || 0).toLocaleString() + '*\n';
    if (i < bills.length - 1) msg += '\n' + THIN + '\n';
  });
  return msg + FOOTER;
}

async function generateBillsXlsx(bills, date) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Bills');
  ws.columns = [
    { header: '#', key: 'no', width: 5 },
    { header: 'Person', key: 'personName', width: 20 },
    { header: 'Purpose', key: 'billPurpose', width: 18 },
    { header: 'Items', key: 'itemsCount', width: 8 },
    { header: 'Total (Rs.)', key: 'totalAmount', width: 14 },
    { header: 'Date', key: 'date', width: 12 },
    { header: 'Time', key: 'time', width: 10 },
  ];
  ws.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
  ws.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF10B981' } };
  bills.forEach((b, i) => {
    ws.addRow({
      no: i + 1,
      personName: b.personName || '',
      billPurpose: b.billPurpose || '',
      itemsCount: Array.isArray(b.items) ? b.items.length : 0,
      totalAmount: Number(b.totalAmount) || 0,
      date: b.date || '',
      time: b.time || '',
    });
  });
  const total = bills.reduce((a, b) => a + (Number(b.totalAmount) || 0), 0);
  ws.addRow({});
  const totalRow = ws.addRow({ no: '', personName: 'TOTAL', totalAmount: total });
  totalRow.font = { bold: true, size: 12 };
  await addExcelFooter(ws);
  const buf = await wb.xlsx.writeBuffer();
  return Buffer.from(buf);
}

async function generateBillImage(bills, date) {
  const total = bills.reduce((a, b) => a + (Number(b.totalAmount) || 0), 0);
  const dt = new Date(date + 'T00:00:00');
  const dateLabel = dt.toLocaleDateString('en-GB', { weekday: 'long', day: '2-digit', month: 'long', year: 'numeric' });
  let billsHtml = '';
  bills.forEach(b => {
    const items = Array.isArray(b.items) ? b.items : [];
    const itemRows = items.map(it =>
      '<tr><td style="padding:6px 8px;border-bottom:1px dashed #e5e7eb;">' + (it.name || '') + '</td>' +
      '<td style="padding:6px 8px;border-bottom:1px dashed #e5e7eb;text-align:center;">' + (it.quantity || 1) + '</td>' +
      '<td style="padding:6px 8px;border-bottom:1px dashed #e5e7eb;text-align:right;">Rs. ' + Number(it.price || 0).toLocaleString() + '</td></tr>'
    ).join('');
    billsHtml += '<div style="background:#fff;border:2px dashed #d1d5db;border-radius:12px;padding:16px;margin-bottom:16px;">' +
      '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;padding-bottom:8px;border-bottom:1px dashed #e5e7eb;">' +
      '<div><div style="font-weight:700;font-size:16px;">👤 ' + (b.personName || 'N/A') + '</div>' +
      '<div style="font-size:12px;color:#6b7280;">📋 ' + (b.billPurpose || '') + '</div></div>' +
      '<div style="font-size:12px;color:#6b7280;">🕐 ' + (b.time || '') + '</div></div>' +
      '<table style="width:100%;border-collapse:collapse;font-size:13px;">' +
      '<thead><tr style="background:#f9fafb;"><th style="padding:6px 8px;text-align:left;font-size:11px;color:#6b7280;">Item</th>' +
      '<th style="padding:6px 8px;text-align:center;font-size:11px;color:#6b7280;">Qty</th>' +
      '<th style="padding:6px 8px;text-align:right;font-size:11px;color:#6b7280;">Price</th></tr></thead>' +
      '<tbody>' + itemRows + '</tbody></table>' +
      '<div style="display:flex;justify-content:space-between;padding:10px 8px;margin-top:8px;background:#f0fdf4;border-radius:8px;font-weight:700;font-size:15px;color:#059669;">' +
      '<span>Total</span><span>Rs. ' + Number(b.totalAmount || 0).toLocaleString() + '</span></div></div>';
  });
  const html = '<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{font-family:Arial,sans-serif;margin:0;padding:24px;background:#f8fafc;}</style></head><body>' +
    '<div style="max-width:500px;margin:0 auto;">' +
    '<div style="text-align:center;margin-bottom:20px;">' +
    '<div style="font-size:22px;font-weight:800;color:#059669;">🧾 BILLS</div>' +
    '<div style="font-size:14px;color:#6b7280;">' + dateLabel + '</div>' +
    '<div style="font-size:13px;color:#059669;font-weight:600;margin-top:4px;">Grand Total: Rs. ' + total.toLocaleString() + '</div></div>' +
    billsHtml +
    '<div style="text-align:center;font-size:11px;color:#9ca3af;margin-top:16px;">' + BOT_NAME + ' • ' + SITE_URL + '</div></div></body></html>';

  // Try puppeteer, fallback to plain buffer
  try {
    const puppeteer = require('puppeteer');
    const browser = await puppeteer.launch({
      headless: 'new',
      args: ['--no-sandbox', '--disable-setuid-sandbox', '--disable-dev-shm-usage']
    });
    const page = await browser.newPage();
    await page.setViewport({ width: 600, height: 800, deviceScaleFactor: 2 });
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const body = await page.$('body');
    const buf = await body.screenshot({ type: 'png' });
    await browser.close();
    return Buffer.from(buf);
  } catch (puppeteerErr) {
    log('[BILL-IMG] Puppeteer not available, generating simple image: ' + puppeteerErr.message);
    throw new Error('Puppeteer not available');
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

async function addExcelFooter(ws) {
  ws.addRow({});
  const footerRow1 = ws.addRow({ no: '', name: BOT_NAME });
  footerRow1.font = { bold: true, size: 11, color: { argb: 'FF25D366' } };
  const footerRow2 = ws.addRow({ no: '', name: SITE_URL });
  footerRow2.font = { color: { argb: 'FF0066CC' }, underline: true };
  const footerRow3 = ws.addRow({ no: '', name: 'Admin: +' + ADMIN_NUMBERS.join(', +') });
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

  const inEntries = data.filter(e => e.type === 'in');
  const outEntries = data.filter(e => e.type !== 'in');
  const totalIn = inEntries.reduce((a, e) => a + Number(e.amount || 0), 0);
  const totalOut = outEntries.reduce((a, e) => a + Number(e.amount || 0), 0);
  const balance = totalIn - totalOut;
  const maxRows = Math.max(inEntries.length, outEntries.length);

  // ── Side-by-side: LEFT = IN (green), RIGHT = OUT (red) ──
  ws.columns = [
    { header: '#', key: 'in_no', width: 4 },
    { header: '📥 RECEIVED FROM', key: 'in_from', width: 28 },
    { header: 'AMOUNT (Rs)', key: 'in_amount', width: 15 },
    { header: 'DATE', key: 'in_date', width: 14 },
    { header: '', key: 'gap', width: 3 },
    { header: '#', key: 'out_no', width: 4 },
    { header: '📤 SPENT ON', key: 'out_for', width: 28 },
    { header: 'AMOUNT (Rs)', key: 'out_amount', width: 15 },
    { header: 'DATE', key: 'out_date', width: 14 },
  ];

  // Header styling - GREEN for IN side, RED for OUT side
  const headerRow = ws.getRow(1);
  headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
  [1,2,3,4].forEach(c => { headerRow.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00AA00' } }; });
  [6,7,8,9].forEach(c => { headerRow.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC0000' } }; });
  headerRow.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

  // Data rows
  for (let i = 0; i < maxRows; i++) {
    const rowData = { gap: '' };
    if (i < inEntries.length) {
      const e = inEntries[i];
      rowData.in_no = i + 1;
      rowData.in_from = e.personOrPurpose || '';
      rowData.in_amount = Number(e.amount) || 0;
      rowData.in_date = e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB') : '';
    }
    if (i < outEntries.length) {
      const e = outEntries[i];
      rowData.out_no = i + 1;
      rowData.out_for = e.personOrPurpose || '';
      rowData.out_amount = Number(e.amount) || 0;
      rowData.out_date = e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB') : '';
    }
    const row = ws.addRow(rowData);
    // Light green tint for IN side, light red tint for OUT side
    if (i < inEntries.length) {
      [1,2,3,4].forEach(c => { row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } }; });
    }
    if (i < outEntries.length) {
      [6,7,8,9].forEach(c => { row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4EC' } }; });
    }
  }

  // Summary rows
  ws.addRow({});
  const sumIn = ws.addRow({ in_from: 'TOTAL RECEIVED', in_amount: totalIn });
  sumIn.font = { bold: true, color: { argb: 'FF00AA00' }, size: 12 };
  const sumOut = ws.addRow({ out_for: 'TOTAL SPENT', out_amount: totalOut });
  sumOut.font = { bold: true, color: { argb: 'FFCC0000' }, size: 12 };
  ws.addRow({});
  const balRow = ws.addRow({ in_from: 'BALANCE', in_amount: balance });
  balRow.font = { bold: true, size: 14 };

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
  const sheetName = section.charAt(0).toUpperCase() + section.slice(1);
  const ws = wb.addWorksheet(sheetName);
  if (!data.length) return null;

  // Separate IN and OUT for samples/clipping
  const hasInOut = ['samples', 'clipping'].includes(section);
  const inEntries = hasInOut ? data.filter(e => e.type === 'in') : [];
  const outEntries = hasInOut ? data.filter(e => e.type !== 'in') : [];

  if (hasInOut && inEntries.length && outEntries.length) {
    // ── Side-by-side layout ──
    const keys = Object.keys(data[0]).filter(k => k !== 'id');
    const halfCols = keys.map(k => ({ header: k.charAt(0).toUpperCase() + k.slice(1), key: k, width: 18 }));
    ws.columns = [
      { header: '#', key: 'in_no', width: 4 },
      ...halfCols.map(c => ({ ...c, key: 'in_' + c.key })),
      { header: '', key: 'gap', width: 3 },
      { header: '#', key: 'out_no', width: 4 },
      ...halfCols.map(c => ({ ...c, key: 'out_' + c.key })),
    ];
    const headerRow = ws.getRow(1);
    headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
    const inColCount = halfCols.length + 1;
    for (let c = 1; c <= inColCount; c++) headerRow.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF00AA00' } };
    for (let c = inColCount + 2; c <= ws.columnCount; c++) headerRow.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCC0000' } };
    headerRow.getCell(inColCount + 1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

    const maxRows = Math.max(inEntries.length, outEntries.length);
    for (let i = 0; i < maxRows; i++) {
      const rowData = { gap: '' };
      if (i < inEntries.length) {
        rowData.in_no = i + 1;
        keys.forEach(k => { rowData['in_' + k] = inEntries[i][k] != null ? String(inEntries[i][k]) : ''; });
      }
      if (i < outEntries.length) {
        rowData.out_no = i + 1;
        keys.forEach(k => { rowData['out_' + k] = outEntries[i][k] != null ? String(outEntries[i][k]) : ''; });
      }
      const row = ws.addRow(rowData);
      if (i < inEntries.length) for (let c = 1; c <= inColCount; c++) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE8F5E9' } };
      if (i < outEntries.length) for (let c = inColCount + 2; c <= ws.columnCount; c++) row.getCell(c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4EC' } };
    }
  } else {
    // Fallback: single table layout
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
  }
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

    // Extract text from various message types (text, button clicks, list responses, native flow)
    let text = '';
    const msgType = Object.keys(msg.message)[0];
    if (msgType === 'interactiveResponseMessage') {
      // Native flow button click
      const ir = msg.message.interactiveResponseMessage;
      if (ir.nativeFlowResponseMessage) {
        try { text = JSON.parse(ir.nativeFlowResponseMessage.paramsJson).id || ''; } catch(e) { text = ''; }
      } else { text = ir.body || ''; }
    } else if (msgType === 'templateButtonReplyMessage') {
      text = msg.message.templateButtonReplyMessage.selectedId || '';
    } else if (msg.message.buttonsResponseMessage) {
      text = msg.message.buttonsResponseMessage.selectedButtonId || '';
    } else if (msg.message.listResponseMessage) {
      text = msg.message.listResponseMessage.singleSelectReply?.selectedRowId || '';
    } else {
      text = msg.message.conversation
        || msg.message.extendedTextMessage?.text
        || '';
    }

    if (!text || text.trim().length < 1) return;

    // Strip ALL leading symbols (., /, !, #, *, etc.) so .items /items !items items all work
    const stripped = text.trim().replace(/^[^a-zA-Z0-9]+/, '');
    if (!stripped) return;
    const parts = stripped.split(/\s+/);
    const cmd = parts[0].toLowerCase();

    try {
      // .ping
      if (cmd === 'ping') {
        const start = Date.now();
        await sock.sendMessage(chatId, { text: 'Pong!' });
        const latency = Date.now() - start;
        await sock.sendMessage(chatId, {
          text: '*' + BOT_NAME + ' Bot*\nLatency: ' + latency + 'ms\nStatus: Online'
        });
      }

      // .sendoutbox
      if (cmd === 'sendoutbox') {
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
      if (cmd === 'status') {
        const outbox = readOutbox();
        const pending = outbox.filter(m => !m.sent).length;
        await sock.sendMessage(chatId, {
          text: '*' + BOT_NAME + '*\nBot: Online\nAccount: ' + (sock.user?.id || 'unknown') + '\nPending messages: ' + pending + '\nUptime: ' + Math.floor(process.uptime()) + 's'
        });
      }

      // ─── Stock Manager Commands ────────────────────────────────────────────

      // .items - send items.xlsx
      if (cmd === 'items') {
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
      if (cmd === 'items2') {
        await sock.sendMessage(chatId, { text: '📦 Fetching items text...' });
        const data = await fetchStockData('items');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No items data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatItemsText(data) });
      }

      // .wallet - send wallet.xlsx
      if (cmd === 'wallet') {
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
      if (cmd === 'wallet2' || cmd === 'wallettxt') {
        const data = await fetchStockData('wallet');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No wallet data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatWalletText(data) });
      }

      // .person - send person.xlsx
      if (cmd === 'person') {
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
      if (cmd === 'person2') {
        await sock.sendMessage(chatId, { text: '👷 Fetching person text...' });
        const data = await fetchStockData('person');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No person data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatPersonText(data) });
      }

      // .maintenance - send maintenance.xlsx
      if (cmd === 'maintenance') {
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
      if (cmd === 'maintenance2') {
        await sock.sendMessage(chatId, { text: '🔧 Fetching maintenance text...' });
        const data = await fetchStockData('maintenance');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No maintenance data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatMaintenanceText(data) });
      }

      // .samples - send samples.xlsx
      if (cmd === 'samples') {
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
      if (cmd === 'samples2') {
        await sock.sendMessage(chatId, { text: '🧪 Fetching samples text...' });
        const data = await fetchStockData('samples');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No samples data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatSamplesText(data) });
      }

      // .clipping - send clipping.xlsx
      if (cmd === 'clipping') {
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
      if (cmd === 'clipping2') {
        await sock.sendMessage(chatId, { text: '✂️ Fetching clipping text...' });
        const data = await fetchStockData('clipping');
        if (!data.length) { await sock.sendMessage(chatId, { text: '❌ No clipping data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatClippingText(data) });
      }

      // .salary - Calculate salary for a person
      if (cmd === 'salary') {
        const ZAID_SALARY = 42000;
        // Format: .salary,{name},{monthly_salary} or .salary {name} {monthly_salary}
        // Special: .salary zaid (no salary needed, fixed 42000)
        const raw = stripped.substring(parts[0].length).trim();
        let name = '', salaryStr = '';
        if (raw.includes(',')) {
          const salaryParts = raw.split(',').map(s => s.trim());
          name = salaryParts[0] || '';
          salaryStr = salaryParts[1] || '';
        } else {
          const salaryParts = raw.split(/\s+/);
          name = salaryParts[0] || '';
          salaryStr = salaryParts[1] || '';
        }
        const isZaid = name.toUpperCase() === 'ZAID';
        if (!name || (!salaryStr && !isZaid)) {
          await sock.sendMessage(chatId, { text: '💰 *Salary Calculator*\n\nUsage: `salary,{name},{monthly_salary}`\nExample: `salary,ADNAN,25000`\n\nSpecial: `salary zaid` (fixed Rs. 42,000)' });
          return;
        }
        let monthlySalary;
        if (isZaid) {
          monthlySalary = ZAID_SALARY;
        } else {
          monthlySalary = parseFloat(salaryStr);
          if (isNaN(monthlySalary) || monthlySalary <= 0) {
            await sock.sendMessage(chatId, { text: '❌ Invalid salary amount. Example: `salary,ADNAN,25000`' });
            return;
          }
        }
        await sock.sendMessage(chatId, { text: '💰 Calculating salary for *' + name.toUpperCase() + '*...' });

        try {
          const personData = await fetchStockData('person');
          if (!personData.length) {
            await sock.sendMessage(chatId, { text: '❌ No attendance data found.' });
            return;
          }

          // Always use previous full month (1st to last day)
          const today = new Date();
          const periodStart = new Date(today.getFullYear(), today.getMonth() - 1, 1);
          const periodEnd = new Date(today.getFullYear(), today.getMonth(), 0, 23, 59, 59, 999);
          const daysInMonth = new Date(today.getFullYear(), today.getMonth(), 0).getDate();
          const monthName = periodStart.toLocaleString('default', { month: 'long', year: 'numeric' });

          let daysWorked = 0;
          let absents = 0;
          const actualName = isZaid ? 'ZAID' : '';

          if (!isZaid) {
            // Find person case-insensitive
            const personName = personData.find(e => (e.personName || '').toUpperCase() === name.toUpperCase());
            if (!personName) {
              await sock.sendMessage(chatId, { text: '❌ Person "' + name + '" not found in attendance records.' });
              return;
            }
            const realName = personName.personName;

            // Count days worked in the previous full month
            const personEntries = personData
              .filter(e => e.personName === realName && e.action === 'enter')
              .map(e => ({ ...e, _date: new Date(e.timestamp) }))
              .filter(e => e._date >= periodStart && e._date <= periodEnd);

            const workDates = new Set();
            personEntries.forEach(e => {
              workDates.add(e._date.toDateString());
            });
            daysWorked = workDates.size;
            absents = daysInMonth - daysWorked;

            // Update actualName for display
            name = realName;
          } else {
            // ZAID: 0 absents, full month salary
            daysWorked = daysInMonth;
            absents = 0;
          }

          const dailyRate = monthlySalary / daysInMonth;
          const earnedSalary = Math.round(dailyRate * daysWorked);
          const periodStartStr = periodStart.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
          const periodEndStr = periodEnd.toLocaleDateString('en-GB', { day: '2-digit', month: 'short', year: 'numeric' });

          const msg = [
            '💰 ✦ *' + toBold('SALARY CALCULATOR') + '* ✦ 💰',
            LINE,
            '👤 *Name:* ' + (isZaid ? actualName + ' ⭐' : name),
            '💵 *Monthly Salary:* Rs. ' + monthlySalary.toLocaleString(),
            '📅 *Period:* ' + periodStartStr + ' — ' + periodEndStr,
            '📆 *Month:* ' + monthName + ' (' + daysInMonth + ' days)',
            DIV,
            '📊 *' + toBold('CALCULATION') + '*',
            THIN,
            '📅 *Days Worked:* ' + daysWorked,
            '❌ *Absents:* ' + absents,
            '💰 *Daily Rate:* Rs. ' + Math.round(dailyRate).toLocaleString() + ' (' + monthlySalary.toLocaleString() + ' ÷ ' + daysInMonth + ')',
            THIN,
            '✅ *Total Salary:* Rs. ' + earnedSalary.toLocaleString(),
            '📝 *Formula:* ' + daysWorked + ' × ' + Math.round(dailyRate).toLocaleString() + ' = ' + earnedSalary.toLocaleString(),
            isZaid ? '⭐ *ZAID: 0 absents, full month salary*' : '',
            FOOTER
          ].filter(Boolean).join('\n');
          await sock.sendMessage(chatId, { text: msg });
        } catch (err) {
          await sock.sendMessage(chatId, { text: '❌ Error calculating salary: ' + err.message });
        }
      }

      // .stock / .help - show all commands
      if (cmd === 'stock' || cmd === 'help') {
        const helpMsg = [
          '🤖 *' + toBold(BOT_NAME) + '*',
          LINE,
          '',
          '📂 *' + toBold('STOCK & INVENTORY') + '*',
          THIN,
          '📦 *items* — Items Excel file',
          '📝 *items2* — Items text details',
          '',
          '💰 *' + toBold('WALLET & FINANCE') + '*',
          THIN,
          '💵 *wallet* — Wallet Excel (IN/OUT side-by-side)',
          '📝 *wallet2* — Wallet text details',
          '',
          '👷 *' + toBold('ATTENDANCE') + '*',
          THIN,
          '📋 *person* — Attendance Excel',
          '📝 *person2* — Attendance text',
          '',
          '🔧 *' + toBold('MAINTENANCE') + '*',
          THIN,
          '🔧 *maintenance* — Maintenance Excel',
          '📝 *maintenance2* — Maintenance text',
          '',
          '🧪 *' + toBold('SAMPLES') + '*',
          THIN,
          '🧪 *samples* — Samples Excel (IN/OUT)',
          '📝 *samples2* — Samples text',
          '',
          '✂️ *' + toBold('CLIPPING') + '*',
          THIN,
          '✂️ *clipping* — Clipping Excel (IN/OUT)',
          '📝 *clipping2* — Clipping text',
          '',
          '⚙️ *' + toBold('BOT CONTROLS') + '*',
          THIN,
          '🏓 *ping* — Check latency',
          '📊 *status* — Bot status & stats',
          '📨 *sendoutbox* — Process message queue',
          '',
          '💰 *' + toBold('SALARY CALCULATOR') + '*',
          THIN,
          '💰 *salary,{name},{amount}* — Calculate salary',
          '_Example: salary,ADNAN,25000_',
          '',
          '💡 _Send any command with or without symbols_',
          '_Example: .items /items items !items all work_',
          '',
          LINE,
          '🏢 *' + BOT_NAME + '*',
          '🌐 ' + SITE_URL,
          '📱 Admin: +' + ADMIN_NUMBER
        ].join('\n');
        await sock.sendMessage(chatId, { text: helpMsg });
      }

      // ─── ZAID Interactive Buttons (hierarchical) ──────────────────────
      if (cmd === 'zaid') {
        if (sendButtons) {
          try {
            await sendButtons(sock, chatId, {
              text: '🤖 *' + BOT_NAME + '*\n\n👋 Welcome! Tap a category:',
              footer: '🌐 ' + SITE_URL + ' • 📱 +' + ADMIN_NUMBER,
              buttons: [
                { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📦 Items', id: 'zaid_items' }) },
                { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '✂️ Clipping', id: 'zaid_clipping' }) },
                { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '🧪 Samples', id: 'zaid_samples' }) },
                { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '💰 Wallet', id: 'zaid_wallet' }) },
                { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '🧾 Bills', id: 'zaid_bills' }) },
                { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '❓ Help', id: 'help' }) },
              ]
            });
          } catch (btnErr) {
            log('[ZAID] Button send error: ' + btnErr.message);
            await sock.sendMessage(chatId, { text: '🤖 *' + BOT_NAME + '*\n\nCommands: items, wallet, clipping, samples, bills, help' });
          }
        } else {
          await sock.sendMessage(chatId, { text: '🤖 *' + BOT_NAME + '*\n\nCommands: items, wallet, clipping, samples, bills, help' });
        }
      }

      // ─── Category → Format buttons ───────────────────────────────────
      if (cmd === 'zaid_items') {
        if (sendButtons) {
          await sendButtons(sock, chatId, {
            text: '📦 *Items*\nChoose format:', footer: BOT_NAME,
            buttons: [
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📝 Text', id: 'items2' }) },
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📊 Excel File', id: 'items' }) },
            ]
          });
        } else { await sock.sendMessage(chatId, { text: 'Send: items (Excel) or items2 (Text)' }); }
      }
      if (cmd === 'zaid_clipping') {
        if (sendButtons) {
          await sendButtons(sock, chatId, {
            text: '✂️ *Clipping*\nChoose format:', footer: BOT_NAME,
            buttons: [
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📝 Text', id: 'clipping2' }) },
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📊 Excel File', id: 'clipping' }) },
            ]
          });
        } else { await sock.sendMessage(chatId, { text: 'Send: clipping (Excel) or clipping2 (Text)' }); }
      }
      if (cmd === 'zaid_samples') {
        if (sendButtons) {
          await sendButtons(sock, chatId, {
            text: '🧪 *Samples*\nChoose format:', footer: BOT_NAME,
            buttons: [
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📝 Text', id: 'samples2' }) },
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📊 Excel File', id: 'samples' }) },
            ]
          });
        } else { await sock.sendMessage(chatId, { text: 'Send: samples (Excel) or samples2 (Text)' }); }
      }
      if (cmd === 'zaid_wallet') {
        if (sendButtons) {
          await sendButtons(sock, chatId, {
            text: '💰 *Wallet*\nChoose format:', footer: BOT_NAME,
            buttons: [
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📝 Text', id: 'wallet2' }) },
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📊 Excel File', id: 'wallet' }) },
            ]
          });
        } else { await sock.sendMessage(chatId, { text: 'Send: wallet (Excel) or wallet2 (Text)' }); }
      }

      // ─── Bills → Date buttons ─────────────────────────────────────────
      if (cmd === 'zaid_bills') {
        const billsData = await fetchStockData('bills');
        if (!billsData.length) { await sock.sendMessage(chatId, { text: '🧾 No bills available yet.' }); return; }
        const dates = [...new Set(billsData.map(b => b.date || (b.timestamp ? new Date(b.timestamp).toISOString().slice(0, 10) : null)).filter(Boolean))].sort().reverse();
        if (sendButtons && dates.length > 0) {
          const dateButtons = dates.slice(0, 8).map(d => {
            const dt = new Date(d + 'T00:00:00');
            const label = dt.toLocaleDateString('en-GB', { day: '2-digit', month: 'short' });
            return { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📅 ' + label, id: 'bill_date_' + d }) };
          });
          await sendButtons(sock, chatId, {
            text: '🧾 *Bills*\nSelect a date:', footer: BOT_NAME, buttons: dateButtons
          });
        } else {
          await sock.sendMessage(chatId, { text: '🧾 *Bills dates:*\n' + dates.join('\n') });
        }
      }

      // ─── Bill date selected → format buttons ─────────────────────────
      if (cmd.startsWith('bill_date_')) {
        const date = cmd.replace('bill_date_', '');
        if (sendButtons) {
          await sendButtons(sock, chatId, {
            text: '🧾 *Bills — ' + date + '*\nChoose format:', footer: BOT_NAME,
            buttons: [
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '🖼️ Image', id: 'bill_img_' + date }) },
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📝 Text', id: 'bill_txt_' + date }) },
              { name: 'quick_reply', buttonParamsJson: JSON.stringify({ display_text: '📊 Excel', id: 'bill_xls_' + date }) },
            ]
          });
        } else { await sock.sendMessage(chatId, { text: 'Formats: bill_txt_' + date + ' / bill_xls_' + date }); }
      }

      // ─── Bill text ──────────────────────────────────────────────────
      if (cmd.startsWith('bill_txt_')) {
        const date = cmd.replace('bill_txt_', '');
        const billsData = await fetchStockData('bills');
        const dateBills = billsData.filter(b => (b.date || (b.timestamp ? new Date(b.timestamp).toISOString().slice(0, 10) : '')) === date);
        if (!dateBills.length) { await sock.sendMessage(chatId, { text: '❌ No bills for ' + date }); return; }
        await sock.sendMessage(chatId, { text: formatBillsText(dateBills, date) });
      }

      // ─── Bill Excel ──────────────────────────────────────────────────
      if (cmd.startsWith('bill_xls_')) {
        const date = cmd.replace('bill_xls_', '');
        const billsData = await fetchStockData('bills');
        const dateBills = billsData.filter(b => (b.date || (b.timestamp ? new Date(b.timestamp).toISOString().slice(0, 10) : '')) === date);
        if (!dateBills.length) { await sock.sendMessage(chatId, { text: '❌ No bills for ' + date }); return; }
        const buf = await generateBillsXlsx(dateBills, date);
        const tmpFile = path.join(__dirname, 'tmp', 'bills_' + date + '_' + Date.now() + '.xlsx');
        if (!fs.existsSync(path.join(__dirname, 'tmp'))) fs.mkdirSync(path.join(__dirname, 'tmp'), { recursive: true });
        fs.writeFileSync(tmpFile, buf);
        const total = dateBills.reduce((a, b) => a + (Number(b.totalAmount) || 0), 0);
        await sock.sendMessage(chatId, {
          document: fs.readFileSync(tmpFile), fileName: 'bills_' + date + '.xlsx',
          mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          caption: '🧾 *BILLS — ' + date + '*\n📦 ' + dateBills.length + ' bills\n💰 Rs. ' + total.toLocaleString() + '\n🏢 ' + BOT_NAME
        });
        fs.unlinkSync(tmpFile);
      }

      // ─── Bill Image ──────────────────────────────────────────────────
      if (cmd.startsWith('bill_img_')) {
        const date = cmd.replace('bill_img_', '');
        const billsData = await fetchStockData('bills');
        const dateBills = billsData.filter(b => (b.date || (b.timestamp ? new Date(b.timestamp).toISOString().slice(0, 10) : '')) === date);
        if (!dateBills.length) { await sock.sendMessage(chatId, { text: '❌ No bills for ' + date }); return; }
        await sock.sendMessage(chatId, { text: '🖼️ Generating bill image...' });
        try {
          const imgBuf = await generateBillImage(dateBills, date);
          await sock.sendMessage(chatId, {
            image: imgBuf,
            caption: '🧾 *BILLS — ' + date + '*\n💰 Rs. ' + dateBills.reduce((a, b) => a + (Number(b.totalAmount) || 0), 0).toLocaleString() + '\n🏢 ' + BOT_NAME
          });
        } catch (imgErr) {
          log('[BILL-IMG] Error: ' + imgErr.message);
          await sock.sendMessage(chatId, { text: '⚠️ Image failed, sending text.\n\n' + formatBillsText(dateBills, date) });
        }
      }

      // ─── .bills direct command ───────────────────────────────────────
      if (cmd === 'bills') {
        const billsData = await fetchStockData('bills');
        if (!billsData.length) { await sock.sendMessage(chatId, { text: '🧾 No bills data found.' }); return; }
        const buf = await generateBillsXlsx(billsData, 'all');
        const tmpFile = path.join(__dirname, 'tmp', 'bills_' + Date.now() + '.xlsx');
        if (!fs.existsSync(path.join(__dirname, 'tmp'))) fs.mkdirSync(path.join(__dirname, 'tmp'), { recursive: true });
        fs.writeFileSync(tmpFile, buf);
        const total = billsData.reduce((a, b) => a + (Number(b.totalAmount) || 0), 0);
        await sock.sendMessage(chatId, {
          document: fs.readFileSync(tmpFile), fileName: 'bills.xlsx',
          mimetype: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
          caption: '🧾 *BILLS* (' + billsData.length + ' | 💰 Rs. ' + total.toLocaleString() + ')\n🏢 ' + BOT_NAME
        });
        fs.unlinkSync(tmpFile);
      }

      if (cmd === 'bills2') {
        const billsData = await fetchStockData('bills');
        if (!billsData.length) { await sock.sendMessage(chatId, { text: '🧾 No bills data found.' }); return; }
        await sock.sendMessage(chatId, { text: formatBillsText(billsData, 'all') });
      }

      // ─── TILLA Search ────────────────────────────────────────────────────
      if (cmd === 'tilla') {
        await sock.sendMessage(chatId, { text: '🔍 Searching for *TILLA* in items...' });
        try {
          const itemsData = await fetchStockData('items');
          const tillaItems = itemsData.filter(e => {
            const searchFields = [e.name, e.number, e.model, e.person].filter(Boolean).join(' ').toLowerCase();
            return searchFields.includes('tilla');
          });
          if (!tillaItems.length) {
            await sock.sendMessage(chatId, { text: '⚠️ No items found matching *TILLA*.' });
            return;
          }
          const div = '──────────────────────────────';
          const lines = tillaItems.map((e, i) => [
            `${i+1}. 📋 *${e.name || 'N/A'}*`,
            `   🔢 Serial: ${e.number || 'N/A'}`,
            `   👤 Person: ${e.person || 'N/A'}`,
            `   📐 Model: ${e.model || 'N/A'}`,
            `   🔢 Qty: ${e.quantity || 1}`,
            `   📊 Status: ${e.status || 'available'}`
          ].join('\n')).join('\n' + div + '\n');

          const msg = [
            '🔍 ✦ *' + toBold('TILLA SEARCH') + '* ✦ 🔍',
            LINE,
            '📦 *Found:* ' + tillaItems.length + ' items',
            DIV,
            lines,
            FOOTER
          ].join('\n');
          await sock.sendMessage(chatId, { text: msg });
        } catch (err) {
          await sock.sendMessage(chatId, { text: '❌ Error searching: ' + err.message });
        }
      }

      // ─── Serial Number Search (3-4 digits) ──────────────────────────────
      if (/^\d{3,4}$/.test(stripped)) {
        const searchNum = stripped;
        await sock.sendMessage(chatId, { text: '🔍 Searching item serial: *' + searchNum + '*...' });
        try {
          const itemsData = await fetchStockData('items');
          // Search by serial number (exact or partial match)
          const matched = itemsData.filter(e => {
            const num = String(e.number || '').trim();
            return num === searchNum || num.includes(searchNum) || num.endsWith(searchNum);
          });
          if (!matched.length) {
            await sock.sendMessage(chatId, { text: '⚠️ *Thread not available!*\n\nNo item found with serial number *' + searchNum + '*.\n\nPlease check the number and try again.' });
            return;
          }
          const div = '──────────────────────────────';
          const lines = matched.map((e, i) => [
            `${i+1}. 📋 *${e.name || 'N/A'}*`,
            `   🔢 Serial: ${e.number || 'N/A'}`,
            `   👤 Person: ${e.person || 'N/A'}`,
            `   📐 Model: ${e.model || 'N/A'}`,
            `   🔢 Qty: ${e.quantity || 1}`,
            `   📊 Status: ${e.status || 'available'}`,
            `   🕐 ${e.timestamp ? new Date(e.timestamp).toLocaleDateString('en-GB', { day:'2-digit', month:'short', year:'numeric' }) : 'N/A'}`
          ].join('\n')).join('\n' + div + '\n');

          const msg = [
            '🔍 ✦ *' + toBold('SERIAL SEARCH') + '* ✦ 🔍',
            LINE,
            '🔢 *Search:* ' + searchNum,
            '📦 *Found:* ' + matched.length + ' item(s)',
            DIV,
            lines,
            FOOTER
          ].join('\n');
          await sock.sendMessage(chatId, { text: msg });
        } catch (err) {
          await sock.sendMessage(chatId, { text: '❌ Error searching: ' + err.message });
        }
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
  // Support multiple targets: if 'to' is given, use it; otherwise send to ALL admins
  const targets = to ? [to] : ADMIN_NUMBERS;
  const outbox = readOutbox();
  const ids = [];
  targets.forEach(target => {
    const entry = {
      id: 'msg_' + Date.now() + '_' + Math.random().toString(36).slice(2, 8),
      to: target,
      message,
      queuedAt: new Date().toISOString(),
      sent: false
    };
    outbox.push(entry);
    ids.push(entry.id);
    log('[API] Queued message for +' + target);
    if (whatsappConnected && sockRef) {
      processImmediateSend(sockRef, entry).catch(err => {
        log('[API] Immediate send failed for +' + target + ': ' + err.message);
      });
    }
  });
  writeOutbox(outbox);
  res.json({ success: true, message: 'Message queued for ' + targets.map(t => '+' + t).join(', '), ids });
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
