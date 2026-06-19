/**
 * BANU SAEED HOSPITAL - WhatsApp API Server
 * 
 * Runs:
 *   1. Express HTTP API on process.env.PORT (for Vercel frontend to call)
 *   2. Clean WhatsApp bot (index.js) as child process with auto-send
 */

const express = require('express');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(express.json());

const PORT = process.env.PORT || 1000;
const API_SECRET = process.env.API_SECRET || 'banu-saeed-secret-2024';
const ADMIN_NUMBER = process.env.ADMIN_NUMBER || '923299931199';
const OUTBOX_PATH = path.join(__dirname, 'data', 'outbox.json');

let whatsappConnected = false;
let botLogs = [];

// ─── Helpers ────────────────────────────────────────────────────────────────

function readOutbox() {
  try {
    if (!fs.existsSync(OUTBOX_PATH)) {
      fs.writeFileSync(OUTBOX_PATH, '[]', 'utf8');
      return [];
    }
    return JSON.parse(fs.readFileSync(OUTBOX_PATH, 'utf8'));
  } catch { return []; }
}

function writeOutbox(data) {
  try {
    const dir = path.dirname(OUTBOX_PATH);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(OUTBOX_PATH, JSON.stringify(data, null, 2), 'utf8');
    return true;
  } catch (err) {
    console.error('Failed to write outbox:', err.message);
    return false;
  }
}

function authMiddleware(req, res, next) {
  const { secret } = req.body;
  if (secret !== API_SECRET) {
    return res.status(401).json({ success: false, error: 'Unauthorized' });
  }
  next();
}

// ─── Health Check ────────────────────────────────────────────────────────────

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

// ─── Status ──────────────────────────────────────────────────────────────────

app.get('/status', (req, res) => {
  const outbox = readOutbox();
  const pending = outbox.filter(m => !m.sent).length;
  const sent = outbox.filter(m => m.sent).length;
  res.json({
    online: true,
    botRunning: botProcess ? !botProcess.killed : false,
    whatsappConnected,
    pendingMessages: pending,
    sentMessages: sent,
    adminNumber: ADMIN_NUMBER,
    uptime: Math.floor(process.uptime()) + 's',
    recentLogs: botLogs.slice(-10)
  });
});

// ─── Send Message (queues for bot to auto-send) ─────────────────────────────

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
    res.json({ success: true, message: 'Message queued for +' + target, id: entry.id });
  } catch (err) {
    res.status(500).json({ success: false, error: err.message });
  }
});

// ─── Pair endpoint ──────────────────────────────────────────────────────────

app.post('/pair', authMiddleware, (req, res) => {
  res.json({
    success: true,
    message: 'Check Railway logs for QR code. Scan in WhatsApp > Linked Devices > Link a Device.'
  });
});

// ─── Debug endpoint ──────────────────────────────────────────────────────────

app.get('/debug', (req, res) => {
  res.json({
    online: true,
    botRunning: botProcess ? !botProcess.killed : false,
    whatsappConnected,
    uptime: Math.floor(process.uptime()) + 's',
    nodeVersion: process.version,
    env: {
      PORT: process.env.PORT || '1000 (default)',
      SESSION_ID: process.env.SESSION_ID ? 'SET (' + process.env.SESSION_ID.substring(0, 8) + '...)' : 'NOT SET',
      API_SECRET: process.env.API_SECRET ? 'SET' : 'NOT SET',
      ADMIN_NUMBER: process.env.ADMIN_NUMBER || '923299931199 (default)',
    },
    recentLogs: botLogs.slice(-20)
  });
});

// ─── Start Server ────────────────────────────────────────────────────────────

let botProcess = null;

app.listen(PORT, () => {
  console.log('═══════════════════════════════════════════════');
  console.log('  BANU SAEED HOSPITAL - WhatsApp API Server');
  console.log('═══════════════════════════════════════════════');
  console.log('  API listening on port ' + PORT);
  console.log('  Admin: +' + ADMIN_NUMBER);
  console.log('═══════════════════════════════════════════════');
  startBot();
});

// ─── Bot Process ─────────────────────────────────────────────────────────────

function parseBotOutput(text) {
  const lower = text.toLowerCase();

  if (lower.includes('connected') || lower.includes('connection open') || lower.includes('bot connected')) {
    if (!whatsappConnected) {
      whatsappConnected = true;
      console.log('[STATUS] WhatsApp CONNECTED');
    }
  }

  if (lower.includes('logged out') || lower.includes('401')) {
    if (whatsappConnected) {
      whatsappConnected = false;
      console.log('[STATUS] WhatsApp DISCONNECTED (logged out)');
    }
  }

  botLogs.push(text.substring(0, 300));
  if (botLogs.length > 50) botLogs.shift();
}

let restartCount = 0;

function startBot() {
  console.log('[BOT] Starting WhatsApp bot... (attempt ' + (restartCount + 1) + ')');
  parseBotOutput('Starting bot process...');

  // Set bot's PORT to avoid conflict with our Express
  const botEnv = { ...process.env, PORT: '3001' };

  botProcess = spawn('node', ['index.js'], {
    cwd: __dirname,
    stdio: ['ignore', 'pipe', 'pipe'],
    env: botEnv
  });

  botProcess.stdout.on('data', (data) => {
    const lines = data.toString().split('\n').filter(l => l.trim());
    lines.forEach(line => {
      console.log('[BOT]', line);
      parseBotOutput(line);
    });
  });

  botProcess.stderr.on('data', (data) => {
    const lines = data.toString().split('\n').filter(l => l.trim());
    lines.forEach(line => {
      console.error('[BOT ERR]', line);
      parseBotOutput(line);
    });
  });

  botProcess.on('exit', (code) => {
    restartCount++;
    const delay = restartCount > 10 ? 30000 : 5000;
    console.log('[BOT] Exited code ' + code + ' (restart #' + restartCount + '). Retrying in ' + (delay/1000) + 's...');
    parseBotOutput('Bot exited code ' + code);
    whatsappConnected = false;
    botProcess = null;
    setTimeout(startBot, delay);
  });

  botProcess.on('error', (err) => {
    restartCount++;
    console.error('[BOT] Start error: ' + err.message);
    parseBotOutput('Bot error: ' + err.message);
    whatsappConnected = false;
    botProcess = null;
    setTimeout(startBot, 5000);
  });
}

// ─── Graceful Shutdown ───────────────────────────────────────────────────────

process.on('SIGTERM', () => {
  if (botProcess) botProcess.kill('SIGTERM');
  process.exit(0);
});
process.on('SIGINT', () => {
  if (botProcess) botProcess.kill('SIGTERM');
  process.exit(0);
});
process.on('uncaughtException', (err) => {
  console.error('[SERVER] Uncaught:', err.message);
  parseBotOutput('Uncaught: ' + err.message);
});
process.on('unhandledRejection', (reason) => {
  console.error('[SERVER] Unhandled:', reason);
  parseBotOutput('Unhandled: ' + reason);
});
