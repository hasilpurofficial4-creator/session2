/**
 * SendOutbox Command - Processes message queue for Banu Saeed Hospital
 * 
 * Usage:
 *   .sendoutbox        - Process queue once
 *   .sendoutbox auto   - Start auto-processing every 15 seconds
 *   .sendoutbox stop   - Stop auto-processing
 *   .sendoutbox status - Show queue status
 */

const fs = require('fs');
const path = require('path');

const OUTBOX_PATH = path.join(__dirname, '..', 'data', 'outbox.json');
const ADMIN_NUMBER = process.env.ADMIN_NUMBER || '923299931199';

let autoInterval = null;

function readOutbox() {
  try {
    if (!fs.existsSync(OUTBOX_PATH)) return [];
    return JSON.parse(fs.readFileSync(OUTBOX_PATH, 'utf8'));
  } catch {
    return [];
  }
}

function writeOutbox(data) {
  try {
    fs.writeFileSync(OUTBOX_PATH, JSON.stringify(data, null, 2), 'utf8');
  } catch (err) {
    console.error('[OUTBOX] Write error:', err.message);
  }
}

async function processQueue(sock) {
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
      console.error('[OUTBOX] Failed: ' + err.message);
    }
  }

  const cleaned = outbox.filter(m => {
    if (!m.sent) return true;
    return (Date.now() - new Date(m.sentAt).getTime()) < 3600000;
  });

  writeOutbox(cleaned);
  return { processed, failed };
}

// 3-argument signature matching other bot commands like ping, alive
async function sendoutboxCommand(sock, chatId, message) {
  // Extract subcommand from message text: ".sendoutbox auto" -> "auto"
  const text = message.message?.conversation 
    || message.message?.extendedTextMessage?.text 
    || '';
  const parts = text.trim().split(/\s+/);
  const subCommand = (parts[1] || '').toLowerCase();

  try {
    if (subCommand === 'status') {
      const outbox = readOutbox();
      const pending = outbox.filter(m => !m.sent).length;
      const sent = outbox.filter(m => m.sent).length;
      const autoStatus = autoInterval ? 'Running (15s)' : 'Stopped';

      await sock.sendMessage(chatId, {
        text: '*BANU SAEED HOSPITAL*\n\nPending: *' + pending + '*\nSent: *' + sent + '*\nAuto: ' + autoStatus + '\n\n.sendoutbox - process now\n.sendoutbox auto - auto every 15s\n.sendoutbox stop - stop auto'
      });
      return;
    }

    if (subCommand === 'auto') {
      if (autoInterval) {
        await sock.sendMessage(chatId, { text: 'Auto-send already running!' });
        return;
      }

      autoInterval = setInterval(async () => {
        try {
          const r = await processQueue(sock);
          if (r.processed > 0) console.log('[AUTO-OUTBOX] Sent ' + r.processed);
        } catch (err) {
          console.error('[AUTO-OUTBOX] Error: ' + err.message);
        }
      }, 15000);

      const result = await processQueue(sock);
      await sock.sendMessage(chatId, {
        text: '*AUTO-SEND ENABLED*\nProcessed: *' + result.processed + '* sent, *' + result.failed + '* failed\nAuto-check every 15s\n\n.sendoutbox stop to disable'
      });
      return;
    }

    if (subCommand === 'stop') {
      if (autoInterval) {
        clearInterval(autoInterval);
        autoInterval = null;
        await sock.sendMessage(chatId, { text: 'Auto-send *stopped*.' });
      } else {
        await sock.sendMessage(chatId, { text: 'Auto-send not running.' });
      }
      return;
    }

    // Default: process once
    const result = await processQueue(sock);
    const outbox = readOutbox();
    const pending = outbox.filter(m => !m.sent).length;

    await sock.sendMessage(chatId, {
      text: '*QUEUE PROCESSED*\nSent: *' + result.processed + '*\nFailed: *' + result.failed + '*\nPending: *' + pending + '*\n\nUse .sendoutbox auto for auto-send'
    });

  } catch (error) {
    console.error('[OUTBOX CMD] Error:', error);
    await sock.sendMessage(chatId, { text: 'Error: ' + error.message });
  }
}

module.exports = sendoutboxCommand;
