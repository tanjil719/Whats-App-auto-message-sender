'use strict';

const fs = require('fs');
const path = require('path');
const ExcelJS = require('exceljs');
const { isExcelFileOpen } = require('./excel-ops');

const STATUS = {
  SENT: 'Sent Successfully',
  FAILED: 'Failed to send',
  INVALID: 'Invalid Number Format',
  NOT_ON_WHATSAPP: 'Not on WhatsApp',
  DUPLICATE: 'Duplicate Number Skipped',
  MISSING_CONTENT: 'Missing message content',
  MISSING_NUMBER: 'Whats app number missing',
  NETWORK_FAIL: 'Network connection fail',
  NO_TICK: 'Sent (no tick)',
};

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// Returns international digits without '+' (e.g. "8801623892013") or null if invalid.
function normalizeBangladeshiNumber(raw) {
  if (raw == null) return null;
  const digits = String(raw).replace(/\D/g, '');

  let local = null;
  if (digits.length === 13 && digits.startsWith('880')) {
    local = '0' + digits.slice(3);
  } else if (digits.length === 11 && digits.startsWith('0')) {
    local = digits;
  } else if (digits.length === 10 && digits.startsWith('1')) {
    local = '0' + digits;
  } else {
    return null;
  }

  if (!/^01[3-9]\d{8}$/.test(local)) return null;
  return '880' + local.slice(1);
}

function cellText(cell) {
  if (!cell) return '';
  const v = cell.value;
  if (v == null) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString();
  if (typeof v === 'object') {
    if (Array.isArray(v.richText)) return v.richText.map((r) => r.text || '').join('');
    if (v.text != null) return String(v.text);
    if (v.result != null) return String(v.result);
    if (v.hyperlink) return String(v.hyperlink);
    return '';
  }
  return String(v);
}

function isBlank(value) {
  return value == null || String(value).trim() === '';
}

/**
 * Run the automation.
 *
 * @param {object} overrides  { excelPath?, startRow?, endRow?, dryRun? }
 * @param {object} ui         { onLog?, onQR?, onSummary?, askHotspotConfirmation? }
 * @returns {Promise<object|null>} summary object, or null when nothing to do
 */
async function runAutomation(overrides = {}, ui = {}) {
  const log = typeof ui.onLog === 'function' ? ui.onLog : () => {};
  const onQR = typeof ui.onQR === 'function' ? ui.onQR : () => {};
  const onSummary = typeof ui.onSummary === 'function' ? ui.onSummary : () => {};
  const askHotspotConfirmation = typeof ui.askHotspotConfirmation === 'function'
    ? ui.askHotspotConfirmation
    : () => Promise.resolve();

  const CONFIG_PATH = path.join(__dirname, 'config.json');
  if (!fs.existsSync(CONFIG_PATH)) {
    throw new Error('config.json not found: ' + CONFIG_PATH);
  }
  const CONFIG = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));

  const {
    excel_path: EXCEL_PATH_DEFAULT,
    sheet_name: SHEET_NAME = null,
    columns: COLUMNS = {},
    auto_start_record: AUTO_START_RECORD = 201,
    start_row: START_ROW = null,
    end_row: END_ROW = null,
    delay_min_sec: DELAY_MIN = 90,
    delay_max_sec: DELAY_MAX = 180,
    batch_rows: BATCH_ROWS = 30,
    ack_timeout_sec: ACK_TIMEOUT_SEC = 60,
    reconnect_poll_sec: RECONNECT_POLL_SEC = 10,
    reconnect_max_wait_sec: RECONNECT_MAX_WAIT_SEC = 300,
    connect_timeout_sec: CONNECT_TIMEOUT_SEC = 600,
    ready_fallback_sec: READY_FALLBACK_SEC = 120,
    protocol_timeout_sec: PROTOCOL_TIMEOUT_SEC = 600,
  } = CONFIG;

  const EXCEL_PATH = overrides.excelPath || EXCEL_PATH_DEFAULT;
  const DRY_RUN = !!overrides.dryRun;
  const WORK_DIR = overrides.workDir || __dirname;
  const CHROME_PATH = overrides.chromePath || null;
  const stopToken = overrides.stopToken || null;
  const isStopRequested = () => !!(stopToken && stopToken.stopped);
  const DETECT_DUPLICATE = overrides.detectDuplicate !== false; // default true

  if (!EXCEL_PATH || !fs.existsSync(EXCEL_PATH)) {
    throw new Error('Excel file not found: ' + EXCEL_PATH);
  }

  if (isExcelFileOpen(EXCEL_PATH)) {
    throw new Error('The Excel file is currently open. Please close it and try again.');
  }

  let clientConnected = false;
  const SESSION_DIR = path.join(WORK_DIR, '.wwebjs_auth', 'session-bd-excel-sender');

  function createClient() {
    const { Client, LocalAuth } = require('whatsapp-web.js');
    const puppeteerOptions = {
      headless: true,
      protocolTimeout: PROTOCOL_TIMEOUT_SEC * 1000,
    };
    if (CHROME_PATH) puppeteerOptions.executablePath = CHROME_PATH;
    return new Client({
      authStrategy: new LocalAuth({
        clientId: 'bd-excel-sender',
        dataPath: path.join(WORK_DIR, '.wwebjs_auth'),
      }),
      authTimeoutMs: 180000,
      puppeteer: puppeteerOptions,
    });
  }

  async function waitForPageReady(client, timeoutMs) {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      try {
        const ready = await client.pupPage.evaluate(() => {
          return typeof window.WWebJS !== 'undefined' && window.WWebJS !== null;
        });
        if (ready) return true;
      } catch (err) {
        // page may be navigating; retry
      }
      await sleep(1000);
    }
    return false;
  }

  async function waitForSentAck(client, chatId, timeoutMs) {
    const deadline = Date.now() + timeoutMs;
    while (Date.now() < deadline) {
      try {
        const result = await client.pupPage.evaluate(async (id) => {
          const chat = await window.WWebJS.getChat(id, { getAsModel: false });
          if (!chat) return null;
          const msgs = chat.msgs.getModelsArray();
          const last = msgs[msgs.length - 1];
          if (!last || !last.id) return null;
          return { fromMe: last.id.fromMe, ack: last.ack };
        }, chatId);

        if (result && result.fromMe && result.ack >= 1) {
          return result.ack;
        }
      } catch (err) {
        // transient; retry on next loop
      }
      await sleep(1000);
    }
    return 0;
  }

  function cleanupStaleLock() {
    const lockFile = path.join(SESSION_DIR, 'DevToolsActivePort');
    try {
      if (fs.existsSync(lockFile)) {
        fs.unlinkSync(lockFile);
        log('[Cleanup] Removed stale lock file.');
      }
    } catch (err) {
      // ignore
    }
  }

  function resetSession() {
    try {
      fs.rmSync(SESSION_DIR, { recursive: true, force: true });
    } catch (e) {
      // ignore
    }
  }

  async function destroyAndReset(client) {
    try {
      await client.destroy();
    } catch (e) {
      // ignore
    }
    await sleep(2000);
    resetSession();
  }

  function connectWhatsApp(retryCount = 0) {
    cleanupStaleLock();
    clientConnected = false;
    const client = createClient();

    return new Promise((resolve, reject) => {
      let settled = false;
      let timer;
      let readyFallbackTimer;

      const finish = (fn, val) => {
        if (settled) return;
        settled = true;
        clearTimeout(timer);
        clearTimeout(readyFallbackTimer);
        fn(val);
      };

      timer = setTimeout(() => {
        finish(reject, new Error(`WhatsApp connection timed out after ${CONNECT_TIMEOUT_SEC}s (no ready received).`));
      }, CONNECT_TIMEOUT_SEC * 1000);

      client.on('loading_screen', (percent, message) => {
        log(`[WhatsApp] Loading... ${percent}%${message ? ' - ' + message : ''}`);
      });

      client.on('qr', (qr) => {
        log('[WhatsApp] Scan this QR code with your phone (WhatsApp -> Settings -> Linked Devices -> Link a Device).');
        onQR(qr);
      });

      client.on('authenticated', () => {
        log('[WhatsApp] Authenticated. Starting session...');
        // whatsapp-web.js sometimes never emits 'ready' after QR auth (a page-navigation
        // race). If 'ready' doesn't arrive shortly, reconnect: the session is already
        // persisted in the Chrome profile, so the reconnect reaches 'ready' without QR.
        if (!readyFallbackTimer) {
          readyFallbackTimer = setTimeout(async () => {
            if (retryCount >= 2) {
              log('[WhatsApp] ready still not received; proceeding anyway.');
              clientConnected = true;
              finish(resolve, client);
              return;
            }
            log('[WhatsApp] ready event not received; reconnecting with saved session...');
            clearTimeout(timer);
            settled = true;
            clearTimeout(readyFallbackTimer);
            try { await client.destroy(); } catch (e) {}
            try {
              const nextClient = await connectWhatsApp(retryCount + 1);
              resolve(nextClient);
            } catch (e) {
              reject(e);
            }
          }, READY_FALLBACK_SEC * 1000);
        }
      });

      client.on('ready', () => {
        log('[WhatsApp] Connected and ready.');
        clientConnected = true;
        finish(resolve, client);
      });

      client.on('auth_failure', (msg) => {
        finish(reject, new Error('WhatsApp auth failure: ' + msg));
      });

      client.on('disconnected', (reason) => {
        log('[WhatsApp] Disconnected: ' + reason);
        clientConnected = false;
      });

      client.initialize().catch((err) => finish(reject, err));
    });
  }

  async function ensureConnected() {
    if (clientConnected) return true;
    const maxMs = RECONNECT_MAX_WAIT_SEC * 1000;
    const start = Date.now();
    while (Date.now() - start < maxMs) {
      if (clientConnected) return true;
      const waitedSec = Math.round((Date.now() - start) / 1000);
      log(`[Network] Connection lost. Waiting for reconnect... (${waitedSec}s / ${RECONNECT_MAX_WAIT_SEC}s)`);
      await sleep(RECONNECT_POLL_SEC * 1000);
    }
    return clientConnected;
  }

  async function sendWhatsApp(client, phone, content) {
    const chatId = phone + '@c.us';

    // Ensure the WhatsApp page is fully injected before sending.
    await waitForPageReady(client, 30000);

    try {
      await client.sendMessage(chatId, content);
    } catch (err) {
      const msg = (err && err.message) ? String(err.message) : String(err);
      // "No LID for user" means the number is not registered on WhatsApp.
      if (/no lid/i.test(msg)) return STATUS.NOT_ON_WHATSAPP;
      log('[send fail] ' + phone + ' ' + msg);
      return clientConnected ? STATUS.FAILED : STATUS.NETWORK_FAIL;
    }

    // Confirm via the actual single-tick (ground truth), regardless of
    // sendMessage's return value. This avoids false "Failed to send" when the
    // message was actually delivered.
    try {
      const ack = await waitForSentAck(client, chatId, ACK_TIMEOUT_SEC * 1000);
      if (ack >= 1) return STATUS.SENT;

      // No tick. Distinguish "not on WhatsApp" from "sent but no tick".
      try {
        const contact = await client.getNumberId(phone);
        return contact ? STATUS.NO_TICK : STATUS.NOT_ON_WHATSAPP;
      } catch (err) {
        return STATUS.NO_TICK;
      }
    } catch (err) {
      log('[ack warn] ' + phone + ' ' + (err && err.message ? err.message : err));
      return STATUS.FAILED;
    }
  }

  function todayStr() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  }

  function randInt(min, max) {
    return Math.floor(min + Math.random() * (max - min + 1));
  }

  async function humanDelay() {
    const sec = randInt(DELAY_MIN, DELAY_MAX);
    const deadline = Date.now() + sec * 1000;
    while (Date.now() < deadline) {
      if (isStopRequested()) return false;
      await sleep(1000);
    }
    return true;
  }

  function buildSummary(counters, totalRecords, autoScopeCount, firstAutoRow) {
    const d = new Date();
    const dateStr = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    const manualCount = Math.min(AUTO_START_RECORD - 1, totalRecords);

    return {
      date: dateStr,
      totalRecords,
      manualCount,
      firstAutoRow,
      autoScopeCount,
      sent: counters.sent,
      failed: counters.failed,
      invalid: counters.invalid,
      notOnWhatsApp: counters.notOnWhatsApp,
      noTick: counters.noTick,
      duplicate: counters.duplicate,
      missingContent: counters.missingContent,
      missingNumber: counters.missingNumber,
      networkFail: counters.networkFail,
      skippedExisting: counters.skippedExisting,
    };
  }

  log('Excel file: ' + EXCEL_PATH);
  log('Mode: ' + (DRY_RUN ? 'DRY-RUN (no send, no write)' : 'LIVE (sending)'));

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(EXCEL_PATH);

  const ws = SHEET_NAME ? workbook.getWorksheet(SHEET_NAME) : workbook.worksheets[0];
  if (!ws) {
    throw new Error('Worksheet not found. Available sheets: ' +
      workbook.worksheets.map((w) => w.name).join(', '));
  }

  const headerRow = ws.getRow(1);
  const headerMap = {};
  headerRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const name = cellText(cell).trim();
    if (name) headerMap[name] = colNumber;
  });

  const colPhone = headerMap[COLUMNS.phone];
  const colMessage = headerMap[COLUMNS.message];
  const colStatus = headerMap[COLUMNS.status];

  if (!colPhone || !colMessage || !colStatus) {
    throw new Error('Could not find required columns. Expected: ' +
      [COLUMNS.phone, COLUMNS.message, COLUMNS.status].join(', ') +
      '. Found: ' + Object.keys(headerMap).join(', '));
  }

  const totalRecords = Math.max(0, ws.actualRowCount - 1);
  const startRowOverride = overrides.startRow ? Number(overrides.startRow) : null;
  const endRowOverride = overrides.endRow ? Number(overrides.endRow) : null;
  if (startRowOverride && endRowOverride && startRowOverride > endRowOverride) {
    throw new Error(`Invalid range: Start row (${startRowOverride}) is greater than End row (${endRowOverride}).`);
  }
  const firstAutoRow = startRowOverride || START_ROW || (AUTO_START_RECORD + 1);
  const lastRow = endRowOverride || (END_ROW ? Math.min(END_ROW, ws.actualRowCount) : ws.actualRowCount);
  const autoScopeCount = lastRow >= firstAutoRow ? (lastRow - firstAutoRow + 1) : 0;

  log(`Total data records : ${totalRecords}`);
  log(`Manual scope       : first ${AUTO_START_RECORD - 1} entries`);
  log(`Auto scope         : ${autoScopeCount} entries (row ${firstAutoRow} to ${lastRow})`);

  if (autoScopeCount <= 0) {
    log('Excel has 200 or fewer data records. No automatic operation will be performed.');
    return null;
  }

  const STATE_PATH = path.join(WORK_DIR, 'state.json');
  const today = todayStr();
  let state = { date: today, sentToday: 0 };
  if (fs.existsSync(STATE_PATH)) {
    try { state = JSON.parse(fs.readFileSync(STATE_PATH, 'utf8')); } catch (e) {}
  }
  if (state.date !== today) {
    state.date = today;
    state.sentToday = 0;
  }
  const saveState = () => fs.writeFileSync(STATE_PATH, JSON.stringify(state));

  const counters = {
    sent: 0,
    failed: 0,
    invalid: 0,
    notOnWhatsApp: 0,
    noTick: 0,
    duplicate: 0,
    missingContent: 0,
    missingNumber: 0,
    networkFail: 0,
    skippedExisting: 0,
  };
  const sentThisRun = new Set();
  const dupKey = (phone, content) => JSON.stringify([phone, content]);

  const saveWorkbook = async () => {
    await workbook.xlsx.writeFile(EXCEL_PATH);
  };

  let client = null;

  const handleRow = async (rowNumber) => {
    const row = ws.getRow(rowNumber);
    const phoneRaw = cellText(row.getCell(colPhone));
    const content = cellText(row.getCell(colMessage));

    const existingStatus = cellText(row.getCell(colStatus)).trim();
    if (!DRY_RUN && existingStatus && existingStatus !== STATUS.NETWORK_FAIL) {
      counters.skippedExisting++;
      log(`Row ${rowNumber} | Skipped (already: ${existingStatus})`);
      return;
    }

    let status;
    let networkFailed = false;
    if (isBlank(phoneRaw)) {
      status = STATUS.MISSING_NUMBER;
      counters.missingNumber++;
    } else if (isBlank(content)) {
      status = STATUS.MISSING_CONTENT;
      counters.missingContent++;
    } else {
      const phone = normalizeBangladeshiNumber(phoneRaw);
      if (!phone) {
        status = STATUS.INVALID;
        counters.invalid++;
      } else if (DETECT_DUPLICATE && sentThisRun.has(dupKey(phone, content))) {
        status = STATUS.DUPLICATE;
        counters.duplicate++;
      } else if (DRY_RUN) {
        status = STATUS.SENT;
        sentThisRun.add(dupKey(phone, content));
        counters.sent++;
      } else {
        if (!(await ensureConnected())) {
          status = STATUS.NETWORK_FAIL;
          counters.networkFail++;
          networkFailed = true;
        } else {
          status = await sendWhatsApp(client, phone, content);
          if (status === STATUS.SENT) {
            sentThisRun.add(dupKey(phone, content));
            counters.sent++;
          } else if (status === STATUS.NOT_ON_WHATSAPP) {
            counters.notOnWhatsApp++;
          } else if (status === STATUS.NETWORK_FAIL) {
            counters.networkFail++;
          } else if (status === STATUS.NO_TICK) {
            counters.noTick++;
          } else if (status === STATUS.FAILED) {
            counters.failed++;
          } else {
            counters.invalid++;
          }
          state.sentToday++;
        }
      }
    }

    if (!DRY_RUN) {
      row.getCell(colStatus).value = status;
      await saveWorkbook();
    }

    log(`Row ${rowNumber} | ${status}`);

    const didAttemptSend = status === STATUS.SENT || status === STATUS.NO_TICK || status === STATUS.FAILED;
    if (!DRY_RUN && didAttemptSend && rowNumber < lastRow) {
      const continued = await humanDelay();
      if (!continued) return 'STOPPED';
    }

    if (networkFailed) return 'NETWORK_FAIL';
    return 'DONE';
  };

  // ---- DRY RUN: single pass, no connection, no batching ----
  if (DRY_RUN) {
    for (let rowNumber = firstAutoRow; rowNumber <= lastRow; rowNumber++) {
      await handleRow(rowNumber);
    }
    const summary = buildSummary(counters, totalRecords, autoScopeCount, firstAutoRow);
    onSummary(summary);
    return summary;
  }

  // ---- LIVE: batched processing with account/session rotation ----
  let currentRow = firstAutoRow;
  if (state.nextRow && state.nextRow > firstAutoRow && state.nextRow <= lastRow) {
    currentRow = state.nextRow;
  }

  let shuttingDown = false;
  const shutdown = async () => {
    if (shuttingDown) return;
    shuttingDown = true;
    log('[Shutdown] Interrupted. Saving progress...');
    state.nextRow = currentRow > lastRow ? null : currentRow;
    saveState();
    try {
      await saveWorkbook();
      log('[Saved] Progress saved to Excel.');
    } catch (e) {}
    try {
      if (client) await client.destroy();
    } catch (e) {}
    onSummary(buildSummary(counters, totalRecords, autoScopeCount, firstAutoRow));
    process.exit(130);
  };
  process.on('SIGINT', shutdown);
  process.on('SIGTERM', shutdown);

  let stopped = false;

  while (currentRow <= lastRow && !stopped) {
    resetSession();
    client = await connectWhatsApp();

    let processed = 0;
    while (currentRow <= lastRow && processed < BATCH_ROWS && !stopped) {
      if (isStopRequested()) {
        stopped = true;
        break;
      }
      const result = await handleRow(currentRow);
      if (result === 'NETWORK_FAIL') {
        break;
      }
      currentRow++;
      processed++;
      state.nextRow = currentRow > lastRow ? null : currentRow;
      saveState();
      if (result === 'STOPPED') {
        stopped = true;
        break;
      }
    }

    state.nextRow = currentRow > lastRow ? null : currentRow;
    saveState();
    await saveWorkbook();

    if (stopped) {
      log('[Stop] Stop requested. Cleaning up...');
      if (client) { await client.destroy().catch(() => {}); client = null; }
      break;
    }

    if (currentRow <= lastRow) {
      await destroyAndReset(client);
      client = null;
      log(`[Batch done] Processed ${processed} rows. Next row: ${currentRow}.`);
      await askHotspotConfirmation();
      log('[Confirmed] Connecting next account...');
    } else {
      await destroyAndReset(client);
      client = null;
    }
  }

  await saveWorkbook();
  log('[Saved] All statuses written to Excel.');
  const summary = buildSummary(counters, totalRecords, autoScopeCount, firstAutoRow);
  onSummary(summary);
  return summary;
}

module.exports = {
  STATUS,
  normalizeBangladeshiNumber,
  runAutomation,
};
