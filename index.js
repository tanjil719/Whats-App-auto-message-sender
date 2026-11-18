'use strict';

const { runAutomation } = require('./core');
const qrcode = require('qrcode-terminal');
const readline = require('readline');

// Run:  node index.js            -> real send
//       node index.js --dry-run  -> validate Excel + phone normalization only, no send, no write
const DRY_RUN = process.argv.includes('--dry-run');

function log(msg) {
  console.log(msg);
}

function askHotspotConfirmation() {
  return new Promise((resolve) => {
    const rl = readline.createInterface({ input: process.stdin, output: process.stdout });
    const ask = () => {
      rl.question('[Batch done] Did you connect the WiFi to the phone hotspot? Press Y to continue: ', (answer) => {
        if (String(answer).trim().toLowerCase() === 'y') {
          rl.close();
          resolve();
        } else {
          ask();
        }
      });
    };
    ask();
  });
}

function formatSummary(s) {
  const lines = [
    '',
    '====================================',
    '          TASK SUMMARY',
    '====================================',
    `Task Date: ${s.date}`,
    'Task Status: Completed',
    `Total Excel Records: ${s.totalRecords}`,
    `Manual Sending Scope: First ${s.manualCount} entries`,
    `Auto Sending Scope: Starting from row ${s.firstAutoRow}, ${s.autoScopeCount} entries in total`,
    `Auto Sent Successfully: ${s.sent}`,
    `Failed to Send: ${s.failed}`,
    `Sent (no tick): ${s.noTick}`,
    `Invalid Number Format: ${s.invalid}`,
    `Not on WhatsApp: ${s.notOnWhatsApp}`,
    `Duplicate Numbers Skipped: ${s.duplicate}`,
    `Missing message content: ${s.missingContent}`,
    `Whats app number missing: ${s.missingNumber}`,
    `Network connection fail: ${s.networkFail}`,
    `Skipped (already completed): ${s.skippedExisting}`,
    '====================================',
    '',
  ];
  return lines.join('\n');
}

runAutomation(
  { dryRun: DRY_RUN },
  {
    onLog: log,
    onQR: (qr) => {
      log('');
      log('[WhatsApp] Scan this QR code with your phone:');
      log('  WhatsApp -> Settings -> Linked Devices -> Link a Device');
      log('');
      qrcode.generate(qr, { small: true });
    },
    askHotspotConfirmation,
    onSummary: (s) => log(formatSummary(s)),
  },
)
  .then(() => process.exit(0))
  .catch((err) => {
    console.error('');
    console.error('[FATAL]', err && err.message ? err.message : err);
    console.error('Task Status: Failed');
    process.exit(1);
  });
