'use strict';

// Runs inside a pure-Node child process (spawned by Electron main with
// ELECTRON_RUN_AS_NODE=1). This avoids Electron main-process module patching
// that breaks the WebSocket (ws) used by whatsapp-web.js/puppeteer.

const { runAutomation } = require('./core');

let stopToken = null;

process.on('message', (msg) => {
  if (!msg) return;

  if (msg.type === 'stop') {
    if (stopToken) stopToken.stopped = true;
    return;
  }

  if (msg.type !== 'start') return;

  stopToken = { stopped: false };

  const ui = {
    onLog: (m) => process.send({ type: 'log', data: m }),
    onQR: (qr) => process.send({ type: 'qr', data: qr }),
    onSummary: (s) => process.send({ type: 'summary', data: s }),
    askHotspotConfirmation: () =>
      new Promise((resolve) => {
        const onMsg = (m) => {
          if (m && m.type === 'confirm-hotspot') {
            process.removeListener('message', onMsg);
            resolve();
          }
        };
        process.on('message', onMsg);
        process.send({ type: 'prompt-hotspot' });
      }),
  };

  const payload = Object.assign({}, msg.payload || {}, { stopToken });

  runAutomation(payload, ui)
    .then(() => process.send({ type: 'done' }))
    .catch((err) => {
      process.send({
        type: 'error',
        data: (err && (err.stack || err.message)) || String(err),
      });
    });
});
