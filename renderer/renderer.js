'use strict';

const els = {
  excelPath: document.getElementById('excelPath'),
  startRow: document.getElementById('startRow'),
  endRow: document.getElementById('endRow'),
  dryRun: document.getElementById('dryRun'),
  startBtn: document.getElementById('startBtn'),
  stopBtn: document.getElementById('stopBtn'),
  browseBtn: document.getElementById('browseBtn'),
  removeDupBtn: document.getElementById('removeDupBtn'),
  sortBtn: document.getElementById('sortBtn'),
  detectDuplicate: document.getElementById('detectDuplicate'),
  maintenancePassword: document.getElementById('maintenancePassword'),
  unlockBtn: document.getElementById('unlockBtn'),
  qrBox: document.getElementById('qrBox'),
  qrImg: document.getElementById('qrImg'),
  hotspotBox: document.getElementById('hotspotBox'),
  confirmBtn: document.getElementById('confirmBtn'),
  sidePanel: document.getElementById('sidePanel'),
  logPane: document.getElementById('logPane'),
  summaryBox: document.getElementById('summaryBox'),
  statusBar: document.getElementById('statusBar'),
};

function logLine(msg) {
  els.logPane.textContent += msg + '\n';
  els.logPane.scrollTop = els.logPane.scrollHeight;
}

function setStatus(text) {
  els.statusBar.textContent = text;
}

function updateStartEnabled() {
  els.startBtn.disabled = !els.excelPath.value.trim();
}

function hidePanels() {
  els.qrBox.classList.remove('visible');
  els.hotspotBox.classList.remove('visible');
  els.sidePanel.classList.remove('visible');
  els.summaryBox.classList.remove('visible');
  els.qrImg.style.display = 'none';
}

function renderSummary(s) {
  const rows = [
    ['Task Date', s.date],
    ['Total Excel Records', s.totalRecords],
    ['Manual Sending Scope', 'First ' + s.manualCount + ' entries'],
    ['Auto Sending Scope', 'From row ' + s.firstAutoRow + ', ' + s.autoScopeCount + ' entries'],
    ['Auto Sent Successfully', s.sent],
    ['Failed to Send', s.failed],
    ['Sent (no tick)', s.noTick],
    ['Invalid Number Format', s.invalid],
    ['Not on WhatsApp', s.notOnWhatsApp],
    ['Duplicate Numbers Skipped', s.duplicate],
    ['Missing message content', s.missingContent],
    ['Whats app number missing', s.missingNumber],
    ['Network connection fail', s.networkFail],
    ['Skipped (already completed)', s.skippedExisting],
  ];

  let html = '<table>';
  for (const [k, v] of rows) {
    html += `<tr><td>${k}</td><td>${v}</td></tr>`;
  }
  html += '</table>';
  els.summaryBox.innerHTML = html;
  els.summaryBox.classList.add('visible');
}

function handleEvent(data) {
  const { type, data: payload } = data || {};

  switch (type) {
    case 'log':
      logLine(payload);
      break;
    case 'qr':
      els.qrImg.src = payload;
      els.qrImg.style.display = 'block';
      els.qrBox.classList.add('visible');
      els.sidePanel.classList.add('visible');
      els.hotspotBox.classList.remove('visible');
      setStatus('Waiting for QR scan...');
      break;
    case 'prompt-hotspot':
      els.qrImg.style.display = 'none';
      els.qrBox.classList.remove('visible');
      els.hotspotBox.classList.add('visible');
      els.sidePanel.classList.add('visible');
      setStatus('Waiting for hotspot confirmation...');
      break;
    case 'summary':
      renderSummary(payload);
      setStatus('Summary ready.');
      break;
    case 'done':
      els.sidePanel.classList.remove('visible');
      setStatus('Finished.');
      els.startBtn.disabled = false;
      els.stopBtn.disabled = true;
      break;
    case 'error':
      els.sidePanel.classList.remove('visible');
      logLine('[ERROR] ' + payload);
      setStatus('Failed.');
      els.startBtn.disabled = false;
      els.stopBtn.disabled = true;
      break;
  }
}

window.api.onEvent(handleEvent);

(async function prefill() {
  try {
    const cfg = await window.api.getConfig();
    if (cfg && cfg.startRow && !els.startRow.value) els.startRow.value = cfg.startRow;
    if (cfg && cfg.endRow && !els.endRow.value) els.endRow.value = cfg.endRow;
  } catch (e) {
    // ignore
  }
  updateStartEnabled();
})();

els.browseBtn.addEventListener('click', async () => {
  const file = await window.api.selectFile();
  if (file) {
    els.excelPath.value = file;
    updateStartEnabled();
  }
});

els.excelPath.addEventListener('input', updateStartEnabled);

els.removeDupBtn.addEventListener('click', async () => {
  const excelPath = els.excelPath.value.trim();
  if (!excelPath) {
    logLine('[ERROR] Please add the Excel file from browse.');
    setStatus('Missing Excel file.');
    return;
  }
  try {
    const res = await window.api.removeDuplicates(excelPath);
    logLine('[Duplicate] Removed ' + res.removed + ' duplicate row(s). ' + res.remaining + ' unique row(s) remain.');
    setStatus('Duplicate rows removed.');
  } catch (err) {
    logLine('[ERROR] ' + (err && err.message ? err.message : err));
    setStatus('Remove duplicates failed.');
  }
});

els.sortBtn.addEventListener('click', async () => {
  const excelPath = els.excelPath.value.trim();
  if (!excelPath) {
    logLine('[ERROR] Please add the Excel file from browse.');
    setStatus('Missing Excel file.');
    return;
  }
  try {
    const res = await window.api.sortByColumn(excelPath);
    logLine('[Sort] Sorted ' + res.sorted + ' row(s) by Customer Manager (A-Z).');
    setStatus('Sort completed.');
  } catch (err) {
    logLine('[ERROR] ' + (err && err.message ? err.message : err));
    setStatus('Sort failed.');
  }
});

els.unlockBtn.addEventListener('click', () => {
  if (els.maintenancePassword.value === 'mRzNEz007') {
    els.removeDupBtn.disabled = false;
    els.sortBtn.disabled = false;
    els.maintenancePassword.value = '';
    setStatus('Maintenance tools unlocked.');
    logLine('[Unlock] Remove duplicates and sort are now enabled.');
  } else {
    els.maintenancePassword.value = '';
    setStatus('Wrong password.');
    logLine('[ERROR] Wrong password. Tools remain locked.');
  }
});

els.startBtn.addEventListener('click', () => {
  hidePanels();
  els.logPane.textContent = '';

  const excelPath = els.excelPath.value.trim();
  const startVal = els.startRow.value.trim();
  const endVal = els.endRow.value.trim();

  if (!excelPath) {
    logLine('[ERROR] Please add the Excel file from browse.');
    setStatus('Missing Excel file.');
    return;
  }
  if (!startVal) {
    logLine('[ERROR] Please provide the Start row.');
    setStatus('Missing Start row.');
    return;
  }
  if (!endVal) {
    logLine('[ERROR] Please provide the End row.');
    setStatus('Missing End row.');
    return;
  }

  const startRow = Number(startVal);
  const endRow = Number(endVal);
  if (!Number.isInteger(startRow) || startRow < 2) {
    logLine('[ERROR] Start row must be a whole number of 2 or greater.');
    setStatus('Invalid Start row.');
    return;
  }
  if (!Number.isInteger(endRow) || endRow < 2) {
    logLine('[ERROR] End row must be a whole number of 2 or greater.');
    setStatus('Invalid End row.');
    return;
  }
  if (startRow > endRow) {
    logLine('[ERROR] Start row cannot be greater than End row.');
    setStatus('Invalid row range.');
    return;
  }

  els.startBtn.disabled = true;
  els.stopBtn.disabled = false;
  setStatus('Starting...');

  window.api.start({
    excelPath,
    startRow,
    endRow,
    dryRun: els.dryRun.checked,
    detectDuplicate: els.detectDuplicate.checked,
  });
});

els.stopBtn.addEventListener('click', () => {
  logLine('[Stop] Stop requested...');
  setStatus('Stopping...');
  els.stopBtn.disabled = true;
  window.api.stop();
});

// Keep the Start-End range to at most 30 rows (End <= Start + 29).
function applyRowRangeLimit() {
  const start = Number(els.startRow.value);
  if (Number.isInteger(start) && start >= 2) {
    els.endRow.max = String(start + 29);
    const end = Number(els.endRow.value);
    if (Number.isInteger(end) && end > start + 29) {
      els.endRow.value = String(start + 29);
    }
  } else {
    els.endRow.removeAttribute('max');
  }
}
els.startRow.addEventListener('input', applyRowRangeLimit);
els.endRow.addEventListener('input', applyRowRangeLimit);

els.confirmBtn.addEventListener('click', () => {
  els.hotspotBox.classList.remove('visible');
  window.api.confirmHotspot();
  setStatus('Continuing...');
});
