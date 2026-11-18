'use strict';

const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { fork } = require('child_process');
const path = require('path');
const QRCode = require('qrcode');
const excelOps = require('./excel-ops');

let mainWindow = null;
let childProcess = null;
let running = false;

function send(type, data) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send('automation-event', { type, data });
  }
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 1080,
    title: 'WhatsApp Excel Sender',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));
}

function stopChild() {
  if (childProcess) {
    try {
      childProcess.kill();
    } catch (e) {}
    childProcess = null;
  }
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('before-quit', () => {
  stopChild();
});

ipcMain.handle('select-file', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Select Excel file',
    properties: ['openFile'],
    filters: [{ name: 'Excel files', extensions: ['xlsx', 'xls'] }],
  });
  return result.canceled ? null : result.filePaths[0];
});

ipcMain.handle('get-config', () => {
  const fs = require('fs');
  const configPath = path.join(__dirname, 'config.json');
  let config = {};
  try {
    config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
  } catch (e) {}
  return {
    excelPath: config.excel_path || '',
    startRow: config.start_row || '',
    endRow: config.end_row || '',
  };
});

ipcMain.handle('remove-duplicates', async (_event, excelPath) => {
  return excelOps.removeDuplicateRows(excelPath);
});

ipcMain.handle('sort-by-column', async (_event, excelPath) => {
  const fs = require('fs');
  const configPath = path.join(__dirname, 'config.json');
  let columnName = 'Customer Manager';
  try {
    const config = JSON.parse(fs.readFileSync(configPath, 'utf8'));
    if (config.sort_column) columnName = config.sort_column;
  } catch (e) {}
  return excelOps.sortByColumn(excelPath, columnName);
});

ipcMain.on('confirm-hotspot', () => {
  if (childProcess) {
    childProcess.send({ type: 'confirm-hotspot' });
  }
});

ipcMain.on('stop', () => {
  if (childProcess) {
    childProcess.send({ type: 'stop' });
  }
});

ipcMain.on('start', (_event, payload) => {
  if (running) {
    send('log', '[WARN] A run is already in progress.');
    return;
  }

  const excelPath = payload && payload.excelPath;
  if (!excelPath) {
    send('error', 'Please add the Excel file from browse.');
    return;
  }
  if (excelOps.isExcelFileOpen(excelPath)) {
    send('error', 'The Excel file is currently open. Please close it and try again.');
    return;
  }

  running = true;
  stopChild();

  const workDir = app.getPath('userData');
  const chromePath = app.isPackaged
    ? path.join(process.resourcesPath, 'chrome', 'chrome.exe')
    : undefined;

  const nodeExecPath = app.isPackaged
    ? path.join(process.resourcesPath, 'node.exe')
    : 'node';

  childProcess = fork(path.join(__dirname, 'automation-child.js'), [], {
    execPath: nodeExecPath,
    env: process.env,
    windowsHide: true,
  });

  if (childProcess.stdout) childProcess.stdout.on('data', () => {});
  if (childProcess.stderr) childProcess.stderr.on('data', () => {});

  childProcess.on('message', (msg) => {
    if (!msg) return;

    switch (msg.type) {
      case 'log':
        send('log', msg.data);
        break;
      case 'qr':
        QRCode.toDataURL(msg.data, { margin: 1, width: 300 })
          .then((url) => send('qr', url))
          .catch((err) => send('log', '[QR] Failed to render: ' + (err && err.message ? err.message : err)));
        break;
      case 'summary':
        send('summary', msg.data);
        break;
      case 'prompt-hotspot':
        send('prompt-hotspot', null);
        break;
      case 'done':
        send('done', null);
        running = false;
        childProcess = null;
        break;
      case 'error':
        send('error', msg.data);
        running = false;
        childProcess = null;
        break;
    }
  });

  childProcess.on('exit', (code) => {
    if (running && code !== 0) {
      send('error', 'Automation process exited unexpectedly (code ' + code + ').');
      running = false;
      childProcess = null;
    }
  });

  childProcess.send({
    type: 'start',
    payload: {
      ...(payload || {}),
      workDir,
      chromePath,
    },
  });
});
