'use strict';

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {
  selectFile: () => ipcRenderer.invoke('select-file'),
  getConfig: () => ipcRenderer.invoke('get-config'),
  removeDuplicates: (excelPath) => ipcRenderer.invoke('remove-duplicates', excelPath),
  sortByColumn: (excelPath) => ipcRenderer.invoke('sort-by-column', excelPath),
  start: (payload) => ipcRenderer.send('start', payload),
  stop: () => ipcRenderer.send('stop'),
  confirmHotspot: () => ipcRenderer.send('confirm-hotspot'),
  onEvent: (callback) => ipcRenderer.on('automation-event', (_event, data) => callback(data)),
});
