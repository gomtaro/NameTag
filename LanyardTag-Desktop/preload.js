const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  printHtml: (html) => ipcRenderer.invoke('print-html', html),
});
