// Cầu nối an toàn giữa trang và Electron: chỉ mở các lệnh tệp cần thiết, không lộ Node
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('xlDesktop', {
  platform: process.platform,
  startupFile: () => ipcRenderer.invoke('xl:startup-file'),
  openDialog: () => ipcRenderer.invoke('xl:open-dialog'),
  save: (o) => ipcRenderer.invoke('xl:save', o),
  onOpen: (cb) => ipcRenderer.on('xl:open', (_e, f) => cb(f)),
  onCloseRequest: (cb) => { ipcRenderer.on('xl:close-request', () => cb()); ipcRenderer.send('xl:close-ready'); },
  closeOk: () => ipcRenderer.send('xl:close-ok'),
  closePending: () => ipcRenderer.send('xl:close-pending'),
  closeCancel: () => ipcRenderer.send('xl:close-cancel'),
  setTitleBar: (o) => ipcRenderer.send('xl:titlebar', { color: String(o?.color ?? ''), symbolColor: String(o?.symbolColor ?? '') }),
  printers: () => ipcRenderer.invoke('xl:printers'),
  print: (o) => ipcRenderer.invoke('xl:print', o),
  pdf: (o) => ipcRenderer.invoke('xl:pdf', o),
  win: (a) => ipcRenderer.send('xl:win', String(a)),
  openExternal: (u) => ipcRenderer.send('xl:open-external', String(u)),
  license: () => ipcRenderer.invoke('xl:license'),
  activate: (k) => ipcRenderer.invoke('xl:activate', String(k)),
});
