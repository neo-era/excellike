// Cầu nối an toàn giữa giao diện và tiến trình chính (không lộ Node, không lộ khóa bí mật cho trang)
const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('km', {
  state: () => ipcRenderer.invoke('km:state'),
  importSigning: (text) => ipcRenderer.invoke('km:import-signing', text ?? null),
  exportSigning: () => ipcRenderer.invoke('km:export-signing'),
  create: (o) => ipcRenderer.invoke('km:create', o),
  update: (serial, patch) => ipcRenderer.invoke('km:update', serial, patch),
  verify: (text, machine) => ipcRenderer.invoke('km:verify', String(text), machine ? String(machine) : ''),
  copy: (t) => ipcRenderer.invoke('km:copy', String(t)),
  saveText: (name, text) => ipcRenderer.invoke('km:save-text', String(name), String(text)),
  restore: () => ipcRenderer.invoke('km:restore'),
  openData: () => ipcRenderer.invoke('km:open-data'),
});
