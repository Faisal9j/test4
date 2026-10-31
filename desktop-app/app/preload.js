// الجسر بين المبنى والبرنامج: ما يفتح للصفحة إلا هذي الدوال.
const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('revenue', {
  snapshot: () => ipcRenderer.invoke('store:snapshot'),
  addChat: (m) => ipcRenderer.invoke('store:addChat', m),
  onChange: (fn) => { const h = (_e, c) => fn(c); ipcRenderer.on('store:change', h); return () => ipcRenderer.removeListener('store:change', h); },
  request: (text) => ipcRenderer.invoke('request', text),
  approve: (comment) => ipcRenderer.invoke('approve', comment),
  sendBack: (comment) => ipcRenderer.invoke('sendBack', comment),
  sample: (turns) => ipcRenderer.invoke('sample', turns),
  getSettings: () => ipcRenderer.invoke('settings:get'),
  saveSettings: (s) => ipcRenderer.invoke('settings:save', s),
  checkSettings: () => ipcRenderer.invoke('settings:check'),
  pickFolder: () => ipcRenderer.invoke('pick:folder'),
  pickFile: (base) => ipcRenderer.invoke('pick:file', base),
  openPath: (p) => ipcRenderer.invoke('open:path', p)
});
