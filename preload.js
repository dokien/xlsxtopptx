const { contextBridge, ipcRenderer } = require('electron')
contextBridge.exposeInMainWorld(
  'electron',
  {
    importExcel: () => ipcRenderer.invoke('read-xlsx'),
    exportPptx: () => ipcRenderer.invoke('write-pptx')
  }
)