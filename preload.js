const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('officepro', {

  pdfToWord: (filePath) => ipcRenderer.invoke('pdf-to-word', filePath),
  pdfToExcel: (filePath) => ipcRenderer.invoke('pdf-to-excel', filePath),
  pdfToPPT: (filePath) => ipcRenderer.invoke('pdf-to-ppt', filePath),
  wordToPDF: (htmlContent) => ipcRenderer.invoke('word-to-pdf', htmlContent),
  selectPDFFile: () => ipcRenderer.invoke('select-pdf-file'),

  openFileDialog: () => ipcRenderer.invoke('open-file-dialog'),
  openFileByPath: (filePath) => ipcRenderer.invoke('open-file-by-path', filePath),

  ocrImage: (imageInput) => ipcRenderer.invoke('ocr-image', imageInput),
  ocrBatch: (imageInputs) => ipcRenderer.invoke('ocr-batch', imageInputs),

  platform: process.platform,
  versions: {
    node: process.versions.node,
    chrome: process.versions.chrome,
    electron: process.versions.electron
  },

  onFileOpened: (callback) => {
    const handler = (event, data) => callback(data);
    ipcRenderer.on('file-opened', handler);
  },

  onMenuAction: (callback) => {
    const acts = ['new','open-file-dialog','save','pdf','find','guide','about'];
    acts.forEach(a => ipcRenderer.on(a, () => callback(a)));
  }
});
