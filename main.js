const { app, BrowserWindow, Menu, shell, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const pdfParse = require('pdf-parse');
const mammoth = require('mammoth');
const XLSX = require('xlsx');
const JSZip = require('jszip');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle, ImageRun } = require('docx');

let mainWindow;
let pendingFile = null;


function extractFileArg(argv) {
  if (!Array.isArray(argv)) return null;
  for (let i = 1; i < argv.length; i++) {
    const arg = argv[i];
    if (!arg || arg.startsWith('--')) continue;
    if (arg.toLowerCase().endsWith('.exe')) continue;
    try { if (fs.existsSync(arg) && fs.statSync(arg).isFile()) return arg; } catch (e) {}
  }
  return null;
}


function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1366,
    height: 860,
    minWidth: 900,
    minHeight: 600,
    title: 'OfficePro 专业办公套件',
    icon: path.join(__dirname, 'icon.png'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      enableRemoteModule: false,
      preload: path.join(__dirname, 'preload.js')
    },
    frame: true,
    backgroundColor: '#f5f7fa'
  });

  mainWindow.loadFile('index.html', { encoding: 'UTF-8' });
  mainWindow.on('closed', () => { mainWindow = null; });

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });


  mainWindow.webContents.on('will-navigate', (event, navUrl) => {
    try {
      if (navUrl && navUrl.toLowerCase().startsWith('file:///')) {
        let fp = decodeURIComponent(navUrl.replace(/^file:\/\//, ''));
        fp = fp.replace(/\//g, path.sep);
        const ext = path.extname(fp).toLowerCase();
        const dragOk = ['.doc','.docx','.docm','.dotx','.wps','.wpt','.dot','.xls','.xlsx','.xlsm','.et','.csv','.tsv','.ods','.ppt','.pptx','.dps','.ppsx','.pdf','.txt','.md','.rtf','.odt','.odp','.epub','.mobi','.html','.htm','.json','.xml','.log','.sql','.jpg','.jpeg','.png','.bmp','.gif','.webp','.svg','.tif','.tiff'];
        if (ext && dragOk.includes(ext) && fs.existsSync(fp)) {
          event.preventDefault();
          openFileByPath(fp);
          return;
        }
      }
    } catch (e) {}
    event.preventDefault();
  });


  if (pendingFile) {
    mainWindow.webContents.once('did-finish-load', () => {
      openFileByPath(pendingFile);
      pendingFile = null;
    });
  }
}

const { parseFile, getFileType, escapeHtml, decodeTextBuffer, extractOleText } = require('./file-parser');


ipcMain.handle('open-file-dialog', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '打开文件',
    properties: ['openFile'],
    filters: [
      { name: '所有支持的文档', extensions: ['txt','md','markdown','mdx','rst','csv','tsv','json','jsonl','xml','yaml','yml','toml','ini','conf','html','htm','xhtml','mht','mhtml','rtf','log','sql','doc','docx','docm','dotx','dotm','wps','wpt','dot','odt','ott','pages','xls','xlsx','xlsm','xlsb','xltx','et','ods','ots','numbers','dif','ppt','pptx','pptm','ppsx','potx','dps','odp','key','pdf','xps','epub','mobi','azw3','fb2','eml','msg','zip','jar','tar','gz','7z','rar','xmind','vsdx','ttf','otf','woff','woff2','jpg','jpeg','png','bmp','gif','webp','tiff','tif','svg','ico','avif','heic','mp3','wav','flac','aac','ogg','m4a','mp4','mkv','avi','mov','webm','js','ts','py','java','c','cpp','h','cs','go','rs','rb','php','css','sh','bat','ps1','vue'] },
      { name: 'Word/文字文档', extensions: ['doc','docx','docm','dotx','dotm','wps','wpt','dot','odt','ott','rtf','pages'] },
      { name: 'Excel/表格', extensions: ['xls','xlsx','xlsm','xlsb','xlt','xltx','et','ods','ots','numbers','csv','tsv','dif','prn'] },
      { name: 'PPT/演示', extensions: ['ppt','pptx','pptm','pps','ppsx','pot','potx','dps','dpt','odp','key'] },
      { name: 'PDF与电子书', extensions: ['pdf','xps','epub','mobi','azw3','fb2'] },
      { name: '文本与源代码', extensions: ['txt','md','json','xml','yaml','yml','toml','ini','conf','log','sql','html','htm','css','js','ts','py','java','c','cpp','h','cs','go','rs','rb','php','sh','bat','ps1','vue'] },
      { name: '图片', extensions: ['jpg','jpeg','png','bmp','gif','webp','tiff','tif','svg','ico','avif','heic'] },
      { name: '音频与视频', extensions: ['mp3','wav','flac','aac','ogg','m4a','mp4','mkv','avi','mov','webm','wmv'] },
      { name: '压缩包/邮件/字体', extensions: ['zip','jar','tar','gz','7z','rar','xmind','eml','msg','ttf','otf','woff','woff2'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });
  if (result.canceled || result.filePaths.length === 0) {
    return { success: false, error: '用户取消选择' };
  }
  const filePath = result.filePaths[0];
  try {
    const parsed = await parseFile(filePath);
    return { success: true, filePath, ...parsed };
  } catch (e) {
    return { success: false, error: '解析失败：' + (e && e.message ? e.message : e) };
  }
});


async function openFileByPath(filePath) {
  if (!mainWindow) return;
  if (!fs.existsSync(filePath)) return;
  try {
    const parsed = await parseFile(filePath);
    mainWindow.webContents.send('file-opened', { filePath, ...parsed });
  } catch (e) {
    mainWindow.webContents.send('file-opened', { filePath, type: 'error', error: '解析失败：' + (e && e.message ? e.message : e), fileName: path.basename(filePath) });
  }
}

ipcMain.handle('open-file-by-path', async (event, filePath) => {
  if (!fs.existsSync(filePath)) return { success: false, error: '文件不存在' };
  try {
    const parsed = await parseFile(filePath);
    return { success: true, filePath, ...parsed };
  } catch (e) {
    return { success: false, error: '解析失败：' + (e && e.message ? e.message : e) };
  }
});


let ocrWorker = null;

async function getOcrWorker() {
  if (ocrWorker) return ocrWorker;
  const { createWorker } = require('tesseract.js');
  const tessdataPath = path.join(__dirname, 'tessdata');
  ocrWorker = await createWorker('chi_sim', 1, {
    logger: () => {},
    langPath: tessdataPath,
    gzip: true
  });
  return ocrWorker;
}

ipcMain.handle('ocr-image', async (event, imageInput) => {
  try {
    const worker = await getOcrWorker();
    let imageData = imageInput;

    if (typeof imageInput === 'string' && fs.existsSync(imageInput)) {
      imageData = fs.readFileSync(imageInput);
    }
    const { data: { text, confidence } } = await worker.recognize(imageData);
    return { success: true, text, confidence };
  } catch(e) {
    return { success: false, error: e.message };
  }
});


ipcMain.handle('ocr-batch', async (event, imageInputs) => {
  try {
    const worker = await getOcrWorker();
    const results = [];
    for (const input of imageInputs) {
      let imageData = input;
      if (typeof input === 'string' && fs.existsSync(input)) {
        imageData = fs.readFileSync(input);
      }
      const { data } = await worker.recognize(imageData);
      results.push({ text: data.text, confidence: data.confidence });
    }
    return { success: true, results };
  } catch(e) {
    return { success: false, error: e.message };
  }
});


ipcMain.handle('pdf-to-word', async (event, filePath) => {
  try {
    const dataBuffer = fs.readFileSync(filePath);
    const data = await pdfParse(dataBuffer);
    const text = data.text;
    const lines = text.split('\n').filter(l => l.trim());
    const children = [];
    let currentPara = [];
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) {
        if (currentPara.length > 0) {
          children.push(new Paragraph({ children: [new TextRun({ text: currentPara.join(' '), size: 24 })], spacing: { after: 200 } }));
          currentPara = [];
        }
        continue;
      }
      if (line.length < 50 && (line === line.toUpperCase() || /^\d+[\.、]/.test(line))) {
        if (currentPara.length > 0) {
          children.push(new Paragraph({ children: [new TextRun({ text: currentPara.join(' '), size: 24 })], spacing: { after: 200 } }));
          currentPara = [];
        }
        children.push(new Paragraph({ children: [new TextRun({ text: line, bold: true, size: 32 })], heading: HeadingLevel.HEADING_1, spacing: { before: 300, after: 200 } }));
      } else {
        currentPara.push(line);
      }
    }
    if (currentPara.length > 0) {
      children.push(new Paragraph({ children: [new TextRun({ text: currentPara.join(' '), size: 24 })], spacing: { after: 200 } }));
    }
    const doc = new Document({ sections: [{ properties: { page: { margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } } }, children: children }] });
    const buffer = await Packer.toBuffer(doc);
    const savePath = dialog.showSaveDialogSync(mainWindow, { title: '保存Word文档', defaultPath: path.basename(filePath, '.pdf') + '.docx', filters: [{ name: 'Word文档', extensions: ['docx'] }] });
    if (savePath) {
      fs.writeFileSync(savePath, buffer);
      return { success: true, text: text, pages: data.numpages, savePath: savePath };
    }
    return { success: false, error: '用户取消保存' };
  } catch (error) {
    return { success: false, error: error.message };
  }
});


ipcMain.handle('pdf-to-excel', async (event, filePath) => {
  try {
    const dataBuffer = fs.readFileSync(filePath);
    const data = await pdfParse(dataBuffer);
    const text = data.text;
    const lines = text.split('\n');
    let csvContent = '';
    for (const line of lines) {
      if (line.trim()) {
        const cells = line.split(/\s{2,}|\t/).filter(c => c.trim());
        if (cells.length > 1) {
          csvContent += cells.map(c => '"' + c.replace(/"/g, '""') + '"').join(',') + '\n';
        } else {
          csvContent += '"' + line.trim().replace(/"/g, '""') + '"\n';
        }
      }
    }
    const savePath = dialog.showSaveDialogSync(mainWindow, { title: '保存Excel文档', defaultPath: path.basename(filePath, '.pdf') + '.csv', filters: [{ name: 'CSV文件', extensions: ['csv'] }] });
    if (savePath) {
      fs.writeFileSync(savePath, '\ufeff' + csvContent, 'utf8');
      return { success: true, text: text, pages: data.numpages, savePath: savePath };
    }
    return { success: false, error: '用户取消保存' };
  } catch (error) {
    return { success: false, error: error.message };
  }
});


ipcMain.handle('pdf-to-ppt', async (event, filePath) => {
  try {
    const dataBuffer = fs.readFileSync(filePath);
    const data = await pdfParse(data);
    const text = data.text;
    const pages = text.split(/\f/);
    let htmlContent = '';
    pages.forEach((page, index) => {
      const lines = page.split('\n').filter(l => l.trim());
      if (lines.length > 0) {
        const title = lines[0].trim().substring(0, 50);
        const content = lines.slice(1).join('\n').substring(0, 500);
        htmlContent += `<div class="slide-page" data-slide="${index+1}"><h1>${title}</h1><pre>${content}</pre></div>\n`;
      }
    });
    const savePath = dialog.showSaveDialogSync(mainWindow, { title: '保存PPT内容', defaultPath: path.basename(filePath, '.pdf') + '_ppt.html', filters: [{ name: 'HTML文件', extensions: ['html'] }] });
    if (savePath) {
      fs.writeFileSync(savePath, htmlContent, 'utf8');
      return { success: true, text: text, pages: data.numpages, savePath: savePath };
    }
    return { success: false, error: '用户取消保存' };
  } catch (error) {
    return { success: false, error: error.message };
  }
});


ipcMain.handle('word-to-pdf', async (event, htmlContent) => {
  try {
    const tempHtml = path.join(app.getPath('temp'), 'officepro_word_to_pdf.html');
    const fullHtml = `<!DOCTYPE html><html><head><meta charset="UTF-8"><style>
      body { font-family: 'Microsoft YaHei', SimSun, sans-serif; margin: 40px; line-height: 1.8; font-size: 14px; }
      h1 { font-size: 24px; font-weight: bold; margin: 20px 0 10px; }
      h2 { font-size: 20px font-weight: bold; margin: 16px 0 8px; }
      h3 { font-size: 16px; font-weight: bold; margin: 12px 0 6px; }
      p { margin: 8px 0; }
      table { border-collapse: collapse; width: 100%; margin: 10px 0; }
      td, th { border: 1px solid #333; padding: 6px 10px; }
      img { max-width: 100%; }
      @page { size: A4; margin: 20mm; }
    </style></head><body>${htmlContent}</body></html>`;
    fs.writeFileSync(tempHtml, fullHtml, 'utf8');
    const printWindow = new BrowserWindow({ show: false });
    await printWindow.loadFile(tempHtml);
    const pdfBuffer = await printWindow.webContents.printToPDF({ printBackground: true, pageSize: 'A4', margins: { top: 20, bottom: 20, left: 20, right: 20 } });
    printWindow.close();
    const savePath = dialog.showSaveDialogSync(mainWindow, { title: '保存PDF文档', defaultPath: 'document.pdf', filters: [{ name: 'PDF文档', extensions: ['pdf'] }] });
    if (savePath) {
      fs.writeFileSync(savePath, pdfBuffer);
      try { fs.unlinkSync(tempHtml); } catch(e) {}
      return { success: true, savePath: savePath };
    }
    try { fs.unlinkSync(tempHtml); } catch(e) {}
    return { success: false, error: '用户取消保存' };
  } catch (error) {
    return { success: false, error: error.message };
  }
});


ipcMain.handle('select-pdf-file', async () => {
  const result = dialog.showOpenDialogSync(mainWindow, {
    title: '选择PDF文件',
    filters: [{ name: 'PDF文档', extensions: ['pdf'] }],
    properties: ['openFile']
  });
  if (result && result.length > 0) {
    return { success: true, filePath: result[0], fileName: path.basename(result[0]), fileSize: fs.statSync(result[0]).size };
  }
  return { success: false, error: '用户取消选择' };
});


function createMenu() {
  const template = [
    {
      label: '文件',
      submenu: [
        { label: '新建', accelerator: 'Ctrl+N', click: () => { mainWindow.webContents.send('new'); } },
        { label: '打开文件', accelerator: 'Ctrl+O', click: () => { mainWindow.webContents.send('open-file-dialog'); } },
        { label: '保存', accelerator: 'Ctrl+S', click: () => { mainWindow.webContents.send('save'); } },
        { type: 'separator' },
        { label: '导出PDF', accelerator: 'Ctrl+E', click: () => { mainWindow.webContents.send('pdf'); } },
        { label: '打印', accelerator: 'Ctrl+P', click: () => { mainWindow.webContents.print(); } },
        { type: 'separator' },
        { label: '退出', accelerator: 'Ctrl+Q', click: () => { app.quit(); } }
      ]
    },
    {
      label: '编辑',
      submenu: [
        { label: '撤销', accelerator: 'Ctrl+Z', role: 'undo' },
        { label: '重做', accelerator: 'Ctrl+Y', role: 'redo' },
        { type: 'separator' },
        { label: '剪切', accelerator: 'Ctrl+X', role: 'cut' },
        { label: '复制', accelerator: 'Ctrl+C', role: 'copy' },
        { label: '粘贴', accelerator: 'Ctrl+V', role: 'paste' },
        { label: '全选', accelerator: 'Ctrl+A', role: 'selectAll' },
        { type: 'separator' },
        { label: '查找替换', accelerator: 'Ctrl+F', click: () => { mainWindow.webContents.send('find'); } }
      ]
    },
    {
      label: '视图',
      submenu: [
        { label: '放大', accelerator: 'Ctrl+=', role: 'zoomIn' },
        { label: '缩小', accelerator: 'Ctrl+-', role: 'zoomOut' },
        { label: '重置缩放', accelerator: 'Ctrl+0', role: 'resetZoom' },
        { type: 'separator' },
        { label: '全屏', accelerator: 'F11', role: 'togglefullscreen' },
        { label: '开发者工具', accelerator: 'F12', role: 'toggleDevTools' }
      ]
    },
    {
      label: '帮助',
      submenu: [
        { label: '操作指引', click: () => { mainWindow.webContents.send('guide'); } },
        { type: 'separator' },
        { label: '关于 OfficePro', click: () => { mainWindow.webContents.send('about'); } }
      ]
    }
  ];
  Menu.setApplicationMenu(Menu.buildFromTemplate(template));
}


app.whenReady().then(() => {

  const firstFile = extractFileArg(process.argv);
  if (firstFile) pendingFile = firstFile;
  createWindow();
  createMenu();
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});


app.on('open-file', (event, filePath) => {
  event.preventDefault();
  if (mainWindow) {
    openFileByPath(filePath);
  } else {
    pendingFile = filePath;
  }
});


const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  app.quit();
} else {
  app.on('second-instance', (event, commandLine) => {

    const fileArg = extractFileArg(commandLine);
    if (fileArg && mainWindow) {
      openFileByPath(fileArg);
    }
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });
}


app.on('before-quit', async () => {
  if (ocrWorker) {
    try { await ocrWorker.terminate(); } catch(e) {}
    ocrWorker = null;
  }
});
