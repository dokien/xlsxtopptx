// Modules to control application life and create native browser window
const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const { readXlsx } = require('./lib/import-xlsx')
const { exportPptx } = require('./lib/export-pptx')
const path = require('path')
let mainWindow;
let xlsxData;
function createWindow() {
  // Create the browser window.
  mainWindow = new BrowserWindow({
    width: 1920,
    height: 800,
    fullscreen: true,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  })

  // and load the index.html of the app.
  mainWindow.loadFile('index.html')
  mainWindow.webContents.openDevTools()
}
app.whenReady().then(() => {
  createWindow()

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit()
})

ipcMain.handle('read-xlsx', async () => {
  const xlsxPath = dialog.showOpenDialogSync(mainWindow, {
    filters: [{ name: 'Excel File', extensions: ['xls', 'xlsx', 'xlsm'] }],
    properties: ['openFile']
  })
  if (xlsxPath && xlsxPath.length > 0) {
    try {
      xlsxData = await readXlsx(xlsxPath[0]);
      return xlsxData;
    } catch (e) {
      console.log(e);
      return false;
    }
  }
  return false;
})
ipcMain.handle('write-pptx', async () => {
  const pptxPath = dialog.showSaveDialogSync(mainWindow, {
    defaultPath: `${app.getPath('desktop')}/test.pptx`,
    filters: [{ name: 'PowerPoint File', extensions: ['ppt', 'pptx'] }],
  })
  if (pptxPath) {
    console.log(pptxPath);
    try {
      await exportPptx(pptxPath, xlsxData);
      return true;
    } catch (e) {
      console.log(e);
      return false;
    }
  }
  return false;
})
