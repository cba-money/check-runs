const { app, BrowserWindow, ipcMain, dialog } = require("electron");
const path = require("path");

const isDev = require("./util").isDev;

const { startProcess } = require("./checkruns/check-runs");

function createWindow() {
  const win = new BrowserWindow({
    width: 700,
    height: 500,
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true
    },
    nodeIntegration: false,
    contextIsolation: true,
    sandbox: true,
    webSecurity: true,
    devTools: isDev(),
  });

  // Load main window page
  win.loadFile(
    path.join(__dirname, "renderer", "index.html")
  );
}

async function selectLocalFile() {
  const focusedWindow = BrowserWindow.getFocusedWindow();
  
  const result = await dialog.showOpenDialog(focusedWindow, {
    properties: ['openFile'],
    filters: [{ name: 'Excel Spreadsheets', extensions: ['xlsx', 'csv', 'xlsm', 'xml'] }]
  });

  if (!result.canceled) {
    const filePath = result.filePaths[0]; // Contains the full absolute path string
    console.log('Selected file:', filePath);
    return filePath;
  }
}

async function selectLocalFolder(){
    const focusedWindow = BrowserWindow.getFocusedWindow();
   
    const result = await dialog.showOpenDialog(focusedWindow, {
      properties: ['openDirectory'],
    });
    
     if (result.canceled || result.filePaths.length === 0) {
        return null;
    }

    return result.filePaths[0];
}

app.whenReady().then(createWindow);

ipcMain.handle("upload-file", async (_) => {
    return await selectLocalFile();
});


ipcMain.handle("process-file", async (_, data) => {
  const output = await startProcess(data.registerFilePath, data.runFilePath);
  return output;
});


ipcMain.handle("save-output", async (_, sourcePath) => {
  const result = await dialog.showSaveDialog({
    defaultPath: "processed.xlsx"
  });

  if (result.canceled) {
    return null;
  }

  const fs = require("fs");

  fs.copyFileSync(sourcePath, result.filePath);

  return result.filePath;
});