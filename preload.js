const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("electronAPI", {
  uploadFile: () => ipcRenderer.invoke("upload-file"),
  processFile: (data) =>
    ipcRenderer.invoke("process-file", data),

  saveOutput: (filePath) =>
    ipcRenderer.invoke("save-output", filePath),
  
  chooseOutputPath: () => ipcRenderer.invoke("choose-output-path"),
  getFilePath: (file) => webUtils.getPathForFile(file)
});