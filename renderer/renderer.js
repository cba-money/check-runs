let generatedFile = null;

function isFilePath(value) {
  if (typeof value !== "string" || !value.trim()) {
    return false;
  }

  const str = value.trim();

  // Windows absolute paths (C:\folder\file.txt, C:/folder/file.txt)
  const windowsAbsolute = /^[a-zA-Z]:[\\/]/;
  // UNC paths (\\server\share\file.txt)
  const uncPath = /^\\\\[^\\]+\\[^\\]+/;
  // Unix absolute paths (/home/user/file.txt, /etc/hosts)
  const unixAbsolute = /^\/(?:[^/]+\/?)*$/;
  // Relative paths containing directory separators
  const relativePath = /^(?:\.{1,2}[\\/]|[^<>:"|?*\r\n]+[\\/])/;
  // Paths ending with a filename extension
  const fileExtension = /(?:^|[\\/])[^\\/]+\.[a-zA-Z0-9_-]+$/;

  return (
    windowsAbsolute.test(str) ||
    uncPath.test(str) ||
    unixAbsolute.test(str) ||
    relativePath.test(str) ||
    fileExtension.test(str)
  );
}

function isFolderPath(value) {
  if (typeof value !== "string" || !value.trim()) {
    return false;
  }

  const str = value.trim();

  // Exclude URLs and URI schemes.
  if (/^[a-zA-Z][a-zA-Z\d+.-]*:\/\//.test(str)) {
    return false;
  }
  // Windows absolute paths (C:\Users\Tyler or C:/Users/Tyler).
  if (/^[a-zA-Z]:[\\/]/.test(str)) {
    return true;
  }
  // Windows UNC paths (\\server\share or \\server\share\folder).
  if (/^\\\\[^\\]+\\[^\\]+/.test(str)) {
    return true;
  }
  // Unix absolute paths (/home/user, /var/log).
  if (str.startsWith("/")) {
    return true;
  }
  // Relative paths (./folder, ../folder).
  if (/^(?:\.{1,2})[\\/]/.test(str)) {
    return true;
  }
  // Paths containing directory separators.
  if (/[/\\]/.test(str)) {
    return true;
  }
  // Bare folder names without extensions.
  if (/^[^<>:"/\\|?*\r\n.]+(?: [^<>:"/\\|?*\r\n.]+)*$/.test(str)) {
    return true;
  }

  return false;
}

async function uploadRegisterFile(){
  const file = await window.electronAPI.uploadFile();
  if(file === undefined || file === null || file.length <= 0){
      document.getElementById('uploadedRegisterFilePath').value = "Please select a folder...";
      return;
  }
  document.getElementById('uploadedRegisterFilePath').value = file;
}

async function uploadRunFile(){
  const file = await window.electronAPI.uploadFile();
  if(file === undefined || file === null || file.length <= 0){
      document.getElementById('uploadedRunFilePath').value = "Please select a folder...";
      return;
  }
  document.getElementById('uploadedRunFilePath').value = file;
}

async function setOutputPath(){
  const file = await window.electronAPI.chooseOutputPath();
  if(file === undefined || file === null || file.length <= 0){
      document.getElementById('outputPath').value = "Please select a folder...";
      return;
  }
  document.getElementById('outputPath').value = file;
}


document.getElementById('registerFileUpload').addEventListener('click', uploadRegisterFile);
document.getElementById('runFileUpload').addEventListener('click', uploadRunFile);
document.getElementById('outputPathUpload').addEventListener('click', setOutputPath);
document.getElementById('uploadedRunFilePath').addEventListener('click', uploadRunFile);
document.getElementById('uploadedRegisterFilePath').addEventListener('click', uploadRegisterFile);
document.getElementById('outputPath').addEventListener('click', setOutputPath);

document.getElementById('checkRunsForm').addEventListener('submit', async (event) => {
  event.preventDefault(); // Prevent the default form submission behavior
  const form = event.currentTarget;
  const formData = new FormData(form);
  const values = Object.fromEntries(formData.entries());
  const normalizedRegisterPath = values.uploadedRegisterFilePath.replace(/\\\\/g, "\\");
  const normalizedRunPath = values.uploadedRunFilePath.replace(/\\\\/g, "\\");
  const normalizedOutputPath = values.outputFilePath.replace(/\\\\/g, "\\");
  const saveLogs = values.saveLogs === "yes" ? true : false;
  //console.log(`Register: ${normalizedRegisterPath}, Run: ${normalizedRunPath}, Output Path: ${normalizedOutputPath}`);

  document.getElementById('status').innerText = "Validating inputs...";
  document.getElementById('uploadedRegisterFilePath').classList.remove("error");
  document.getElementById('uploadedRunFilePath').classList.remove("error");
  document.getElementById('outputPath').classList.remove("error");
  document.getElementById('error-message').innerHTML = ``;
  document.getElementById('error-message').style.display = "none";

  // Form Validation
    if (!normalizedRegisterPath || !isFilePath(normalizedRegisterPath) 
          || !normalizedRunPath || !isFilePath(normalizedRunPath) 
          || !normalizedOutputPath || !isFolderPath(normalizedOutputPath)) {
      document.getElementById('error-message').innerHTML = `<span style="color: red;">Please fill in all fields.</span>`;
      document.getElementById('error-message').style.display = "block";
      document.getElementById('status').innerHTML = '';

      if(!normalizedRegisterPath || !isFilePath(normalizedRegisterPath)){
        document.getElementById('uploadedRegisterFilePath').value = "Please select a file...";
        document.getElementById('uploadedRegisterFilePath').classList.add("error");
        document.getElementById('status').innerHTML += `<span style="color: red;">Please select a register file.</span>`;
      }
      if(!normalizedRunPath || !isFilePath(normalizedRunPath)){
        document.getElementById('uploadedRunFilePath').value = "Please select a file...";
        document.getElementById('uploadedRunFilePath').classList.add("error");
        document.getElementById('status').innerHTML += `<span style="color: red;">Please select a run file.</span>`;
      }
      if(!normalizedOutputPath || !isFolderPath(normalizedOutputPath)){
        document.getElementById('outputPath').value = "Please select a folder...";
        document.getElementById('outputPath').classList.add("error");
        document.getElementById('status').innerHTML += `<span style="color: red;">Please select an output folder.</span>`;
      }
      return;
    }

    document.getElementById('status').innerText = "Processing...";

     try{
      generatedFile =
        await window.electronAPI.processFile({
          registerFilePath: normalizedRegisterPath,
          runFilePath: normalizedRunPath,
          outputPath: normalizedOutputPath,
          saveLogs: saveLogs
        });
    } catch (error) {
        console.error("Error processing file:", error);
        document.getElementById('error-message').innerHTML = `<span style="color: red;">An error occurred while processing the file.</span>`;
        document.getElementById('error-message').style.display = "block";
        document.getElementById('status').innerHTML = `<span style="color: red;">${error.message}</span>`;
        return;
    }

    document.getElementById('status').innerText = "Finished.";
    document.getElementById('uploadedRegisterFilePath').classList.remove("error");
    document.getElementById('uploadedRunFilePath').classList.remove("error");
    document.getElementById('outputPath').classList.remove("error");
    document.getElementById('error-message').innerHTML = ``;
    document.getElementById('error-message').style.display = "none";

});
