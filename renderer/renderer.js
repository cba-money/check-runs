let generatedFile = null;

async function uploadRegisterFile(){
  const file = await window.electronAPI.uploadFile();
  if(file === undefined || file === null || file.length <= 0){
      document.getElementById('uploadedRegisterFilePath').value = "Please select a folder...";
      return;
  }
  console.log(`File uploaded:${file}`);
  document.getElementById('uploadedRegisterFilePath').value = file;
}

async function uploadRunFile(){
  const file = await window.electronAPI.uploadFile();
  if(file === undefined || file === null || file.length <= 0){
      document.getElementById('uploadedRunFilePath').value = "Please select a folder...";
      return;
  }
  console.log(`File uploaded:${file}`);
  document.getElementById('uploadedRunFilePath').value = file;
}

document.getElementById('registerFileUpload').addEventListener('click', uploadRegisterFile);
document.getElementById('runFileUpload').addEventListener('click', uploadRunFile);

document
  .getElementById('submitBtn')
  .addEventListener('click', async () => {

    let absolutePath = null;

    const registerFilePath = document.getElementById('uploadedRegisterFilePath').value;
    const runFilePath = document.getElementById('uploadedRunFilePath').value;
    /*
    if (!file) {
      alert("Please select a file.");
      return;
    }
    */

    //const dateRanges = document.getElementById('dateRanges').value;

    document.getElementById('status').innerText = "Processing...";

    generatedFile =
      await window.electronAPI.processFile({
        registerFilePath: registerFilePath,
        runFilePath: runFilePath
      });

    document.getElementById('status').innerText = "Finished.";

    //document.getElementById('downloadBtn').style.display = "inline-block";
  });

/*
document
  .getElementById('downloadBtn')
  .addEventListener('click', async () => {
    await window.electronAPI.saveOutput(
      generatedFile
    );
  });
*/