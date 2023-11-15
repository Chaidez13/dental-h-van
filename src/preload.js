// See the Electron documentation for details on how to use preload scripts:
// https://www.electronjs.org/docs/latest/tutorial/process-model#preload-scripts

const { ipcRenderer } = require("electron");

document.addEventListener("DOMContentLoaded", () => {
  let myButton = document.getElementById("myButton");
  myButton.addEventListener("click", () => {
    const patientData = {
      name: document.getElementById("patientName").value,
      phone: document.getElementById("patientPhone").value,
      direction: document.getElementById("patientDirection").value,
    };

    //ipcRenderer.send("savePatient", patientData);
    ipcRenderer.send("openFile", patientData);
  });
});
