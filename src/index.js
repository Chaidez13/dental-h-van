const { app, BrowserWindow, ipcMain } = require("electron");
const { createExcelTemplate } = require("./template/excelTemplate.js");
const path = require("path");
var fs = require("fs");

try {
  require("electron-reloader")(module);
} catch (_) {}

// Handle creating/removing shortcuts on Windows when installing/uninstalling.
if (require("electron-squirrel-startup")) {
  app.quit();
}

const createWindow = () => {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: true,
      preload: path.join(__dirname, "preload.js"),
    },
  });

  // and load the index.html of the app.
  mainWindow.loadFile(path.join(__dirname, "index.html"));

  // Open the DevTools.
  mainWindow.webContents.openDevTools();
};

ipcMain.on("savePatient", (event, patientData) => {
  var Excel = require("exceljs");
  var workbook = new Excel.Workbook();

  workbook.creator = "Erasmo";
  workbook.lastModifiedBy = "";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.lastPrinted = new Date();

  var sheet = workbook.addWorksheet(patientData.name);
  sheet.columns = [
    { width: 18.5, key: "A" }, //A
    { width: 11, key: "B" }, //B
    { width: 11, key: "C" }, //C
    { width: 11, key: "D" }, //D
    { width: 11, key: "E" }, //E
    { width: 11, key: "F" }, //F
    { width: 11, key: "G" }, //G
    { width: 8, key: "H" }, //H
    { width: 12, key: "I" }, //I
    { width: 12, key: "J" }, //J
    { width: 12, key: "K" }, //K
    { width: 8, key: "L" }, //L
    { width: 12, key: "M" }, //M
    { width: 12, key: "N" }, //N
    { width: 12, key: "O" }, //O
  ];

  //Poner logo en la celda A1 a A6
  const escudo = workbook.addImage({
    filename: "./src/assets/escudo-negro.gif",
    extension: "gif",
  });
  sheet.addImage(escudo, "A1:A5");

  //Poner logo en la celda F1 a G5
  const logo = workbook.addImage({
    filename: "./src/assets/logo.png",
    extension: "png",
  });
  sheet.addImage(logo, "G2:G4");

  //Cargar la plantilla
  sheet = createExcelTemplate(sheet, patientData);

  workbook.xlsx
    .writeFile("c:\\_temp\\" + patientData.name + ".xlsx")
    .then(function () {
      console.log("Archivo guardado!");
    });
  //Crear Archivo Asincrono
  //   fs.writeFile("c:\\_temp\\file1.txt", txtValue.toString(), function (err) {
  //     if (err) throw err;
  //     console.log("Archivo guardado!");
  //   });

  //Crear Archivo Sincrono
  //   fs.writeFileSync("c:\\_temp\\file1.txt", txtValue);

  //Agregar texto a un archivo
  //   fs.appendFileSync(
  //     "c:\\_temp\\file1.txt",
  //     txtValue.toString(),
  //     function (err) {
  //       if (err) throw err;
  //       console.log("Archivo guardado!");
  //     }
  //   );
});

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on("ready", createWindow);

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on("window-all-closed", () => {
  if (process.platform !== "darwin") {
    app.quit();
  }
});

app.on("activate", () => {
  // On OS X it's common to re-create a window in the app when the
  // dock icon is clicked and there are no other windows open.
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow();
  }
});

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and import them here.
