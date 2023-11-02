function createNewExcelFile() {
  var Excel = require("exceljs");
  var workbook = new Excel.Workbook();

  workbook.creator = "Erasmo";
  workbook.lastModifiedBy = "";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.lastPrinted = new Date();

  var sheet = workbook.addWorksheet("My Sheet");
  sheet.columns = [
    { header: "Id", key: "id", width: 10 },
    { header: "Name", key: "name", width: 32 },
    { header: "D.O.B.", key: "DOB", width: 10, outlineLevel: 1 },
  ];

  sheet.addRow({ id: 1, name: "John Doe", dob: new Date(1970, 1, 1) });
  sheet.addRow({ id: 2, name: "Jane Doe", dob: new Date(1965, 1, 7) });
  sheet.addRow({ id: 3, name: "Jane Doe", dob: new Date(1965, 1, 7) });

  workbook.xlsx.writeFile("c:\\_temp\\file1.xlsx").then(function () {
    console.log("Archivo guardado!");
  });
}
