const createExcelTemplate = (sheet, patientData) => {
  //Historia Clininca [1ra pagina]
  sheet.mergeCells("A5:G5");
  sheet.getCell("A5").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("A5").style.font = {
    name: "Calibri",
    size: 18,
    bold: true,
  };
  sheet.getCell("A5").value = "HISTORIA CLINICA";
  sheet.getCell("A7").style.font = { bold: true };
  sheet.getCell("A7").value = "C.D: ERANDIDI RAMIREZ RAMIREZ";
  sheet.getCell("D7").style.font = { bold: true };
  sheet.getCell("D7").value = "CEPROF. 6740782";
  sheet.getCell("A8").value =
    "Dirección: Javier Mina 1950 e. Bravo y Rosales Col. Olivos";
  sheet.getCell("A9").value = "En La Paz, Baja California Sur";
  sheet.getCell("A10").value = "Fecha: "; //'=CONCAT("Fecha: ", TEXTO(HOY(), "dd/mm/aaaa"))';
  sheet.mergeCells("A12:G12");
  sheet.getCell("A12").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("A12").style.font = {
    name: "Calibri",
    size: 18,
    bold: true,
  };
  sheet.getCell("A12").value = "FICHA DE INDENTIFICACIÓN";
  sheet.getCell("A14").style.font = { bold: true };
  sheet.getCell("A14").value = "Nombre:";
  sheet.getCell("B14").value = patientData.name;
  sheet.getCell("A15").style.font = { bold: true };
  sheet.getCell("A15").value = "Fecha de Nacimiento:";
  sheet.getCell("B15").value = patientData.birthDate;
  sheet.getCell("D15").style.font = { bold: true };
  sheet.getCell("D15").value = "Edad:";
  sheet.getCell("E15").value = patientData.age;
  sheet.getCell("A16").style.font = { bold: true };
  sheet.getCell("A16").value = "Dirección:";
  sheet.getCell("B16").value = patientData.direction;
  sheet.getCell("A17").style.font = { bold: true };
  sheet.getCell("A17").value = "Télefono:";
  sheet.getCell("B17").value = patientData.phone;
  sheet.getCell("D17").style.font = { bold: true };
  sheet.getCell("D17").value = "Celular:";
  sheet.getCell("E17").value = patientData.cell;
  sheet.getCell("A18").style.font = { bold: true };
  sheet.getCell("A18").value = "Ocupación:";
  sheet.getCell("B18").value = patientData.ocupacion;
  sheet.getCell("A19").style.font = { bold: true };
  sheet.getCell("A19").value = "Religión:";
  sheet.getCell("B19").value = patientData.religion;
  sheet.getCell("A20").style.font = { bold: true };
  sheet.mergeCells("A20:B20");
  sheet.getCell("A20").value = "Motivo de la consulta:";
  sheet.getCell("C20").value = patientData.motivoConsulta;
  sheet.getCell("A21").style.font = { bold: true };
  sheet.getCell("A21").value = "Alergias:";
  sheet.getCell("B21").value = patientData.alergias;
  //Datos antececentes heredo-familiares [1ra pagina]
  sheet.mergeCells("A25:G25");
  sheet.getCell("A25").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("A25").style.font = {
    name: "Calibri",
    size: 18,
    bold: true,
  };
  sheet.getCell("A25").value = "ANTECEDENTES HEREDO-FAMILIARES";
  sheet.mergeCells("A27:C27");
  sheet.getCell("A27").value = "Tiene familiares con enfermedades como:";
  sheet.getCell("A28").value = "D.M.";
  sheet.getCell("C28").value = "Asma:";
  sheet.getCell("E28").value = "Cancer:";
  sheet.getCell("A29").value = "Problemas Cardiacos:";
  sheet.getCell("C29").value = "Epilepsia:";
  sheet.getCell("E29").value = "SIDA:";
  sheet.getCell("A30").value = "Prognatismo:";
  sheet.getCell("C30").value = "Otros:";

  //Seccion Dx y Tx (Dientes) [2da pagina]
  sheet.mergeCells("H2:O2");
  sheet.getCell("H2").style.font = {
    name: "Calibri",
    size: 16,
    bold: true,
  };
  sheet.getCell("H2").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H2").value = "Dx y Tx";
  //Superior Derecho
  sheet.mergeCells("H3:K3");
  sheet.getCell("H3").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H3").value = "Superior Derecho";
  sheet.getCell("H4").style.font = { bold: true };
  sheet.getCell("H4").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H4").value = "51/11";
  sheet.mergeCells("I4:K4");
  sheet.getCell("H5").style.font = { bold: true };
  sheet.getCell("H5").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H5").value = "52/12";
  sheet.mergeCells("I5:K5");
  sheet.getCell("H6").style.font = { bold: true };
  sheet.getCell("H6").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H6").value = "53/13";
  sheet.mergeCells("I6:K6");
  sheet.getCell("H7").style.font = { bold: true };
  sheet.getCell("H7").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H7").value = "54/14";
  sheet.mergeCells("I7:K7");
  sheet.getCell("H8").style.font = { bold: true };
  sheet.getCell("H8").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H8").value = "55/15";
  sheet.mergeCells("I8:K8");
  sheet.getCell("H9").style.font = { bold: true };
  sheet.getCell("H9").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H9").value = "16";
  sheet.mergeCells("I9:K9");
  sheet.getCell("H10").style.font = { bold: true };
  sheet.getCell("H10").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H10").value = "17";
  sheet.mergeCells("I10:K10");
  //Superior Izquierdo
  sheet.mergeCells("L3:O3");
  sheet.getCell("L3").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L3").value = "Superior Izquierdo";
  sheet.getCell("L4").style.font = { bold: true };
  sheet.getCell("L4").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L4").value = "61/21";
  sheet.mergeCells("M4:O4");
  sheet.getCell("L5").style.font = { bold: true };
  sheet.getCell("L5").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L5").value = "62/22";
  sheet.mergeCells("M5:O5");
  sheet.getCell("L6").style.font = { bold: true };
  sheet.getCell("L6").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L6").value = "63/23";
  sheet.mergeCells("M6:O6");
  sheet.getCell("L7").style.font = { bold: true };
  sheet.getCell("L7").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L7").value = "64/24";
  sheet.mergeCells("M7:O7");
  sheet.getCell("L8").style.font = { bold: true };
  sheet.getCell("L8").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L8").value = "65/25";
  sheet.mergeCells("M8:O8");
  sheet.getCell("L9").style.font = { bold: true };
  sheet.getCell("L9").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L9").value = "26";
  sheet.mergeCells("M9:O9");
  sheet.getCell("L10").style.font = { bold: true };
  sheet.getCell("L10").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L10").value = "27";
  sheet.mergeCells("M10:O10");
  //Inferior Derecho
  sheet.mergeCells("H12:K12");
  sheet.getCell("H12").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H12").value = "Inferior Derecho";
  sheet.getCell("H13").style.font = { bold: true };
  sheet.getCell("H13").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H13").value = "81/41";
  sheet.mergeCells("I13:K13");
  sheet.getCell("H14").style.font = { bold: true };
  sheet.getCell("H14").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H14").value = "82/42";
  sheet.mergeCells("I14:K14");
  sheet.getCell("H15").style.font = { bold: true };
  sheet.getCell("H15").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H15").value = "83/43";
  sheet.mergeCells("I15:K15");
  sheet.getCell("H16").style.font = { bold: true };
  sheet.getCell("H16").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H16").value = "84/44";
  sheet.mergeCells("I16:K16");
  sheet.getCell("H17").style.font = { bold: true };
  sheet.getCell("H17").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H17").value = "85/45";
  sheet.mergeCells("I17:K17");
  sheet.getCell("H18").style.font = { bold: true };
  sheet.getCell("H18").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H18").value = "46";
  sheet.mergeCells("I18:K18");
  sheet.getCell("H19").style.font = { bold: true };
  sheet.getCell("H19").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("H19").value = "47";
  sheet.mergeCells("I19:K19");
  //Inferior Izquierdo
  sheet.mergeCells("L12:O12");
  sheet.getCell("L12").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L12").value = "Inferior Izquierdo";
  sheet.getCell("L13").style.font = { bold: true };
  sheet.getCell("L13").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L13").value = "71/31";
  sheet.mergeCells("M13:O13");
  sheet.getCell("L14").style.font = { bold: true };
  sheet.getCell("L14").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L14").value = "72/32";
  sheet.mergeCells("M14:O14");
  sheet.getCell("L15").style.font = { bold: true };
  sheet.getCell("L15").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L15").value = "73/33";
  sheet.mergeCells("M15:O15");
  sheet.getCell("L16").style.font = { bold: true };
  sheet.getCell("L16").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L16").value = "74/34";
  sheet.mergeCells("M16:O16");
  sheet.getCell("L17").style.font = { bold: true };
  sheet.getCell("L17").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L17").value = "75/35";
  sheet.mergeCells("M17:O17");
  sheet.getCell("L18").style.font = { bold: true };
  sheet.getCell("L18").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L18").value = "36";
  sheet.mergeCells("M18:O18");
  sheet.getCell("L19").style.font = { bold: true };
  sheet.getCell("L19").alignment = { horizontal: "center", vertical: "center" };
  sheet.getCell("L19").value = "37";
  sheet.mergeCells("M19:O19");

  return sheet;
};

module.exports = {
  createExcelTemplate,
};
