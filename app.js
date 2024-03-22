/**
 * This script handles the functionality for generating an Excel file from configuration data.
 * It utilizes the XLSX library for creating Excel files.
 *    - 0.17.4/xlsx.full.min.js
 *    - 2.0.5/FileSaver.min.js
 */

document.addEventListener("DOMContentLoaded", function () {
  var fileUploadInput = document.getElementById("file-upload");
  var cfgTextarea = document.getElementById("cfg-textarea");
  var generateExcelBtn = document.getElementById("generate-excel-btn");
  var resetBtn = document.getElementById("reset-btn");

  resetBtn.addEventListener("click", function () {
    fileUploadInput.value = "";
    fileUploadInput.disabled = false;
    cfgTextarea.value = "";
    cfgTextarea.disabled = false;
    generateExcelBtn.disabled = true;
    resetBtn.disabled = true;
  });


  fileUploadInput.addEventListener("change", function (event) {
    var file = event.target.files[0];
    if (file) {
      cfgTextarea.disabled = true;
      generateExcelBtn.disabled = false;
    } else {
      cfgTextarea.disabled = false;
      generateExcelBtn.disabled = true;
    }
    resetBtn.disabled = false;
  });

  cfgTextarea.addEventListener("input", function () {
    if (cfgTextarea.value.trim() !== "") {
      fileUploadInput.disabled = true;
      generateExcelBtn.disabled = false;
    } else {
      fileUploadInput.disabled = false;
      generateExcelBtn.disabled = true;
    }
    resetBtn.disabled = false;
  });

  generateExcelBtn.addEventListener("click", function () {
    fileName = generateFileName();
    if (fileUploadInput.files.length > 0) {
      var file = fileUploadInput.files[0];
      var uploadFileName = file.name;
      // use regex to remove suffix .cfg from the file name
      uploadFileName = uploadFileName.replace(/\.cfg$/, "");
      // use regex to remove the export prefix from the file name
      fileName = fileName.replace(/^export_/, "");
      fileName = `${uploadFileName}_${fileName}`;
      var reader = new FileReader();
      reader.onload = function (event) {
        var cfgData = event.target.result;
        var workbook = createWorkbookFromCfg(cfgData);
        saveWorkbookAsExcel(workbook, fileName);
      };
      reader.readAsText(file);
    } else {
      var cfgData = cfgTextarea.value;
      var workbook = createWorkbookFromCfg(cfgData);
      saveWorkbookAsExcel(workbook, fileName);
    }
  });

  function createWorkbookFromCfg(cfgData) {
    var lines = cfgData.split("\n");
    var workbook = XLSX.utils.book_new();

    var currentSheet = null;
    var sheetData = [];

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i].trim();
      if (line.startsWith("[")) {
        if (currentSheet !== null && sheetData.length > 0) {
          addSheetToWorkbook(workbook, currentSheet, sheetData);
          sheetData = [];
        }
        currentSheet = line.substring(1, line.length - 1);
      } else {
        var keyValue = line.split("=");
        if (keyValue.length === 2) {
          sheetData.push(keyValue);
        }
      }
    }

    if (currentSheet !== null && sheetData.length > 0) {
      addSheetToWorkbook(workbook, currentSheet, sheetData);
    }

    return workbook;
  }

  function addSheetToWorkbook(workbook, sheetName, data) {
    var ws = XLSX.utils.aoa_to_sheet([...data]);
    XLSX.utils.book_append_sheet(workbook, ws, sheetName);
  }

  function saveWorkbookAsExcel(workbook, fileName) {
    var wbout = XLSX.write(workbook, { type: "binary", bookType: "xlsx" });
    var blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    saveAs(blob, fileName);
  }

  // string to array buffer
  function s2ab(s) {
    var buf = new ArrayBuffer(s.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }

  function generateFileName() {
    var date = new Date();
    var year = date.getFullYear();
    var month = ("0" + (date.getMonth() + 1)).slice(-2);
    var day = ("0" + date.getDate()).slice(-2);
    var epoch = date.getTime();
    return `export_${year}${month}${day}_${epoch}.xlsx`;
  }
});
