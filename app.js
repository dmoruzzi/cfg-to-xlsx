/**
 * This script handles the functionality for generating an Excel file from configuration data.
 * It utilizes the XLSX library for creating Excel files.
 *    - 0.17.4/xlsx.full.min.js
 *    - 2.0.5/FileSaver.min.js
 */

document.addEventListener("DOMContentLoaded", function () {
  const fileUploadInput = document.getElementById("file-upload");
  const cfgTextarea = document.getElementById("cfg-textarea");
  const generateExcelBtn = document.getElementById("generate-excel-btn");
  const resetBtn = document.getElementById("reset-btn");

  resetBtn.addEventListener("click", function () {
    fileUploadInput.value = "";
    fileUploadInput.disabled = false;
    cfgTextarea.value = "";
    cfgTextarea.disabled = false;
    generateExcelBtn.disabled = true;
    resetBtn.disabled = true;
  });

  fileUploadInput.addEventListener("change", function (event) {
    let file = event.target.files[0];
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
    let fileName = generateFileName();
    if (fileUploadInput.files.length > 0) {
      let file = fileUploadInput.files[0];
      let uploadFileName = file.name;
      uploadFileName = uploadFileName.replace(/\.cfg$/, "");
      fileName = fileName.replace(/^export_/, "");
      fileName = `${uploadFileName}_${fileName}`;
      let reader = new FileReader();
      reader.onload = function (event) {
        let cfgData = event.target.result;
        let workbook = createWorkbookFromCfg(cfgData);
        saveWorkbookAsExcel(workbook, fileName);
      };
      reader.readAsText(file);
    } else {
      let cfgData = cfgTextarea.value;
      let workbook = createWorkbookFromCfg(cfgData);
      saveWorkbookAsExcel(workbook, fileName);
    }
  });

  function createWorkbookFromCfg(cfgData) {
    let lines = cfgData.split("\n");
    let workbook = XLSX.utils.book_new();

    let currentSheet = null;
    let sheetData = [];

    for (let i = 0; i < lines.length; i++) {
      let line = lines[i].trim();
      if (line.startsWith("[")) {
        if (currentSheet !== null && sheetData.length > 0) {
          addSheetToWorkbook(workbook, currentSheet, sheetData);
          sheetData = [];
        }
        currentSheet = truncateSheetName(line.substring(1, line.length - 1));
      } else {
        let keyValue = line.split("=");
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
    let ws = XLSX.utils.aoa_to_sheet([...data]);
    XLSX.utils.book_append_sheet(workbook, ws, sheetName);
  }

  function saveWorkbookAsExcel(workbook, fileName) {
    let wbout = XLSX.write(workbook, { type: "binary", bookType: "xlsx" });
    let blob = new Blob([s2ab(wbout)], { type: "application/octet-stream" });
    saveAs(blob, fileName);
  }

  function truncateSheetName(sheetName) {
    const maxSheetLength = 31; // Maximum length for sheet names
    if (sheetName.length > maxSheetLength) {
        const truncateSuffix = "...";
        const truncatedLength = maxSheetLength - truncateSuffix.length;
        return sheetName.substring(0, truncatedLength) + truncateSuffix;
    }
    return sheetName;
  }

  function s2ab(s) {
    let buf = new ArrayBuffer(s.length);
    let view = new Uint8Array(buf);
    for (let i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
    return buf;
  }

  function generateFileName() {
    let date = new Date();
    let year = date.getFullYear();
    let month = ("0" + (date.getMonth() + 1)).slice(-2);
    let day = ("0" + date.getDate()).slice(-2);
    let epoch = date.getTime();
    return `export_${year}${month}${day}_${epoch}.xlsx`;
  }
});
