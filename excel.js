const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

async function mergeSheets(masterFilePath, childFilePath, outputFilePath, newFile) {
  const masterWB = new ExcelJS.Workbook();
  const childWB = new ExcelJS.Workbook();
  const extraFile = new ExcelJS.Workbook();

  // Load master and child files
  await masterWB.xlsx.readFile(masterFilePath);
  await childWB.xlsx.readFile(childFilePath);
  await extraFile.xlsx.readFile(newFile)

  // Remove existing Sheet2 if present
  const existingSheet2 = masterWB.getWorksheet("Asset Codes");
  
  if (existingSheet2) {
    masterWB.removeWorksheet(existingSheet2.id);
  }

  // Get child Sheet2
  const childSheet = childWB.getWorksheet("Sheet1") || childWB.worksheets[0];
  if (!childSheet) {
    console.error("❌ Sheet not found in child file.");
    return;
  }

  // Clone child sheet into master
  const newSheet = masterWB.addWorksheet("Asset Codes");

  childSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const newRow = newSheet.getRow(rowNumber);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const newCell = newRow.getCell(colNumber);
      newCell.value = cell.value;

      // Copy style
      newCell.style = { ...cell.style };
    });
    newRow.commit();
  });

  const detailedSheet = masterWB.getWorksheet("Detailed Model");

  const data = [];

  const rowData = [];
  detailedSheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      rowData.push({
        address: cell.address,
        value: cell.value,
        formula: cell.formula || null,
      });
    });
    data.push(rowData);
  });

  fs.writeFile(path.join(__dirname, "data.json"), JSON.stringify(data), (err) => {
    console.log('error', err)
  });
  // Save final file
  await masterWB.xlsx.writeFile(outputFilePath);
  console.log(`✅ Final file created: ${outputFilePath}`);
}

// Usage
const masterFilePath = path.join(__dirname, "example.xlsx");
const childFilePath = path.join(__dirname, "child.xlsx");
const outputFilePath = path.join(__dirname, "result.xlsx");
const newFilePath = path.join(__dirname, "updated_master.xlsx");


mergeSheets(masterFilePath, childFilePath, outputFilePath, newFilePath);
