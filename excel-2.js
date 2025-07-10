const XLSX = require("xlsx");
const fs = require("fs");
const ExcelJS = require("exceljs");
const path = require("path");

// 1. Read master and child workbooks
const masterWB = XLSX.readFile("example.xlsx");
const childWB = XLSX.readFile("child.xlsx");

// 2. Extract Sheet2 from child (assuming name is 'Sheet2')
const sheet2FromChild = childWB.Sheets["Sheet1"];

// 3. Replace or add 'Sheet2' in master
masterWB.Sheets["Asset Codes"] = sheet2FromChild;

// Optional: Ensure 'Sheet2' exists in master sheet names
if (!masterWB.SheetNames.includes("Asset Codes")) {
  masterWB.SheetNames.push("Asset Codes");
}

// const sheetName = "Asset Codes"; // Replace with the actual sheet name
// // 4. Write the updated master workbook
// masterWB.SheetNames = masterWB.SheetNames.filter((name) => name !== sheetName);

XLSX.writeFile(masterWB, "updated_master.xlsx");

const addStyle = async (newFile, masterFile, outputFilePath) => {
  const masterWB = new ExcelJS.Workbook();
  const extraFile = new ExcelJS.Workbook();

  await masterWB.xlsx.readFile(masterFile);
  await extraFile.xlsx.readFile(newFile);

  const masterSheet = masterWB.getWorksheet("Detailed Model");
  const targetSheet = extraFile.getWorksheet("Detailed Model");

  targetSheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const masterRow = masterSheet.getRow(rowNumber);

    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const masterCell = masterRow.getCell(colNumber);

      // Only apply styles from masterSheet, do not change value
      cell.style = { ...masterCell.style };

      // Optional: preserve number format
      if (masterCell.numFmt) {
        cell.numFmt = masterCell.numFmt;
      }
    });

    row.commit();
  });

  await extraFile.xlsx.writeFile(outputFilePath);
};


const masterFilePath = path.join(__dirname, "example.xlsx");
const newFilePath = path.join(__dirname, "updated_master.xlsx");
const outputFilePath = path.join(__dirname, "result_new.xlsx");

addStyle(newFilePath, masterFilePath, outputFilePath);
