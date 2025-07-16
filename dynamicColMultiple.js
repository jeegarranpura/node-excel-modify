const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

async function mergeSheets(masterFilePath, outputFilePath, childFilePath) {
  const body = [
    {
      cellId: I131,
      value: "'Sheet1'!$H$1:$BZ$1",
    },
    {
      cellId: J131,
      value: "'Sheet1'!$B$2:$B$500",
    },
    {
      cellId: K131,
      value: "'Sheet1'!$B$2:$B$500",
    },
    {
      cellId: I132,
      value: "'Sheet2'!$H$1:$BZ$1",
    },
    {
      cellId: J132,
      value: "'Sheet2'!$B$2:$B$500",
    },
    {
      cellId: K132,
      value: "'Sheet2'!$B$2:$B$500",
    },
  ];
  const masterWB = new ExcelJS.Workbook();
  const childWB = new ExcelJS.Workbook();

  await masterWB.xlsx.readFile(masterFilePath);
  await childWB.xlsx.readFile(childFilePath);
  const worksheet = masterWB.getWorksheet("Detailed Model");

  body.filter((items) => {
    worksheet.getCell(items.cell).value = items.value;
  });

  const childSheet = childWB.worksheets[0];
  if (!childSheet) {
    console.error("❌ Sheet not found in child file.");
    return;
  }

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

  // Mapping of source to target columns
  const columnMap = {
    I: "M",
    J: "N",
    K: "O",
  };

  // Utility to update formula/refs safely
  function replaceColumnReferencesInObject(obj, oldCol, newCol) {
    let stringified = JSON.stringify(obj);
    const refRegex = new RegExp(`\\$?${oldCol}\\$?\\d+`, "gi");

    stringified = stringified.replace(refRegex, (match) => {
      return match.replace(new RegExp(`\\$?${oldCol}`, "i"), (colMatch) => {
        return colMatch.replace(oldCol, newCol);
      });
    });

    return JSON.parse(stringified);
  }

  // Iterate over each source-target column pair
  Object.entries(columnMap).forEach(([sourceCol, targetCol]) => {
    const sourceColumn = worksheet.getColumn(sourceCol);

    sourceColumn.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
      const targetCell = worksheet.getCell(`${targetCol}${rowNumber}`);

      let newValue = cell.value;

      // Replace all references from sourceCol to targetCol
      if (newValue !== null && typeof newValue === "object") {
        newValue = replaceColumnReferencesInObject(
          newValue,
          sourceCol,
          targetCol
        );
      }

      targetCell.value = newValue;
      targetCell.style = cell.style;
    });
  });

  await masterWB.xlsx.writeFile(outputFilePath);
  console.log(`✅ Final file created: ${outputFilePath}`);
}

// Usage
const masterFilePath = path.join(__dirname, "sample_final.xlsx");
const outputFilePath = path.join(__dirname, "updated_sample.xlsx");
const childfilePath = path.join(__dirname, "child.xlsx");

mergeSheets(masterFilePath, outputFilePath, childfilePath);
