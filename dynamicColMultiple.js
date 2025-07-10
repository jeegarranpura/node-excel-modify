const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

async function mergeSheets(masterFilePath, outputFilePath) {
  const masterWB = new ExcelJS.Workbook();
  await masterWB.xlsx.readFile(masterFilePath);
  const worksheet = masterWB.getWorksheet("Detailed Model");

  // Mapping of source to target columns
  const columnMap = {
    I: "M",
    J: "N",
    K: "O"
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
      if (newValue !== null && typeof newValue === 'object') {
        newValue = replaceColumnReferencesInObject(newValue, sourceCol, targetCol);
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

mergeSheets(masterFilePath, outputFilePath);
