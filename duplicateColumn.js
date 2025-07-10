const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

async function mergeSheets(masterFilePath, outputFilePath) {
  const masterWB = new ExcelJS.Workbook();

  // Load master and child files
  await masterWB.xlsx.readFile(masterFilePath);

  // Remove existing Sheet2 if present
  const worksheet = masterWB.getWorksheet("Detailed Model");

  const sourceColumn = worksheet.getColumn("I");
  const targetColumn = worksheet.getColumn("AB");

  function replaceColumnReferencesInObject(obj, oldCol, newCol) {
    let stringified = JSON.stringify(obj);

    // Match cell refs: I19, $I$19, I$94, I19:I19, etc. — not IF, INDIRECT
    const refRegex = new RegExp(`\\$?${oldCol}\\$?\\d+`, "gi");

    stringified = stringified.replace(refRegex, (match) => {
      return match.replace(new RegExp(`\\$?${oldCol}`, "i"), (colMatch) => {
        return colMatch.replace(oldCol, newCol);
      });
    });

    return JSON.parse(stringified);
  }

  sourceColumn.eachCell({ includeEmpty: true }, (cell, rowNumber) => {
    const targetCell = worksheet.getCell(`AB${rowNumber}`);
    targetCell.value = replaceColumnReferencesInObject(cell.value, 'I', 'AB')
    targetCell.value = replaceColumnReferencesInObject(targetCell.value, 'J', 'AC')

    // Optional: Copy style
    targetCell.style = cell.style;
  });

  //   existingSheet2.eachRow({ includeEmpty: true }, (row, rowNumber) => {
  //     const newRow = newSheet.getRow(rowNumber);
  //     row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
  //       const newCell = newRow.getCell(colNumber);
  //       newCell.value = cell.value;

  //       // Copy style
  //       newCell.style = { ...cell.style };
  //     });
  //     newRow.commit();
  //   });

  //   const detailedSheet = masterWB.getWorksheet("Detailed Model");

  //   const data = [];

  //   const rowData = [];
  //   detailedSheet.eachRow((row, rowNumber) => {
  //     row.eachCell((cell, colNumber) => {
  //       rowData.push({
  //         address: cell.address,
  //         value: cell.value,
  //         formula: cell.formula || null,
  //       });
  //     });
  //     data.push(rowData);
  //   });

  //   fs.writeFile(path.join(__dirname, "data.json"), JSON.stringify(data), (err) => {
  //     console.log('error', err)
  //   });
  // Save final file
  await masterWB.xlsx.writeFile(outputFilePath);
  console.log(`✅ Final file created: ${outputFilePath}`);
}

// Usage
const masterFilePath = path.join(__dirname, "result_final.xlsx");
// const childFilePath = path.join(__dirname, "child.xlsx");
const outputFilePath = path.join(__dirname, "updated.xlsx");
// const newFilePath = path.join(__dirname, "updated_master.xlsx");

mergeSheets(masterFilePath, outputFilePath);
