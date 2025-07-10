const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

async function mergeSheets(masterFilePath, outputFilePath) {
  const masterWB = new ExcelJS.Workbook();

  // Load master and child files
  await masterWB.xlsx.readFile(masterFilePath);

  //   // Remove existing Sheet2 if present
  const existingSheet2 = masterWB.getWorksheet("Detailed Model");

  // Get child Sheet2
  //   const childSheet = childWB.getWorksheet("Sheet1") || childWB.worksheets[0];
  //   if (!childSheet) {
  //     console.error("❌ Sheet not found in child file.");
  //     return;
  //   }

  // Clone child sheet into master
  const newSheet = masterWB.addWorksheet("New Sheet");
  const columnsI = existingSheet2.getColumn('I');

  existingSheet2.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const newRow = newSheet.getRow(rowNumber);
    row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const newCell = newRow.getCell(colNumber);
    //   if (cell.type === Excel.ValueType.Formula) {
    //     cell.value = cell.result;
    //   }

      newCell.value = typeof cell.value === "object" && cell.formula ? cell.result : cell.value;
      newCell._column.hidden = cell._column.hidden
      newCell._column.width = cell._column.width
      newCell._row.height = cell._row.height
      newCell._row.hidden = cell._row.hidden

      // Copy style
      newCell.style = { ...cell.style };
    });
    newRow.commit();
  });

  const detailedSheet = masterWB.getWorksheet("Detailed Model");
  console.log('ColumnI', columnsI)

  const data = [];

  const rowData = [];
  detailedSheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      rowData.push({
        address: cell.address,
        value: cell.value,
        type: cell.type,
        formula: cell.formula || null,
      });
    });
    data.push(rowData);
  });

  fs.writeFile(
    path.join(__dirname, "data.json"),
    JSON.stringify(data),
    (err) => {
      console.log("error", err);
    }
  );
  // Save final file
  await masterWB.xlsx.writeFile(outputFilePath);
  console.log(`✅ Final file created: ${outputFilePath}`);
}

// Usage
const masterFilePath = path.join(__dirname, "output.xlsx/result.xlsx");
const childFilePath = path.join(__dirname, "child.xlsx");
const outputFilePath = path.join(__dirname, "result_final.xlsx");
const newFilePath = path.join(__dirname, "updated_master.xlsx");

mergeSheets(masterFilePath, outputFilePath);
