// readExcel.js (Node.js server-side)
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

const rowData = [];
async function readExcel(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.worksheets[1]; // assuming first sheet

  const data = [];

  sheet.eachRow((row, rowNumber) => {
    
    row.eachCell((cell, colNumber) => {
         rowData.push({
        address: cell.address,
        value: JSON.stringify(cell.value),
        formula: cell.formula || null,
        result: cell.result,
        // cell: cell
      });
    });
    data.push(rowData);
  });
  fs.writeFile(path.join(__dirname, "data1.json"), JSON.stringify(data), (err) => {
      console.log('error', err)
    });
  return data;
}

// Example usage
(async () => {
  const result = await readExcel(path.join(__dirname, '/output.xlsx/result.xlsx'));

  console.log(result);
})();



// async function createExcelFromRowData(rowData, outputPath) {
//   const workbook = new ExcelJS.Workbook();
//   const sheet = workbook.addWorksheet("Sheet1");

//   rowData.forEach((row, rowIndex) => {
//     row.forEach((cell, colIndex) => {
//       const excelCell = sheet.getCell(rowIndex + 1, colIndex + 1); // rows and cols are 1-based
//       if (cell.formula) {
//         excelCell.value = {
//           formula: cell.formula,
//           result: cell.value?.result ?? null,
//         };
//       } else {
//         excelCell.value = cell.value;
//       }
//     });
//   });

//   await workbook.xlsx.writeFile(outputPath);
//   console.log(`Excel file saved to ${outputPath}`);
// }

// // Usage
// setTimeout(() => {
//     createExcelFromRowData(rowData, path.join(__dirname, 'output.xlsx'));
// }, 10000)
