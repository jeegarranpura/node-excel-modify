const ExcelJS = require("exceljs");
const path = require("path");
const fs = require("fs");

async function mergeSheets(masterFilePath, outputFilePath, childFilePath) {
  // const body = [
  //   {
  //     cellId: "I131",
  //     value: "'Sheet1'!$H$1:$BZ$1",
  //   },
  //   {
  //     cellId: J131,
  //     value: "'Sheet1'!$B$2:$B$500",
  //   },
  //   {
  //     cellId: K131,
  //     value: "'Sheet1'!$B$2:$B$500",
  //   },
  //   {
  //     cellId: I132,
  //     value: "'Sheet2'!$H$1:$BZ$1",
  //   },
  //   {
  //     cellId: J132,
  //     value: "'Sheet2'!$B$2:$B$500",
  //   },
  //   {
  //     cellId: K132,
  //     value: "'Sheet2'!$B$2:$B$500",
  //   },
  // ];
  const sheetBody = [
    {
      sheet_index: 0,
      sheet_name: "Sheet1",
      ranges: [
        {
          cellId: "I131",
          value: "$H$1:$BZ$1",
          cell: "$A$131",
        },
        {
          cellId: "J131",
          value: "$B$2:$B$500",
          cell: "$B$131",
        },
        {
          cellId: "K131",
          value: "$H$2:$BZ$500",
          cell: "$C$131",
        },
      ],
    },
    {
      sheet_index: 1,
      sheet_name: "Sheet2",
      ranges: [
        {
          cellId: "I132",
          value: "$H$1:$BZ$1",
          cell: "$A$132",
        },
        {
          cellId: "J132",
          value: "$B$2:$B$500",
          cell: "$B$132",
        },
        {
          cellId: "K132",
          value: "$H$2:$BZ$500",
          cell: "$C$132",
        },
      ],
    },
  ];
  const masterWB = new ExcelJS.Workbook();
  const childWB = new ExcelJS.Workbook();

  await masterWB.xlsx.readFile(masterFilePath);
  await childWB.xlsx.readFile(childFilePath);
  const worksheet = masterWB.getWorksheet("Detailed Model");

  sheetBody.filter((items) => {
    items.ranges.filter((subItems) => {
      // =CONCATENATE("'","Asset Codes","'","!","$A$1:$A$10")
      worksheet.getCell(subItems.cell).value = {
        formula: `=CONCATENATE("'","${items.sheet_name}","'","!","${subItems.value}")`,
      };
    });

    const childSheet = childWB.getWorksheet(items.sheet_name);
    if (!childSheet) {
      console.error("❌ Sheet not found in child file.");
      return;
    }

    const newSheet = masterWB.addWorksheet(items.sheet_name);

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
  });

  // Mapping of source to target columns
  const columnMap = {
    I: "M",
    J: "N",
    K: "O",
  };

  const formulaReplaceMap = {
    $A$131: "$A$132",
    $B$131: "$B$132",
    $C$131: "$C$132",
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

  function replaceSpecificFormulaValues(obj, replaceMap) {
    let stringified = JSON.stringify(obj);

    Object.entries(replaceMap).forEach(([oldVal, newVal]) => {
      const regex = new RegExp(oldVal.replace(/\$/g, "\\$"), "g");
      stringified = stringified.replace(regex, newVal);
    });

    return JSON.parse(stringified);
  }

  function replaceColumnLetters(value, columnMap) {
    let stringified = JSON.stringify(value);

    Object.entries(columnMap).forEach(([sourceCol, targetCol]) => {
      const regex = new RegExp(`(\\$?)${sourceCol}(\\$?\\d+)`, "gi");
      stringified = stringified.replace(regex, (match, dollarSign, rowPart) => {
        return `${dollarSign}${targetCol}${rowPart}`;
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
        newValue = replaceColumnLetters(newValue, columnMap);
        newValue = replaceSpecificFormulaValues(newValue, formulaReplaceMap);
      }

      targetCell.value = newValue;
      targetCell.style = cell.style;
    });
  });

  await masterWB.xlsx.writeFile(outputFilePath);
  console.log(`✅ Final file created: ${outputFilePath}`);
}

// Usage
const masterFilePath = path.join(__dirname, "new_example.xlsx");
const outputFilePath = path.join(__dirname, "updated_sample.xlsx");
const childfilePath = path.join(__dirname, "child_new.xlsx");

mergeSheets(masterFilePath, outputFilePath, childfilePath);
