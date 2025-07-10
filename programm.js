const { exec } = require("child_process");
const path = require('path');
const inputPath = "./input.xlsx";

const inputFilePath = path.join(__dirname, "result.xlsx");
const outputFilePath = path.join(__dirname, "output.xlsx");
const outputPath = "./output";

exec(`soffice --headless --convert-to xlsx --calc --outdir ${outputFilePath} ${inputFilePath}`, (error, stdout, stderr) => {
  if (error) {
    console.error(`soffice  error: ${error.message}`);
    return;
  }
  console.log("File converted with recalculation.");
});
