const model = require("../models/excelModel");

// Function to sanitize an Excel file
function sanitizeExcelFile(inputPath, outputPath) {
  model.sanitizeExcel(inputPath, outputPath);
}

module.exports = {
  sanitizeExcelFile,
};
