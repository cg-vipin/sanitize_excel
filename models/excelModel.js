const XLSX = require("xlsx");

// Function to sanitize a cell value based on its type
function sanitizeCellValue(cell) {
  if (cell.t === "n") {
    // Numeric
    return cell.v;
  } else if (cell.t === "s") {
    // String
    return cell.v.trim(); // Trim leading and trailing spaces
  } else if (cell.t === "b") {
    // Boolean
    return cell.v;
  } else if (cell.t === "d") {
    // Date
    return cell.w; // Keep the date string as is
  } else {
    return ""; // For other types, return empty string
  }
}

// Function to sanitize the Excel file
function sanitizeExcel(inputPath, outputPath) {
  try {
    // Load the Excel file
    const workbook = XLSX.readFile(inputPath);

    // Get the first sheet
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // Iterate through each cell in the sheet
    const range = XLSX.utils.decode_range(worksheet["!ref"]);
    for (let rowNum = range.s.r; rowNum <= range.e.r; rowNum++) {
      for (let colNum = range.s.c; colNum <= range.e.c; colNum++) {
        const cellAddress = XLSX.utils.encode_cell({ r: rowNum, c: colNum });
        const cell = worksheet[cellAddress];
        if (cell) {
          // Sanitize the cell value
          const sanitizedValue = sanitizeCellValue(cell);
          // Update the cell value in the worksheet
          worksheet[cellAddress] = { v: sanitizedValue };
        }
      }
    }

    // Create a new sheet with sanitized data
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.aoa_to_sheet(
      XLSX.utils.sheet_to_json(worksheet, { header: 1 })
    );
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Sanitized Sheet");

    // Write the sanitized workbook to a new Excel file
    XLSX.writeFile(newWorkbook, outputPath);

    console.log("Sanitized Excel file created successfully.");
  } catch (error) {
    console.error("Error occurred while sanitizing Excel file:", error);
  }
}

module.exports = {
  sanitizeExcel,
};
