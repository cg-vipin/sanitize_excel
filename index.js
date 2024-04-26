const XLSX = require("xlsx");

// Load the Excel file
const workbook = XLSX.readFile("input.xlsx", { cellStyles: true });

// Iterate through each sheet
workbook.SheetNames.forEach((sheetName) => {
  const worksheet = workbook.Sheets[sheetName];

  // Function to sanitize data
  function sanitizeData(cell) {
    // Check if cell is empty
    if (!cell || cell.t === undefined || cell.v === undefined) {
      return null;
    }

    // Sanitize based on cell type
    switch (cell.t) {
      case "s": // String type
        return cell.v.trim(); // Trim leading and trailing spaces
      case "n": // Numeric type
        return parseFloat(cell.v); // Convert to number
      case "b": // Boolean type
        return cell.v ? true : false; // Convert to boolean
      default:
        return cell.v;
    }
  }

  // Convert sheet to JSON to get the range of cells
  const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  // Iterate through each row and cell
  jsonData.forEach((row, rowIndex) => {
    row.forEach((cell, colIndex) => {
      // Create cell address
      const cellAddress = XLSX.utils.encode_cell({ r: rowIndex, c: colIndex });
      // Get cell object from worksheet
      const xlsCell = worksheet[cellAddress];
      // Sanitize cell data
      const sanitizedValue = sanitizeData(xlsCell);
      // Update cell value with sanitized data
      if (sanitizedValue !== null) {
        xlsCell.v = sanitizedValue;
      }
    });
  });
});

// Write the new workbook to a file in XLSX format
const outputFilePath = "output.xlsx";
XLSX.writeFile(workbook, outputFilePath, { bookType: "xlsx" });

console.log(`Sanitized Excel file created: ${outputFilePath}`);
