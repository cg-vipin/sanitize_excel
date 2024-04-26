# Sanitize Excel

<img src="https://raw.githubusercontent.com/vipinvkumar/sanitizes_excel/main/public/excel_sanitize.png" alt="Sanitize Excel" style="display: block; margin-left: auto; margin-right: auto; width: 50%;">

Sanitize Excel is a web app that allows you to upload an Excel file and then download a new Excel file with sanitized data. It trims leading and trailing spaces from string values, converts numbers to their numerical form, and converts booleans to their boolean form. It also keeps dates as is.

### Installation

1. Clone this repository
2. Run `npm install`
3. Run `node app.js`
4. Go to `http://localhost:3000` in your web browser

### How it works

When you upload an Excel file, the app reads it using the `xlsx` library, sanitizes the data, and then creates a new Excel file using the `exceljs` library. The new file is downloaded to your computer.

### Limitations

Currently, only the first sheet of the Excel file is sanitized. The app does not currently support sanitizing multiple sheets or handling merged cells.

### Future improvements

In the future, the app could be improved to support sanitizing multiple sheets and handling merged cells. It could also be improved to provide more user feedback and error handling.


