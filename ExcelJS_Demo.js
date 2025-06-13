//Section 13.67: Introduction to excelJS node module and setting up JS Project
//

//ExcelJS is used to read, manipulate and write spreadsheet data and styles to XLSX and JSON.

//First import the exceljs module/library:
const ExcelJS = require("exceljs");

async function outputCellValues() {
  //Excel spreadsheets are often called "workbooks", which is a term used with ExcelJS.
  //Create an object of the ExcelJS Class, then we can now use this Class and its methods etc.
  //One method is "Workbook()", which allows for access to Excel workbooks:
  const workbook = new ExcelJS.Workbook();

  //Now link the path of the Excel file you want to work with.
  //You can specify ".xlsx" or ".json" before readFile().
  //Need to use "await" as JS will attempt to execute the subsequent code before actually reading the file:
  await workbook.xlsx.readFile("C:/Users/Roscoe/Downloads/excel_download_test.xlsx");

  //Workbooks can have multiple "sheets" (the tabs at the bottom in Excel).
  //You need to specify the worksheet within the workbook first, using "getWorksheet()":
  const worksheet1 = workbook.getWorksheet(`Sheet1`);

  //A loop to output the value in each cell.
  //1. Iterates through each row, using a rowNumber argument.
  //2. Iterates through each cell within that row, using a colNumber argument.
  worksheet1.eachRow((row, rowNumber) => {
    row.eachCell((cell, columnNumber) => {
      console.log(`Row ${rowNumber}, Col ${columnNumber}, Value: ${cell.value}`);
    });
  });
}

outputCellValues();
