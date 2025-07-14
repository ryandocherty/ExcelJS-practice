//Section 13.68: Build Util functions to read and update excel file strategically

const ExcelJS = require("exceljs");

async function outputCellValues() {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("C:/Users/Roscoe/Desktop/Projects/ExcelJS-practice/excel_download_test.xlsx");
  const worksheet = workbook.getWorksheet(`Sheet1`);

  //An object to dynamically store a cell coordinate:
  //To begin with, row&col are initialised with "-1":
  let output = { row: -1, column: -1 };
  let targetValue = "Hello";
  //Clementine

  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      //console.log(`Row ${rowNumber}, Col ${colNumber}, Value: ${cell.value}`);

      //Print the location of a specific cell value:
      if (cell.value === targetValue) {
        console.log(`"${targetValue}" found in row ${rowNumber}, column ${colNumber}`);

        //Assign the coordinates of "targetValue" to the "output" object properties:
        output.row = rowNumber;
        output.column = colNumber;
      }
    });
  });

  //Replace the value of the cell:
  const cellToReplace = worksheet.getCell(output.row, output.column);
  cellToReplace.value = "Clementine";

  //Save the file after making the modification:
  await workbook.xlsx.writeFile("C:/Users/Roscoe/Desktop/Projects/ExcelJS-practice/excel_download_test.xlsx");
}

outputCellValues();
