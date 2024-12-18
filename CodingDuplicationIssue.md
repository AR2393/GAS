# GAS
Issue with Google Apps Script Duplication Code
This is my [sheet]([url](https://docs.google.com/spreadsheets/d/1fzhZmeAmPSPuk3oSPG24looe_1WHV7sq0Y0P6bj2doQ/edit?usp=sharing) and the code I'm using will be pasted below, but basically I have a button on each of my 8 tabs that is connected to the script and when I hit the button on any tab it should duplicate the table I have on there which includes formulas, data validation, formatting, column titles, etc. For the most part it works well, but after the first time I hit the button to duplicate the table, meaning once I hit the button a second time, it creates more than 1 duplicate table so the second time it will create 2 duplicates the 3rd time it will create 4 the 4th time it will create 8 and so on. I'm not sure why this is happening, but it's frustating being that is the only issue. If someone can please help determine why that is happening and help me with an updated code that will work. I will be so grateful! There should be access to my sheet as well.

function dupTableForSheet(sheetName) {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    Logger.log(`Sheet ${sheetName} not found!`);
    return;
  }

  // Define the range for the header and table data
  const tableTitleRow = 3; // Title row of the table
  const tableHeaderRow = 4; // Column headers start at row 4
  const tableStartRow = tableHeaderRow;
  const tableColumns = sheet.getLastColumn(); // Get the last column of data in the sheet
  const tableEndRow = sheet.getLastRow(); // Get the last row of data in the sheet

  // Find the last table's position by checking titles in the first column
  let lastTableRow = tableEndRow;
  const titlePrefix = "Table"; // Customize if necessary

  // Loop through the rows to find the last table based on its title in column 1
  for (let row = lastTableRow; row >= tableTitleRow; row--) {
    const cellValue = sheet.getRange(row, 1).getValue();
    if (cellValue && cellValue.startsWith(titlePrefix)) {
      lastTableRow = row; // Last table's row found
      break;
    }
  }

  // Calculate the next available row (add 5 rows after the last table's position)
  const nextRow = lastTableRow + 5;

  // Check if the space for the new table is empty (no data or table)
  const nextTableRange = sheet.getRange(nextRow, 1, tableEndRow - tableStartRow + 1, tableColumns);
  const nextTableValues = nextTableRange.getValues();
  const isSpaceAvailable = nextTableValues.every(row => row.every(cell => cell === ""));

  if (!isSpaceAvailable) {
    Logger.log("Space already occupied by another table or data. No new table created.");
    return; // Exit the function if the space is occupied
  }

  // Now, copy the entire range for the content, including data and formatting
  const tableRange = sheet.getRange(tableStartRow, 1, tableEndRow - tableStartRow + 1, tableColumns);
  tableRange.copyTo(sheet.getRange(nextRow, 1), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

  // Copy the title (row 3) separately to maintain formatting
  const titleRange = sheet.getRange(tableTitleRow, 1, 1, tableColumns);
  titleRange.copyTo(sheet.getRange(nextRow - 1, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // Apply header formatting (copying background color, text formatting, etc.)
  const headerRange = sheet.getRange(tableHeaderRow, 1, 1, tableColumns);
  headerRange.copyTo(sheet.getRange(nextRow + 1, 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // Ensure columns 1 and 7 in the newly duplicated table do not have data validation
  let newTableRange = sheet.getRange(nextRow, 1, tableEndRow - tableStartRow + 1, tableColumns);
  let firstColumnRange = newTableRange.offset(0, 0, newTableRange.getNumRows(), 1);
  let seventhColumnRange = newTableRange.offset(0, 6, newTableRange.getNumRows(), 1);
  firstColumnRange.clearDataValidations(); // Clear validation from the first column
  seventhColumnRange.clearDataValidations(); // Clear validation from the seventh column

  // Update formulas in column E for the new rows (dynamically adjusting the C column reference)
  const newTableEndRow = nextRow + (tableEndRow - tableStartRow);

  // Loop through each row in the newly copied table and set the formula for column E
  for (let i = 0; i < newTableEndRow - nextRow; i++) {
    const formulaCell = sheet.getRange(nextRow + i, 5); // Column E
    const rowNumber = nextRow + i; // Dynamic row number for each new row
    const formula = `=MULTIPLY($C${rowNumber}, D${rowNumber})`; // Reference the specific row for C and D
    formulaCell.setFormula(formula); // Set the formula for each row dynamically
  }

  // Apply subtotal formula, excluding the last row in the new table (for column E)
  const subtotalFormulaRange = sheet.getRange(newTableEndRow, 5);
  subtotalFormulaRange.setFormula(`=SUBTOTAL(9, E${nextRow + 1}:E${newTableEndRow - 1})`);

  Logger.log(`Table copied to ${sheetName} at row ${nextRow}`);
}

// Functions for specific sheets (no changes here)
function dupTableDowntownQ1() {
  dupTableForSheet('Downtown Internal Events Budget Q1');
}

function dupTableDowntownQ2() {
  dupTableForSheet('Downtown Internal Events Budget Q2');
}

function dupTableDowntownQ3() {
  dupTableForSheet('Downtown Internal Events Budget Q3');
}

function dupTableDowntownQ4() {
  dupTableForSheet('Downtown Internal Events Budget Q4');
}

function dupTableENYQ1() {
  dupTableForSheet('ENY Internal Events Budget Q1');
}

function dupTableENYQ2() {
  dupTableForSheet('ENY Internal Events Budget Q2');
}

function dupTableENYQ3() {
  dupTableForSheet('ENY Internal Events Budget Q3');
}

function dupTableENYQ4() {
  dupTableForSheet('ENY Internal Events Budget Q4');
}
