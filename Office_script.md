# Office Script language (only to be used online)

At the time of writing this doc, office script language us available only on Excel Online.
Office Script takes syntax from **TypeScript**

## Read and log one cell

This sample reads the value of A1 and prints it to the console.
```
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();
// Get the value of cell A1.
  let range = selectedSheet.getRange("A1");
  
  // Print the value of A1.
  console.log(range.getValue());
}
```

## Read the active cell

This script logs the value of the current active cell. If multiple cells are selected, the top-leftmost cell will be logged.
TypeScriptCopy
```
function main(workbook: ExcelScript.Workbook) {
  // Get the current active cell in the workbook.
  let cell = workbook.getActiveCell();
// Log that cell's value.
  console.log(`The current cell's value is ${cell.getValue()}`);
}
```

## Change an adjacent cell

This script gets adjacent cells using relative references. Note that if the active cell is on the top row, part of the script fails, because it references the cell above the currently selected one.
```
function main(workbook: ExcelScript.Workbook) {
  // Get the currently active cell in the workbook.
  let activeCell = workbook.getActiveCell();
  console.log(`The active cell's address is: ${activeCell.getAddress()}`);
// Get the cell to the right of the active cell and set its value and color.
  let rightCell = activeCell.getOffsetRange(0,1);
  rightCell.setValue("Right cell");
  console.log(`The right cell's address is: ${rightCell.getAddress()}`);
  rightCell.getFormat().getFont().setColor("Magenta");
  rightCell.getFormat().getFill().setColor("Cyan");
// Get the cell to the above of the active cell and set its value and color.
  // Note that this operation will fail if the active cell is in the top row.
  let aboveCell = activeCell.getOffsetRange(-1, 0);
  aboveCell.setValue("Above cell");
  console.log(`The above cell's address is: ${aboveCell.getAddress()}`);
  aboveCell.getFormat().getFont().setColor("White");
  aboveCell.getFormat().getFill().setColor("Black");
}
```

## Change all adjacent cells

This script copies the formatting in the active cell to the neighboring cells. Note that this script only works when the active cell isn't on an edge of the worksheet.

```
function main(workbook: ExcelScript.Workbook) {
  // Get the active cell.
  let activeCell = workbook.getActiveCell();
// Get the cell that's one row above and one column to the left of the active cell.
  let cornerCell = activeCell.getOffsetRange(-1,-1);
// Get a range that includes all the cells surrounding the active cell.
  let surroundingRange = cornerCell.getResizedRange(2, 2)
// Copy the formatting from the active cell to the new range.
  surroundingRange.copyFrom(
    activeCell, /* The source range. */
    ExcelScript.RangeCopyType.formats /* What to copy. */
    );
}
```

## Change each individual cell in a range

This script loops over the currently select range. It clears the current formatting and sets the fill color in each cell to a random color.
```
function main(workbook: ExcelScript.Workbook) {
  // Get the currently selected range.
  let range = workbook.getSelectedRange();
// Get the size boundaries of the range.
  let rows = range.getRowCount();
  let cols = range.getColumnCount();
// Clear any existing formatting
  range.clear(ExcelScript.ClearApplyTo.formats);
// Iterate over the range.
  for (let row = 0; row < rows; row++) {
    for (let col = 0; col < cols; col++) {
      // Generate a random color hex-code.
      let colorString = `#${Math.random().toString(16).substr(-6)}`;
// Set the color of the current cell to that random hex-code.
      range.getCell(row, col).getFormat().getFill().setColor(colorString);
    }
  }
}
```

## Get groups of cells based on special criteria

This script gets all the blank cells in the current worksheet's used range. It then highlights all those cells with a yellow background.
```
function main(workbook: ExcelScript.Workbook) {
    // Get the current used range.
    let range = workbook.getActiveWorksheet().getUsedRange();
    
    // Get all the blank cells.
    let blankCells = range.getSpecialCells(ExcelScript.SpecialCellType.blanks);
// Highlight the blank cells with a yellow background.
    blankCells.getFormat().getFill().setColor("yellow");
}
```

## Collections

These samples work with collections of objects in the workbook.

## Iterate over collections

This script gets and logs the names of all the worksheets in the workbook. It also sets the their tab colors to a random color.
```
function main(workbook: ExcelScript.Workbook) {
  // Get all the worksheets in the workbook.
  let sheets = workbook.getWorksheets();
// Get a list of all the worksheet names.
  let names = sheets.map ((sheet) => sheet.getName());
// Write in the console all the worksheet names and the total count.
  console.log(names);
  console.log(`Total worksheets inside of this workbook: ${sheets.length}`);
  
  // Set the tab color each worksheet to a random color
  for (let sheet of sheets) {
    // Generate a random color hex-code.
    let colorString = `#${Math.random().toString(16).substr(-6)}`;
// Set the color of the current worksheet's tab to that random hex-code.
    sheet.setTabColor(colorString);
  }
}
```

## Query and delete from a collection

This script creates a new worksheet. It checks for an existing copy of the worksheet and deletes it before making a new sheet.
```
function main(workbook: ExcelScript.Workbook) {
  // Name of the worksheet to be added.
  let name = "Index";
// Get any worksheet with that name.
  let sheet = workbook.getWorksheet("Index");
  
  // If `null` wasn't returned, then there's already a worksheet with that name.
  if (sheet) {
    console.log(`Worksheet by the name ${name} already exists. Deleting it.`);
    // Delete the sheet.
    sheet.delete();
  }
  
  // Add a blank worksheet with the name "Index".
  // Note that this code runs regardless of whether an existing sheet was deleted.
  console.log(`Adding the worksheet named ${name}.`);
  let newSheet = workbook.addWorksheet("Index");
// Switch to the new worksheet.
  newSheet.activate();
}
```

## Dates

The samples in this section show how to use the JavaScript Date object.
The following sample gets the current date and time and then writes those values to two cells in the active worksheet.
```
function main(workbook: ExcelScript.Workbook) {
  // Get the cells at A1 and B1.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
  let timeRange = workbook.getActiveWorksheet().getRange("B1");
// Get the current date and time with the JavaScript Date object.
  let date = new Date(Date.now());
// Add the date string to A1.
  dateRange.setValue(date.toLocaleDateString());
// Add the time string to B1.
  timeRange.setValue(date.toLocaleTimeString());
}
```

The next sample reads a date that's stored in Excel and translates it to a JavaScript Date object. It uses the date's numeric serial number as input for the JavaScript Date.
```
function main(workbook: ExcelScript.Workbook) {
  // Read a date at cell A1 from Excel.
  let dateRange = workbook.getActiveWorksheet().getRange("A1");
// Convert the Excel date to a JavaScript Date object.
  let excelDateValue = dateRange.getValue() as number;
  let javaScriptDate = new Date(Math.round((excelDateValue - 25569) * 86400 * 1000));
  console.log(javaScriptDate);
}
```

## Display data

These samples demonstrate how to work with worksheet data and provide users with a better view or organization.

## Apply conditional formatting

This sample applies conditional formatting to the currently used range in the worksheet. The conditional formatting is a green fill for the top 10% of values.
```
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();
// Get the used range in the worksheet.
  let range = selectedSheet.getUsedRange();
// Set the fill color to green for the top 10% of values in the range.
  let conditionalFormat = range.addConditionalFormat(ExcelScript.ConditionalFormatType.topBottom)
  conditionalFormat.getTopBottom().getFormat().getFill().setColor("green");
  conditionalFormat.getTopBottom().setRule({
    rank: 10, // The percentage threshold.
    type: ExcelScript.ConditionalTopBottomCriterionType.topPercent // The type of the top/bottom condition.
  });
}
```

## Create a sorted table

This sample creates a table from the current worksheet's used range, then sorts it based on the first column.
```
function main(workbook: ExcelScript.Workbook) {
  // Get the current worksheet.
  let selectedSheet = workbook.getActiveWorksheet();
// Create a table with the used cells.
  let usedRange = selectedSheet.getUsedRange();
  let newTable = selectedSheet.addTable(usedRange, true);
// Sort the table using the first column.
  newTable.getSort().apply([{ key: 0, ascending: true }]);
}
```

From <https://docs.microsoft.com/en-us/office/dev/scripts/resources/samples/excel-samples>


## Sample code: Insert comma-separated values into a workbook [read a CSV File/Content]

```
function main(workbook: ExcelScript.Workbook, csv: string) {
  /* Convert the CSV data into a 2D array. */
  // Trim the trailing new line.
  csv = csv.trim();
// Split each line into a row.
  let rows = csv.split("\r\n");
  let data : string[][] = [];
  rows.forEach((value) => {
    /*
     * For each row, match the comma-separated sections.
     * For more information on how to use regular expressions to parse CSV files,
     * see this Stack Overflow post: https://stackoverflow.com/a/48806378/9227753
     */
    let row = value.match(/(?:,|\n|^)("(?:(?:"")*[^"]*)*"|[^",\n]*|(?:\n|$))/g);
    
    // Remove the preceding comma.
    row.forEach((cell, index) => {
      row[index] = cell.indexOf(",") === 0 ? cell.substr(1) : cell;
    });
    data.push(row);
  });
// Put the data in the worksheet.
  let sheet = workbook.getWorksheet("Sheet1");
  let range = sheet.getRangeByIndexes(0, 0, data.length, data[0].length);
  range.setValues(data);
// Add any formatting or table creation that you want.
}
```
From <https://docs.microsoft.com/en-us/office/dev/scripts/resources/samples/convert-csv>

