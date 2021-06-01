# Google AppsScript Sheets cookbook
[API reference](https://developers.google.com/apps-script/reference/spreadsheet)

# Basics
### Google Sheets naming conventions
A **spreadsheet** is the name of the overall Google Sheets file. 
A **sheet** is the name of an individual page/tab within the spreadsheet.

### Loading a spreadsheet by id
Where the id can be found in the url - docs.google.com/spreadsheets/d/**spreadsheetId**/edit#gid=0
```js
const spreadsheet = SpreadsheetApp.openById("abcd1234");
```

### Loading a sheet by name
```js
const spreadsheet = SpreadsheetApp.openById("abcd1234");
const sheet = spreadsheet.getSheetByName("Sheet1");
```


---

# Reading data

### EXAMPLE SPREADSHEET
| Game      | Release date | Console  |
| --------- | ------------ | -------- |
| Mario     | 1985         | NES      |
| Runescape | 2001         | PC       |
| Pokemon   | 1996         | Game Boy |


### Reading all data on the sheet
```js
const range = sheet.getDataRange();  
range.getValues();
=>
[ [ 'Game', 'Year', 'Console' ],
  [ 'Mario', 1985, 'NES' ],
  [ 'Runescape', 2001, 'PC' ],
  [ 'Pokemon', 1996, 'Game Boy' ] ]
```
Note: This is functionally equivalent to creating a Range bounded by A1 and (`Sheet.getLastColumn()`, `Sheet.getLastRow()`)

### Reading specific parts of the sheet
**Reading a single cell**
```js
// Using indices
getRange(row, column);
const range = sheet.getRange(2, 1);
range.getValues();
=> [ [ 'Mario' ] ]

// Using A1 notation
const range = sheet.getRange("A2");
range.getValues();
=> [ [ 'Mario' ] ]
```
**Reading multiple cells**
```js
// Using indices
getRange(row, column, numRows, numCols);
const range = sheet.getRange(3, 1, 2, 2);
range.getValues();
=> [ [ 'Runescape', 2001 ], [ 'Pokemon', 1996 ] ]

// Using A1 notation
const range = sheet.getRange("A3:B4");
range.getValues();
=> [ [ 'Runescape', 2001 ], [ 'Pokemon', 1996 ] ]
```

### Reading a column given the column header name
```js
/** Returns a column (array) given the header name of the column. */
function getColByHeader(sheet, headerName, headerRowIndex = 1) {
  const headerRow = sheet.getRange(headerRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headerRow.indexOf(headerName);
  if (colIndex < 0) {
    throw `Could not find header: ${headerName}!`;
  }
  return sheet.getRange(headerRowIndex+1, colIndex+1, sheet.getLastRow()-headerRowIndex, 1).getValues().flat();
}

getColByHeader(sheet, 'Console');
=> [ 'NES', 'PC', 'Game Boy' ]
```

---

# Writing data

### Clearing a sheet of all contents and formatting information
```js
sheet.clear();
```

### Appending a new row to the sheet
```js
// Appends a new row with 3 columns to the bottom of the
// spreadsheet containing the values in the array
sheet.appendRow(["a man", "a plan", "panama"]);
```

### Writing to a specific part of the sheet
**Writing to a single cell**
```js
let cell = sheet.getRange("B2");
cell.setValue(100);
```
**Writing to multiple cells**
```js
// The size of the two-dimensional array must match the size of the range.
let values = [
  [ "2.000", "1,000,000", "$2.99" ]
];
let range = sheet.getRange("B2:D2");
range.setValues(values);
```
**Clearing cells**
```js
let range = sheet.getRange("B2:D2");
range.clear();
```

### Getting the index of a column
```js
/** Returns a column index (0 based) given the header name of the column. */
function getColByHeader(sheet, headerName, headerRowIndex = 1) {
  const headerRow = sheet.getRange(headerRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headerRow.indexOf(headerName);
  if (colIndex < 0) {
    throw `Could not find header: ${headerName}!`;
  }
  return colIndex;
}
```
