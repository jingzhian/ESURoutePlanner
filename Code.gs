function debugMessage(obj) {
  SpreadsheetApp.getUi().alert(JSON.stringify(obj, null, 2));
}

function getColumns(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function getColumn(columns, pattern) {
  return columns.findIndex(name => name.toLowerCase().includes(pattern.toLowerCase())) + 1;
}

// Get the first Sheet from the active spreadsheet that contains `name`
// if sheet does not exist, return undefined
function getSheet(name) {
  const sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  const matches = sheets.filter(sheet => sheet.getName().toLowerCase().includes(name.toLowerCase()));
  if (matches.length > 0) {
    return matches[0];
  } 
  return undefined;
}

function updateChanges() {
  
  // get the route and changes spreadsheets
  const routeSheet = getSheet("route");
  const changeSheet = getSheet("change");
  
  // get the column in the route sheet contains "Changes"
  const routeColumns = getColumns(routeSheet);
  const routeChangeColumn = getColumn(routeColumns, "change");
  const routeLocationColumn = getColumn(routeColumns, "loc");
  const routeInitialColumn = getColumn(routeColumns, "init");

  // get the list of changes from the changes column
  const numPatients = routeSheet.getLastRow() - 1; // -1 for header
  const changes = [];
  for (let i = 0; i < numPatients; i++) {
    const row = i + 2; // +1 for one-indexing and +1 for header
    const description = routeSheet.getRange(row, routeChangeColumn, 1, 1).getValue();
    const location = routeSheet.getRange(row, routeLocationColumn, 1, 1).getValue();
    const initial = routeSheet.getRange(row, routeInitialColumn, 1, 1).getValue();
    const lines = description.split("\n");
    lines.forEach((line, index) => {
      if (description.trim().length > 0) {
        changes.push({
          row: row,
          done: line.includes("✅"),
          description: line.replaceAll("✅", "").trim(),
          location: location,
          initial: initial,
          index: index,
        })
      }
    });
  }
  // throw(JSON.stringify(changes, null, 2));
  changeSheet.getRange(2, 1, changes.length, 1).insertCheckboxes();
  changeSheet.getRange(2, 1, changes.length, 1).setValues(changes.map(change => [change.done ? "true" : "false"]));
  changeSheet.getRange(2, 2, changes.length, 1).setValues(changes.map(change => [change.description]));
  changeSheet.getRange(2, 3, changes.length, 1).setValues(changes.map(change => [change.location]));
  changeSheet.getRange(2, 4, changes.length, 1).setValues(changes.map(change => [change.initial]));
  changeSheet.getRange(2, 5, changes.length, 1).setValues(changes.map(change => [`${change.row}-${change.index}`]));
  changeSheet.hideColumns(5, 1);
  if (changeSheet.getLastRow() > changes.length + 2) {
    changeSheet.deleteRows(changes.length + 2, changeSheet.getLastRow() - (changes.length + 2));
  }
  changeSheet.getRange(1, 2, 5, changeSheet.getLastRow()).protect().setDescription("This table is generated automatically");
}

function onEdit(e) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();
  const columns = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]; 
  const column = e.range.getColumn();
  const columnName = columns.length >= column ? columns[column - 1] : "";

  // Add a time stamp in the column behind a checkbox
  if (sheetName.toLowerCase().includes("route") && columnName.toLowerCase().includes("seen")) {
    if (e.value.toUpperCase() == 'TRUE') {
      e.range.offset(0, 1).setValue(new Date()).setNumberFormat("mm/dd h:mm:ss am/pm");
    } else if (e.value.toUpperCase() == 'FALSE'){
      e.range.offset(0, 1).clearContent();
    }
  }
  
  // change '...' to '; with new line in column called 'Changes'
  if (sheetName.toLowerCase().includes("route") && columnName.toLowerCase().includes("changes")) {
    const value = e.value || "";
    const tokens = value.split(/(?:\.\.\.|…)/);
    const newValue = tokens.map(token => token.trim()).join("\n");
    e.range.offset(0, 0).setValue(newValue);
    updateChanges();
  }

  // synchronize checkbox in change list 
  if (sheetName.toLowerCase().includes("change") && columnName.toLowerCase().includes("done")) {
    // get the route and changes spreadsheets
    const routeSheet = getSheet("route");
    const routeColumns = getColumns(routeSheet);
    const changeColumn = getColumn(routeColumns, "change");
    const changeSheet = getSheet("change");
    const changeColumns = getColumns(changeSheet);
    const indexColumn = getColumn(changeColumns, "index");
    const indexValue = changeSheet.getRange(e.range.getRow(), indexColumn, 1, 1).getValue();
    const routeRow = parseInt(indexValue.split("-")[0], 10);
    const lineIndex = parseInt(indexValue.split("-")[1], 10);
    const routeChangeValue = routeSheet.getRange(routeRow, changeColumn, 1, 1).getValue();
    const lines = routeChangeValue.split("\n");
    if (e.value.toUpperCase() == 'TRUE') {
      if (!lines[lineIndex].includes("✅")) {
        lines[lineIndex] = "✅" + lines[lineIndex];
        const newValue = lines.join("\n");
        routeSheet.getRange(routeRow, changeColumn, 1, 1).setValue(newValue);
      }
    } else if (e.value.toUpperCase() == 'FALSE') {
      if (lines[lineIndex].includes("✅")) {
        lines[lineIndex] = lines[lineIndex].replaceAll("✅", "");
        const newValue = lines.join("\n");
        routeSheet.getRange(routeRow, changeColumn, 1, 1).setValue(newValue);
      }
    }
  }
}
