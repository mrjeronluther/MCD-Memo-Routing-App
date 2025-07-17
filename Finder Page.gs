function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle("MCD Memo Finder");
}

let cachedSheetData = null;

function getSheetData(selectedColumns, startRow) {
  // --- FIX: Dynamically identify the date column ---
  // It is the second item in your selectedColumns array: [8, 14, 2, ...]
  // So selectedColumns[1] will be 14.
  const dateColumnNumber = selectedColumns[1]; 

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("importselectedcol");
  const lastRow = sheet.getLastRow();
  startRow = startRow || 1;

  const numRows = lastRow - startRow + 1;
  if (numRows < 1) return [];

  const fullData = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  const threeMonthsAgo = new Date();
  threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3);

  // --- FIX: Filter using the dynamic dateColumnNumber ---
  const filteredData = fullData.filter(row => {
    const col8Valid = row[7] !== undefined && row[7] !== null && String(row[7]).trim() !== '';
    
    // Get the date from the correct column (e.g., column 14, which is index 13)
    const dateValue = row[dateColumnNumber - 1]; 
    const dateInColumn = dateValue instanceof Date ? dateValue : null;
    const isDateValid = dateInColumn ? dateInColumn >= threeMonthsAgo : true;

    return col8Valid && isDateValid;
  });

  return filteredData.map(row => {
    while (row.length < Math.max(...selectedColumns)) {
      row.push('');
    }

    return selectedColumns.map(colIndex => {
      const cell = row[colIndex - 1];

      // --- FIX: Format the date using the dynamic dateColumnNumber ---
      if (colIndex === dateColumnNumber && cell instanceof Date) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), "MMM d, yyyy H:mm:ss");
      }

      return cell;
    });
  });
}

function clearSheetDataCache() {
  cachedSheetData = null;
}

function validateLogin(username, password) {
  // Map username to { password, actualName }
  const validUsers = {
    "admin": { password: "password123", actualName: "Jeron Luther Castro" },
    "user1": { password: "password123", actualName: "Uploader Name" },
    "GLORIE": { password: "glorie123", actualName: "Glorie Alynna Racelis" },
  };

  if (validUsers[username] && validUsers[username].password === password) {
    return {
      valid: true,
      actualName: validUsers[username].actualName
    };
  } else {
    return {
      valid: false
    };
  }
}
