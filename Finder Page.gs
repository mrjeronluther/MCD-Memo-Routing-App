function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle("MCD Memo Finder");
}

let cachedSheetData = null;

function getSheetData(selectedColumns, startRow) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("importselectedcol");
  const lastRow = sheet.getLastRow();

  // Default to row 1 if not specified
  startRow = startRow || 1;

  const numRows = lastRow - startRow + 1;
  if (numRows < 1) return [];

  const fullData = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  const threeMonthsAgo = new Date();
  threeMonthsAgo.setMonth(threeMonthsAgo.getMonth() - 3); // Calculate date limit

  // Only include rows where column 8 (index 7) has a value AND column 94 (index 93) is within the last 3 months
  const filteredData = fullData.filter(row => {
    const col8Valid = row[7] !== undefined && row[7] !== null && String(row[7]).trim() !== '';
    const col94Date = row[93] instanceof Date ? row[93] : null; // Ensure col 94 is a valid Date
    const col94Valid = col94Date ? col94Date >= threeMonthsAgo : true; // Exclude older dates

    return col8Valid && col94Valid;
  });

  return filteredData.map(row => {
    // Ensure all rows are padded for selected column indexes
    while (row.length < Math.max(...selectedColumns)) {
      row.push('');
    }

    return selectedColumns.map(colIndex => {
      const cell = row[colIndex - 1];

      // Check if colIndex is 94 and convert to string if it's a date
      if (colIndex === 94 && cell instanceof Date) {
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





