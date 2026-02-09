// NEW: Define emails that have admin access.
// These users will see ALL data, bypassing the date and other filters.
// IMPORTANT: Every email listed here MUST also have an entry in the 'validUsers' object below.
const SUPER_USERS = [
  "jecastro@megaworld-lifestyle.com",
  "jmpizarro@megaworld-lifestyle.com",
  "rrlandoy@megaworld-lifestyle.com",
  "mppiamonte@megaworld-lifestyle.com",
  "jdramirez.megaworld@gmail.com",
  "burofornisandrea30@gmail.com"
  // Add other admin emails here, separated by commas
];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setTitle("MCD Memo Finder");
}

let cachedSheetData = null;

/**
 * Fetches and processes data from the spreadsheet.
 * Includes a check to see if the user has special access permissions.
 * @param {number[]} selectedColumns Array of column numbers to fetch.
 * @param {number} startRow The row number to start fetching data from.
 * @param {string} userEmail The email of the user to check against the SUPER_USERS list.
 * @returns {Array<Array<string>>} A 2D array of the processed data.
 */
function getSheetData(selectedColumns, startRow, userEmail) {
  // Check if the current user has superuser privileges
  const isSuperUser = SUPER_USERS.includes(userEmail);
  
  const dateColumnNumber = selectedColumns[1]; 
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("importselectedcol");
  const lastRow = sheet.getLastRow();
  startRow = startRow || 1;

  const numRows = lastRow - startRow + 1;
  if (numRows < 1) return [];

  const fullData = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn()).getValues();
  let dataToProcess;

  // Conditionally filter data based on user role
  if (isSuperUser) {
    // If the user is a superuser, bypass filters and use the full dataset
    Logger.log(`User ${userEmail} is a superuser. Loading all ${fullData.length} rows.`);
    dataToProcess = fullData;
  } else {
    // Regular users now see all dates. Only hiding rows where Column 8 is empty.
    Logger.log(`User ${userEmail} is a regular user. Applying visibility filters.`);

    dataToProcess = fullData.filter(row => {
      // We keep this line so rows with an empty 'Column H' (index 7) stay hidden.
      const col8Valid = row[7] !== undefined && row[7] !== null && String(row[7]).trim() !== '';
      
      // We removed the 'isDateValid' check entirely
      return col8Valid;
    });
  }

  // Map the selected columns from the (now conditionally filtered) data
  return dataToProcess.map(row => {
    while (row.length < Math.max(...selectedColumns)) {
      row.push('');
    }
    return selectedColumns.map(colIndex => {
      const cell = row[colIndex - 1];
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

/**
 * MODIFIED: Validates user credentials and now also returns their role.
 */
/**
 * Validates login against the external Google Spreadsheet.
 * Matching: Exact case-sensitivity for passwords.
 */
function validateLogin(username, password) {
  const spreadsheetId = "1dBO8ThI7FEKb24D9sPVWokfXLuWUx5aCQvisrT9wBvI";
  const tabName = "MRA";

  if (!username || !password) {
    return { valid: false };
  }

  // Handle common domain typos for the username only (Emails/IDs are usually case-insensitive)
  const correctedUsername = username.toLowerCase().replace(/lifetyle\.com|lifestye\.com/, "lifestyle.com").trim();

  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(tabName);
    const data = sheet.getDataRange().getValues();

    // Map headers
    const headers = data[0];
    const col = {
      userId: headers.indexOf("UserIDfinder"),
      password: headers.indexOf("passwordfinder"),
      actualName: headers.indexOf("actualName"),
      
    };

    // Skip the header row and loop
    for (let i = 1; i < data.length; i++) {
      const row = data[i];

      // Convert DB values to strings to prevent issues with numeric-only passwords
      
      const dbUserId = String(row[col.userId]).toLowerCase().trim();
      
      /** 
       * EXACT PASSWORD MATCHING LOGIC:
       * We do NOT .toLowerCase() the password. 
       * We use .trim() to remove accidental hidden spaces in the Google Sheet cell,
       * but the character casing (A vs a) and symbols (@, #) remain strict.
       */
      const dbPassword = String(row[col.password]).trim();

      // Check username match (ID or Email)
      const isUsernameMatch = (correctedUsername === dbUserId);
      
      // Check password match (Exact Match)
      const isPasswordMatch = (password === dbPassword);

      if (isUsernameMatch && isPasswordMatch) {
        
        // Define Super User Logic
        const admins = typeof SUPER_USERS !== 'undefined' ? SUPER_USERS : [];
        const isSuperUser = admins.includes(correctedUsername);

        return {
          valid: true,
          actualName: row[col.actualName],
          isSuperUser: isSuperUser
        };
      }
    }
  } catch (err) {
    Logger.log("Critical Auth Error: " + err.message);
    return { valid: false, error: "Database Connection Error" };
  }

  return { valid: false };
}
