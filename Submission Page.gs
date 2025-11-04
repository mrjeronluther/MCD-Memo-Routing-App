// NEW: Storing IDs in Script Properties is a best practice for security and maintainability.
// To set these, go to Project Settings (gear icon) > Script properties.
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const CONFIG = {
  DRIVE_FOLDER_ID: SCRIPT_PROPS.getProperty("DRIVE_FOLDER_ID") || "1tfPCtJmoFpONmX3MgbSfCe6w8_AIDJQw",
  SPREADSHEET_ID: SCRIPT_PROPS.getProperty("SPREADSHEET_ID") || "1tgjl8T-291AuM1DFaj0W1sTwI7A3lkhugKPVoVpu2lM",
  REF_SHEET_ID: SCRIPT_PROPS.getProperty("REF_SHEET_ID") || "1jACmopzPgV2dPYEbDQwI8iJmIT4SgTHw4L-CXBTHOG4",
  /**
   * NEW: Specify the name of the sheet in the REF_SHEET_ID spreadsheet
   * that contains the headers (Row 1) and corresponding emails (Row 2).
   */
  EMAIL_LOOKUP_SHEET_NAME: "DefaultEmail",
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag("viewport", "width=device-width, initial-scale=1").setTitle("MCD Memo Routing");
}

/**
 * NEW: This function matches selected options to headers in a reference sheet
 * and retrieves corresponding emails from the row below the headers.
 * @param {string[]} additionalOptions - An array of strings selected by the user.
 * @returns {string[]} An array of unique email addresses found.
 */
function getEmailsFromHeaders(additionalOptions) {
    if (!additionalOptions || additionalOptions.length === 0) {
        return [];
    }
    if (!CONFIG.REF_SHEET_ID || !CONFIG.EMAIL_LOOKUP_SHEET_NAME) {
        console.error("Email lookup Spreadsheet ID or Sheet Name is not configured.");
        return [];
    }

    try {
        const spreadsheet = SpreadsheetApp.openById(CONFIG.REF_SHEET_ID);
        const sheet = spreadsheet.getSheetByName(CONFIG.EMAIL_LOOKUP_SHEET_NAME);
        if (!sheet) {
            console.error(`Sheet named "${CONFIG.EMAIL_LOOKUP_SHEET_NAME}" not found in the reference spreadsheet.`);
            return [];
        }

        const lastColumn = sheet.getLastColumn();
        if (lastColumn === 0) return []; // Sheet is empty

        // Read headers from Row 1 and emails from Row 2
        const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
        const emails = sheet.getRange(2, 1, 1, lastColumn).getValues()[0];

        const headerToEmailMap = new Map();
        headers.forEach((header, index) => {
            // Only map if the header and email both exist and are not empty strings
            if (header && String(header).trim() && emails[index] && String(emails[index]).trim()) {
                headerToEmailMap.set(String(header).trim(), String(emails[index]).trim());
            }
        });

        const emailsToNotify = new Set();
        additionalOptions.forEach(option => {
            const cleanOption = option.trim();
            if (headerToEmailMap.has(cleanOption)) {
                emailsToNotify.add(headerToEmailMap.get(cleanOption));
            }
        });

        console.log("Found matching emails for notification: ", Array.from(emailsToNotify));
        return Array.from(emailsToNotify);

    } catch (e) {
        console.error(`Error in getEmailsFromHeaders: ${e.toString()}`);
        return []; // Return empty array on error to prevent process failure
    }
}


function sanitizeString(str, type = "general") {
  if (!str || typeof str !== "string") return "";
  if (type === "mimetype") return str.replace(/[^a-zA-Z0-9./-]/g, "").trim();
  str = str.replace(/[<>&"'`]/g, ""); // Prevent script injection
  if (type === "email") return str.replace(/[^a-zA-Z0-9@._,\s-]/g, "").trim();
  if (type === "date") return str.replace(/[^a-zA-Z0-9\/\-\s]/g, "").trim();
  return str.replace(/[^a-zA-Z0-9 .,\\-_]/g, "").replace(/\s+/g, " ").trim();
}

function checkSheetLimit(sheet) {
  const MAXROWS = 9900000;
  if (sheet.getLastRow() >= MAXROWS) {
    throw new Error(`The '${sheet.getName()}' sheet is full. Please contact the administrator to archive old data.`);
  }
}

// REPLACE this function in your Code.gs
// This version uses our robust reservation pattern to safely write variable data into separate columns,
// and separates the email notification process to run AFTER the lock is released for high concurrency.
function processFormData(data) {
    let mainFile = null;
    let revisedFile = null;
    let wasSuccessful = false; // NEW: Flag to track if the critical section succeeded.

    if (!CONFIG.DRIVE_FOLDER_ID || !CONFIG.SPREADSHEET_ID) {
        throw new Error("Administrator setup error: Script properties for folder/sheet ID are not configured.");
    }

    try {
        // --- Step 1: Pre-computation and File Creation (Outside the lock) ---
        // These operations are safe to run concurrently. File creation is inherently atomic.
        validateFormServerSide(data);
        const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

        const sanitized = {
            name: sanitizeString(data.name),
            email: sanitizeString(data.email, "email"),
            email1: sanitizeString(data.email1, "email"),
            memoType: sanitizeString(data.memoType),
            dateType: sanitizeString(data.dateType, "date"),
            additionalMemoInfo: sanitizeString(data.additionalMemoInfo),
            otherExtraOption: sanitizeString(data.otherExtraOption),
            fileName: sanitizeString(data.fileName),
            fileMimeType: sanitizeString(data.fileMimeType, "mimetype"),
            revisedFileName: sanitizeString(data.revisedFileName || ""),
            revisedFileMimeType: sanitizeString(data.revisedFileMimeType || "", "mimetype"),
            additionalOptions: (data.additionalOptions || []).map(sanitizeString),
            isResubmission: !!data.isResubmission,
            fileData: data.fileData,
            revisedFileData: data.revisedFileData || null,
            approversToNotify: (data.approversToNotify || []).map(a => ({
                name: sanitizeString(a.name),
                email: sanitizeString(a.email, "email"),
                value: sanitizeString(a.value),
                cc: (a.cc || []).map(cc => sanitizeString(cc, "email")).filter(Boolean)
            })),
            selectedCGM: (data.selectedCGM || []).map(sanitizeString),
            selectedGhead: (data.selectedGhead || []).map(sanitizeString),
            selectedDept: (data.selectedDept || []).map(sanitizeString),
            selectedDiv: (data.selectedDiv || []).map(sanitizeString),
        };
        
        // NEW: Get additional emails by matching headers from the reference sheet
        sanitized.additionalEmails = getEmailsFromHeaders(sanitized.additionalOptions);


        const mainBlob = Utilities.newBlob(Utilities.base64Decode(sanitized.fileData), sanitized.fileMimeType, sanitized.fileName);
        mainFile = folder.createFile(mainBlob);

        if (sanitized.isResubmission && sanitized.revisedFileData) {
            const revisedBlob = Utilities.newBlob(Utilities.base64Decode(sanitized.revisedFileData), sanitized.revisedFileMimeType, sanitized.revisedFileName);
            revisedFile = folder.createFile(revisedBlob);
        }

        const refNumber = generateReferenceNumber();
        const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName("Conso");
        if (!sheet) throw new Error('Sheet named "Conso" does not exist.');

        checkSheetLimit(sheet);

          
        const allReviewers = [...sanitized.selectedCGM, ...sanitized.selectedGhead, ...sanitized.selectedDept, ...sanitized.selectedDiv];
        const fixedData = [
            refNumber, new Date(), sanitized.name, sanitized.email, sanitized.memoType,
            sanitized.email1, sanitized.additionalMemoInfo, sanitized.dateType,
            mainFile.getName(), mainFile.getUrl()
        ];

        // --- Step 2: Critical Section (Spreadsheet writing ONLY) ---
        const lock = LockService.getScriptLock();
        lock.waitLock(30000);

        try {
            // This entire block is now extremely fast, minimizing lock time.

            // **FIXED CODE: Find the first blank row from Col E to CD**
            const startRow = 1; // Assuming data starts from row 1
            const maxRows = sheet.getMaxRows();
            const range = sheet.getRange(startRow, 1, maxRows - startRow + 1, 82); // From Col A to CD (82 columns)
            const values = range.getValues();
            let newRowNumber = -1;

            for (let i = 0; i < values.length; i++) {
                if (values[i].every(cell => cell === "")) {
                    newRowNumber = i + startRow;
                    break;
                }
            }

            if (newRowNumber === -1) {
                newRowNumber = sheet.getLastRow() + 1; // Fallback to appending if no empty row is found
            }
            
            const placeholderValue = `reserving_${refNumber}_${new Date().getTime()}`;
            const placeholderRange = sheet.getRange(newRowNumber, 4); 
            placeholderRange.setValue(placeholderValue);
            SpreadsheetApp.flush();
            if (placeholderRange.getValue() !== placeholderValue) {
                throw new Error("Critical concurrency error: Failed to secure a unique row. Please try again.");
            }

            sheet.getRange(newRowNumber, 1).setValue("PENDING");

            const fixedDataRange = sheet.getRange(newRowNumber, 5, 1, fixedData.length);
            fixedDataRange.setValues([fixedData]);
            let lastColumn = fixedDataRange.getLastColumn();

            if (sanitized.additionalOptions.length > 0) {
                const propertiesRange = sheet.getRange(newRowNumber, lastColumn + 1, 1, sanitized.additionalOptions.length);
                propertiesRange.setValues([sanitized.additionalOptions]);
                lastColumn = propertiesRange.getLastColumn();
            }

            const otherPropertyRange = sheet.getRange(newRowNumber, lastColumn + 1);
            otherPropertyRange.setValue(sanitized.otherExtraOption);
            lastColumn = otherPropertyRange.getColumn();

            if (allReviewers.length > 0) {
                const reviewersRange = sheet.getRange(newRowNumber, lastColumn + 1, 1, allReviewers.length);
                reviewersRange.setValues([allReviewers]);
                lastColumn = reviewersRange.getLastColumn();
            }

            const resubmissionRange = sheet.getRange(newRowNumber, lastColumn + 1, 1, 2);
            resubmissionRange.setValues([
                [
                    sanitized.isResubmission && revisedFile ? revisedFile.getName() : '',
                    sanitized.isResubmission && revisedFile ? revisedFile.getUrl() : ''
                ]
            ]);

            placeholderRange.clearContent();
            
            // NEW: Set the success flag ONLY if all spreadsheet operations complete without error.
            wasSuccessful = true;

        } finally {
            // --- Step 3: Release the Lock ---
            // This is CRITICAL. The lock is now released before the slow email process begins.
            lock.releaseLock();
        }

        // --- Step 4: Post-Lock Operations (Slow tasks like emailing) ---
        // This section only runs if the critical section above succeeded.
        // Other users can now acquire the lock while this email is being sent.
        if (wasSuccessful) {
            sendNotificationEmail(sanitized, refNumber, mainFile, revisedFile);
            return { refNumber }; // Return the success object to the client.
        }

    } catch (err) {
        // This catch block handles errors from ANY of the steps above.
        // It also ensures orphaned files are cleaned up if an error occurs after they are created.
        if (mainFile) {
            try { mainFile.setTrashed(true); console.log(`CLEANUP: Trashed orphaned file: ${mainFile.getName()}`); } catch (e) {}
        }
        if (revisedFile) {
            try { revisedFile.setTrashed(true); console.log(`CLEANUP: Trashed orphaned file: ${revisedFile.getName()}`); } catch (e) {}
        }

        console.error("Processing error: " + err.stack);
        throw new Error("Server error during processing: " + err.message);
    }
}

function sendNotificationEmail(sanitized, refNumber, mainAttachment, revisedAttachment) {
    // Check if there are any approvers or additional emails to notify
    if ((!sanitized.approversToNotify || sanitized.approversToNotify.length === 0) && (!sanitized.additionalEmails || sanitized.additionalEmails.length === 0)) {
        console.log("No recipients found for notification email. Skipping send.");
        return;
    }
    
    const subject = `For Approval: ${sanitized.additionalMemoInfo} (Ref: ${refNumber})`;
    const formattedDate = Utilities.formatDate(new Date(sanitized.dateType), Session.getScriptTimeZone(), "MMM d, yyyy");
    
    const createApproverList = (title, selectedValues, allApprovers) => {
      const names = selectedValues
        .map(value => allApprovers.find(a => a.value === value)?.name)
        .filter(Boolean).map(name => `<li>${name}</li>`).join("");
      return names ? `<h4 style="margin-bottom: 5px; color: #333;">${title}:</h4><ul style="margin-top: 0; padding-left: 20px;">${names}</ul>` : "";
    };

    const approverListHtml = createApproverList("CGM/s", sanitized.selectedCGM, sanitized.approversToNotify)
                           + createApproverList("Group Head/s", sanitized.selectedGhead, sanitized.approversToNotify)
                           + createApproverList("Department Head/s", sanitized.selectedDept, sanitized.approversToNotify)
                           + createApproverList("Division Head/s", sanitized.selectedDiv, sanitized.approversToNotify);

    const masterCcList = new Set([sanitized.email1]);
    sanitized.approversToNotify.forEach(a => {
        if(a.email) masterCcList.add(a.email);
        if(a.cc) a.cc.forEach(cc => masterCcList.add(cc));
    });
    
    // NEW: Add the emails found from the header matching to the CC list
    if (sanitized.additionalEmails && sanitized.additionalEmails.length > 0) {
        sanitized.additionalEmails.forEach(email => {
            if (email) { // Ensure the email is not null/empty
                masterCcList.add(email);
            }
        });
    }


    const memmofinder = "https://script.google.com/a/macros/megaworld-lifestyle.com/s/AKfycbwFAhQ7gw4-b5VLc8ksZz2BmBmB6Fquo22eZ9nWs_CmJ04gJIoqYFWsdxyMLv7h0LpH/exec";

    const mainFileUrl = mainAttachment.getUrl();
    const revisedFileUrl = revisedAttachment ? revisedAttachment.getUrl() : null;

    const htmlBody = `
        <div style="font-family: Arial, sans-serif; line-height: 1.6;">
          <p>Hello ${sanitized.name},</p>
          <p>You have successfully submitted a document. A notification has been sent to all required approvers, who are included on this email's CC list. Please see the full list of approvers below:</p>
          <div style="border: 1px solid #eee; padding: 10px 15px; border-radius: 8px; background-color: #f9f9f9;">
            ${approverListHtml}
          </div>
           <p>You may track the document status here: <a href="${memmofinder}" target="_blank">MCD Document Routing</a></p>
          <p><b>Approvers:</b> Please review the document(s) for your approval using the links below.</p>
          <div style="padding-left: 20px;">
            <p style="margin: 5px 0;">
                <a href="${mainFileUrl}" target="_blank" style="font-weight: bold;">View Main Document: ${sanitized.fileName}</a>
            </p>
            ${revisedFileUrl ? 
              `<p style="margin: 5px 0;">
                  <a href="${revisedFileUrl}" target="_blank" style="font-weight: bold;">View Revised Document: ${sanitized.revisedFileName}</a>
               </p>` 
              : ''
            }
          </div>
          <br>
          <table style="border-collapse: collapse; width: 100%; max-width: 600px; border: 1px solid #dddddd;">
            <tr style="border-bottom: 1px solid #dddddd;"><th style="padding: 8px; background-color: #f2f2f2; text-align: left; border-right: 1px solid #dddddd;">Reference #</th><td style="padding: 8px;">${refNumber}</td></tr>
            <tr style="border-bottom: 1px solid #dddddd;"><th style="padding: 8px; background-color: #f2f2f2; text-align: left; border-right: 1px solid #dddddd;">Subject</th><td style="padding: 8px;">${sanitized.additionalMemoInfo}</td></tr>
            <tr style="border-bottom: 1px solid #dddddd;"><th style="padding: 8px; background-color: #f2f2f2; text-align: left; border-right: 1px solid #dddddd;">Submitted By</th><td style="padding: 8px;">${sanitized.name} (${sanitized.email})</td></tr>
            <tr style="border-bottom: 1px solid #dddddd;"><th style="padding: 8px; background-color: #f2f2f2; text-align: left; border-right: 1px solid #dddddd;">Memo Preparer</th><td style="padding: 8px;">${sanitized.memoType} (${sanitized.email1})</td></tr>
            <tr style="border-bottom: 1px solid #dddddd;"><th style="padding: 8px; background-color: #f2f2f2; text-align: left; border-right: 1px solid #dddddd;">Memo Date</th><td style="padding: 8px;">${formattedDate}</td></tr>
          </table>
          <br>
         
          <p>Thank you.<br/>MCD Document Routing System</p>
        </div>`;

    try {
        MailApp.sendEmail({
            to: sanitized.email,
            cc: Array.from(masterCcList).filter(Boolean).join(","),
            subject: subject,
            htmlBody: htmlBody,
            name: "MCD Document for Approval",
        });
    } catch (err) {
        console.error(`Failed to send master notification for Ref ${refNumber}. Error: ${err.message}`);
        try {
            const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
            let logSheet = spreadsheet.getSheetByName("EmailFailures");
            if (!logSheet) {
              logSheet = spreadsheet.insertSheet("EmailFailures");
              logSheet.appendRow(["Timestamp", "Ref Number", "Recipient", "Error Message"]);
            }
            logSheet.appendRow([new Date(), refNumber, sanitized.email, err.message]);
        } catch (logErr) {
            console.error(`Failed to even log the email failure. Logging Error: ${logErr.message}`);
        }
    }
}


function validateFormServerSide(data) {
  // This validation is now fine. It acts as a server-side guard against excessively large
  // uploads that might cause timeouts, but it no longer conflicts with email limits.
  const MAX_SIZE = 32 * 1024 * 1024;
  const base64Size = str => (str ? (str.length * 3) / 4 - (str.match(/=+$/)?.[0].length || 0) : 0);

  if (!data.name || !data.email || !data.memoType || !data.email1 || !data.dateType || !data.additionalMemoInfo) throw new Error("A required information field is missing.");
  if ((data.additionalOptions || []).length === 0 && !data.otherExtraOption) throw new Error("At least one Property/Coverage must be selected or specified.");
  if ((data.selectedCGM || []).length === 0 || (data.selectedGhead || []).length === 0 ||  (data.selectedDept || []).length === 0 || (data.selectedDiv || []).length === 0) throw new Error("An approver must be selected from each group (CGM, Dept, Div).");
  
  if (!data.fileData) throw new Error("Primary file data is missing.");
  if (base64Size(data.fileData) > MAX_SIZE) throw new Error("Primary file size exceeds 32 MB.");
  if (data.isResubmission && !data.revisedFileData) throw new Error("Revised file is required for resubmission.");
  if (data.revisedFileData && base64Size(data.revisedFileData) > MAX_SIZE) throw new Error("Revised file size exceeds 32 MB.");

  ['name','memoType','otherExtraOption','fileName','revisedFileName'].forEach(f => {
    if (data[f] && data[f].length > 255) throw new Error(`${f} exceeds max length.`);
  });
  if(data.additionalMemoInfo && data.additionalMemoInfo.length > 500) throw new Error("Subject of memo exceeds max length.");
}

function generateReferenceNumber() {
  if (!CONFIG.REF_SHEET_ID) {
    throw new Error("Administrator setup error: Reference number log sheet ID is not configured.");
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // Wait up to 30 seconds
  try {
    const spreadsheet = SpreadsheetApp.openById(CONFIG.REF_SHEET_ID);
    let sheet = spreadsheet.getSheetByName("referencenumber");
    if (!sheet) {
      sheet = spreadsheet.insertSheet("referencenumber");
      sheet.appendRow(["Generated Reference Numbers"]);
    }
    
    let existingRefs = new Set();
    if (sheet.getLastRow() > 1) {
      existingRefs = new Set(sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat());
    }

    let newRef, attempts = 0;
    do {
      newRef = `REF#${Math.random().toString(36).substring(2, 8).toUpperCase()}`;
      if (++attempts > 20) throw new Error("Could not generate a unique reference number after 20 attempts.");
    } while (existingRefs.has(newRef));
    
    sheet.appendRow([newRef]);
    return newRef;

  } catch (e) {
    const message = `Critical Error: Could not generate a reference number. ${e.message}`;
    Logger.log(message);
    throw new Error(message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Fetches unique property values from the dvPayorCompany sheet
 * and formats them for the memo routing form.
 *
 * @returns {Array<Object>} An array of objects, each with a 'label' and a 'value'.
 */
function getMemoProperties() {
  try {
    const ss = SpreadsheetApp.openById('1qheN_KURc-sOKSngpzVxLvfkkc8StzGv-1gMvGJZdsc');
    const sheet = ss.getSheetByName('dvPayorCompany');

    if (!sheet) {
      throw new Error('Sheet "dvPayorCompany" not found.');
    }

    const lastRow = sheet.getLastRow();
    // If there are no data rows (header is row 1 & 2, data starts at 3), return an empty array.
    if (lastRow < 3) {
      return [];
    }

    // Get all values from column C, starting from row 3.
    const range = sheet.getRange('C3:C' + lastRow);
    const values = range.getValues();

    // The 'values' variable is a 2D array like [['Value1'], ['Value2'], ['']].
    // 1. .flat() converts it to a 1D array: ['Value1', 'Value2', ''].
    // 2. .filter(String) removes any empty or blank cells.
    // 3. new Set(...) automatically gets only the unique values.
    // 4. [...new Set(...)] converts the Set back into an array.
    const uniqueLabels = [...new Set(values.flat().filter(String))];

    // Map the labels to the required {label, value} format.
    const options = uniqueLabels.map(label => {
      // Create a short 'value' from the 'label' by taking the first letter of each word.
      // Example: "Alabang West Parade" becomes "AWP".
      const words = label.trim().split(/\s+/);
      let value = words.map(word => word[0]).join('').toUpperCase();
      
      // Handle cases where the value might not be unique or is too short
      if (value.length < 2 && label.length > 1) {
         value = label.substring(0, 3).toUpperCase();
      }

      return { label: label.trim(), value: value };
    });

    // Sort the final array alphabetically by label for a better user experience.
    return options.sort((a, b) => a.label.localeCompare(b.label));

  } catch (e) {
    // Log the error for debugging and return a clear error to the frontend.
    console.error('Error fetching memo properties: ' + e.toString());
    throw new Error('Could not retrieve the properties list from the spreadsheet. Please check sheet name and permissions.');
  }
}

