// NEW: Storing IDs in Script Properties is a best practice for security and maintainability.
// To set these, go to Project Settings (gear icon) > Script properties.
const SCRIPT_PROPS = PropertiesService.getScriptProperties();
const CONFIG = {
  DRIVE_FOLDER_ID: SCRIPT_PROPS.getProperty("DRIVE_FOLDER_ID") || "1tfPCtJmoFpONmX3MgbSfCe6w8_AIDJQw",
  SPREADSHEET_ID: SCRIPT_PROPS.getProperty("SPREADSHEET_ID") || "1tgjl8T-291AuM1DFaj0W1sTwI7A3lkhugKPVoVpu2lM",
  REF_SHEET_ID: SCRIPT_PROPS.getProperty("REF_SHEET_ID") || "1jACmopzPgV2dPYEbDQwI8iJmIT4SgTHw4L-CXBTHOG4",
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile("index").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag("viewport", "width=device-width, initial-scale=1").setTitle("MCD Memo Routing");
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

function processFormData(data) {
  let mainFile = null;
  let revisedFile = null;
  
  if (!CONFIG.DRIVE_FOLDER_ID || !CONFIG.SPREADSHEET_ID) {
    throw new Error("Administrator setup error: Script properties for folder/sheet ID are not configured.");
  }

  try {
    validateFormServerSide(data);
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);

    // Sanitize all incoming data
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
        name: sanitizeString(a.name), email: sanitizeString(a.email, "email"),
        value: sanitizeString(a.value), cc: (a.cc || []).map(cc => sanitizeString(cc, "email")).filter(Boolean)
      })),
      selectedCGM: (data.selectedCGM || []).map(sanitizeString),
      selectedDept: (data.selectedDept || []).map(sanitizeString),
      selectedDiv: (data.selectedDiv || []).map(sanitizeString),
    };
    
    const mainBlob = Utilities.newBlob(Utilities.base64Decode(sanitized.fileData), sanitized.fileMimeType, sanitized.fileName);
    mainFile = folder.createFile(mainBlob);
    const mainFileUrl = mainFile.getUrl();

    let revisedFileUrl = "";
    if (sanitized.isResubmission && sanitized.revisedFileData) {
        const revisedBlob = Utilities.newBlob(Utilities.base64Decode(sanitized.revisedFileData), sanitized.revisedFileMimeType, sanitized.revisedFileName);
        revisedFile = folder.createFile(revisedBlob);
        revisedFileUrl = revisedFile.getUrl();
    }

    const refNumber = generateReferenceNumber();
    const spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = spreadsheet.getSheetByName("Conso");
    if (!sheet) throw new Error('Sheet named "Conso" does not exist.');

    checkSheetLimit(sheet);

    const approverCodesForSheet = [...sanitized.selectedCGM, ...sanitized.selectedDept, ...sanitized.selectedDiv];
    let rowData = [refNumber, new Date(), sanitized.name, sanitized.email, sanitized.memoType, sanitized.email1, sanitized.additionalMemoInfo, sanitized.dateType, sanitized.fileName, mainFileUrl];
    
    rowData.push(...sanitized.additionalOptions, sanitized.otherExtraOption, ...approverCodesForSheet);

    if (sanitized.isResubmission && sanitized.revisedFileName && revisedFileUrl) {
      rowData.push(sanitized.revisedFileName, revisedFileUrl);
    }
    sheet.getRange(sheet.getLastRow() + 1, 5, 1, rowData.length).setValues([rowData]);

    // This function call is unchanged; it correctly passes the Drive File objects.
    sendNotificationEmail(sanitized, refNumber, mainFile, revisedFile);
    
    return { refNumber };

  } catch (err) {
    if (mainFile) {
        console.error(`PROCESS FAILED AFTER MAIN FILE CREATION. File to review/delete: ${mainFile.getName()} (ID: ${mainFile.getId()}). Original Error: ${err.message}`);
    }
    if (revisedFile) {
        console.error(`PROCESS FAILED AFTER REVISED FILE CREATION. File to review/delete: ${revisedFile.getName()} (ID: ${revisedFile.getId()}). Original Error: ${err.message}`);
    }
    console.error("Processing error: " + err.stack);
    throw new Error("Server error during processing: " + err.message);
  }
}


// =========================================================================
// ===                    START: MODIFIED CODE SECTION                   ===
// =========================================================================

function sendNotificationEmail(sanitized, refNumber, mainAttachment, revisedAttachment) {
    if (!sanitized.approversToNotify || sanitized.approversToNotify.length === 0) return;
    
    const subject = `For Approval: ${sanitized.additionalMemoInfo} (Ref: ${refNumber})`;
    const formattedDate = Utilities.formatDate(new Date(sanitized.dateType), Session.getScriptTimeZone(), "MMM d, yyyy");
    
    const createApproverList = (title, selectedValues, allApprovers) => {
      const names = selectedValues
        .map(value => allApprovers.find(a => a.value === value)?.name)
        .filter(Boolean).map(name => `<li>${name}</li>`).join("");
      return names ? `<h4 style="margin-bottom: 5px; color: #333;">${title}:</h4><ul style="margin-top: 0; padding-left: 20px;">${names}</ul>` : "";
    };

    const approverListHtml = createApproverList("CGM/s", sanitized.selectedCGM, sanitized.approversToNotify)
                           + createApproverList("Department Head/s", sanitized.selectedDept, sanitized.approversToNotify)
                           + createApproverList("Division Head/s", sanitized.selectedDiv, sanitized.approversToNotify);

    const masterCcList = new Set([sanitized.email1]);
    sanitized.approversToNotify.forEach(a => {
        if(a.email) masterCcList.add(a.email);
        if(a.cc) a.cc.forEach(cc => masterCcList.add(cc));
    });

    const memmofinder = "https://script.google.com/a/macros/megaworld-lifestyle.com/s/AKfycbwFAhQ7gw4-b5VLc8ksZz2BmBmB6Fquo22eZ9nWs_CmJ04gJIoqYFWsdxyMLv7h0LpH/exec";

    // --- CHANGE #1: Get URLs from the Drive file objects. ---
    const mainFileUrl = mainAttachment.getUrl();
    const revisedFileUrl = revisedAttachment ? revisedAttachment.getUrl() : null;

    // --- CHANGE #2: Modify the HTML body to include links instead of mentioning attachments. ---
    const htmlBody = `
        <div style="font-family: Arial, sans-serif; line-height: 1.6;">
          <p>Hello ${sanitized.name},</p>
          <p>You have successfully submitted a document. A notification has been sent to all required approvers, who are included on this email's CC list. Please see the full list of approvers below:</p>
          <div style="border: 1px solid #eee; padding: 10px 15px; border-radius: 8px; background-color: #f9f9f9;">
            ${approverListHtml}
          </div>
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
          <p>You may track the document status here: <a href="${memmofinder}" target="_blank">MCD Document Routing</a></p>
          <p>Thank you.<br/>MCD Document Routing System</p>
        </div>`;

    try {
        // --- CHANGE #3: Remove the 'attachments' property from the sendEmail options. ---
        MailApp.sendEmail({
            to: sanitized.email,
            cc: Array.from(masterCcList).filter(Boolean).join(","),
            subject: subject,
            htmlBody: htmlBody,
            // attachments: attachments, // <-- THIS LINE IS REMOVED
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
// =========================================================================
// ===                     END: MODIFIED CODE SECTION                    ===
// =========================================================================


function validateFormServerSide(data) {
  // This validation is now fine. It acts as a server-side guard against excessively large
  // uploads that might cause timeouts, but it no longer conflicts with email limits.
  const MAX_SIZE = 32 * 1024 * 1024;
  const base64Size = str => (str ? (str.length * 3) / 4 - (str.match(/=+$/)?.[0].length || 0) : 0);

  if (!data.name || !data.email || !data.memoType || !data.email1 || !data.dateType || !data.additionalMemoInfo) throw new Error("A required information field is missing.");
  if ((data.additionalOptions || []).length === 0 && !data.otherExtraOption) throw new Error("At least one Property/Coverage must be selected or specified.");
  if ((data.selectedCGM || []).length === 0 || (data.selectedDept || []).length === 0 || (data.selectedDiv || []).length === 0) throw new Error("An approver must be selected from each group (CGM, Dept, Div).");
  
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
