// File link: https://docs.google.com/spreadsheets/d/1tgjl8T-291AuM1DFaj0W1sTwI7A3lkhugKPVoVpu2lM/edit?gid=0#gid=0
const SPREADSHEET_ID = "1tgjl8T-291AuM1DFaj0W1sTwI7A3lkhugKPVoVpu2lM";

const RESTRICTED_USERS = [
    "Jaye Trich Pizarro",
    "Brix Valdenarro",
    "Joy Melitante",
    "Alex Flores",
    "Aurora Pastolero",
    "Mariano Caleja",
    "Ernesto Andrade",
    "Dustin Sta. Maria",
    "Jenel Ann Cruzgarcia",
    "Doreen Penilla",
    "Denisse Malong",
    "Kevin Lin",
    "Juvi Masilungan",
    "Jhoanalyn Gatdula",
    "Mary Arceo",
    "Annalee Pine ",
    "Arralen Batallones",
    "Ma Dulce Cuenca ",
    "Aizelle Anne Yalong",
    "Henedina Viado",
    "Oauie Banagan",
    "Michelle Ong",
    "Ma Fatima Bausin",
    "Melanie Lingon",
    "Stephen Sumilang",
    "Patricia Mari Quierez",
    "Rosa Cecilia Salvador",
    "Janice Cadog",
    "Camsy Elvina",
    "Mikee Vivo",
    "Vanessa Vicente",
    "Tyron Tan",
    "Rodalyn Landoy",
    "Fermila Pacampara",
    "Lorraine Cicat",
];

function isUserPrivileged(displayName) {
    if (!displayName) return false;
    return !RESTRICTED_USERS.includes(String(displayName).trim());
}

function doGet() {
    return HtmlService.createHtmlOutputFromFile("index").setTitle("MCD Memo Approval").addMetaTag("viewport", "width=device-width, initial-scale=1").setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// REPLACE this function in your Code.gs
function submitRemoveApprovers(data, user) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);

        const required = ["referenceNumber", "removeApprovers"];
        const sanitizedData = validateAndSanitizeInputs(data, required);
        if (!user || !isUserPrivileged(user.displayName)) {
            throw new Error("Unauthorized: You do not have permission to remove approvers.");
        }
        
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName("Conso");
        if (!sheet) throw new Error("Server Error: Sheet 'Conso' not found. Please contact an administrator.");
        
        Logger.log("Privileged action 'submitRemoveApprovers' initiated by: %s", user.displayName);
        
        const referenceNumber = sanitizedData.referenceNumber;
        const rawApproverValues = sanitizedData.removeApprovers;
        
        const approversToRemove = (rawApproverValues || [])
            .map(val => (val && typeof val === 'object' && val.value) ? val.value.trim() : null)
            .filter(Boolean);

        if (approversToRemove.length === 0) {
            // This is a client-side logic error, but we catch it here as a fallback.
            throw new Error("Invalid Request: No valid approvers were selected for removal.");
        }
        
        const lastRow = sheet.getLastRow();
        const refColumn = sheet.getRange("E2:E" + lastRow).getValues();
        const matchRowIndex = refColumn.findIndex((row) => row[0] === referenceNumber);
        
        if (matchRowIndex === -1) {
            throw new Error(`Data mismatch: Reference number '${referenceNumber}' could not be found. The table may have changed. Please refresh.`);
        }
        
        const actualRow = matchRowIndex + 2;
        Logger.log("Matched reference number at row: " + actualRow);
        
        const startColumn = 16; // Column P
        const totalColumns = 60;
        const targetRange = sheet.getRange(actualRow, startColumn, 1, totalColumns);
        const rowData = targetRange.getValues()[0];
        let removedCount = 0;
        
        for (let i = 0; i < rowData.length; i++) {
            if (approversToRemove.includes(rowData[i])) {
                Logger.log("Removing approver: %s from column index %s", rowData[i], i + startColumn);
                rowData[i] = "";
                removedCount++;
            }
        }
        
        if (removedCount === 0) {
            Logger.log("Warning: None of the specified approvers were found to remove, they may have already been actioned by another user.");
        } else {
            targetRange.setValues([rowData]);
            Logger.log("✅ Successfully removed %s approver(s) from row %s.", removedCount, actualRow);
        }
        
        // Explicitly return success
        return { success: true };

    } catch (error) {
        Logger.log("submitRemoveApprovers ERROR: %s \nStack: %s", error.message, error.stack);
        // Propagate a clean error message to the client.
        throw new Error("Remove Approver Failed: " + error.message);
    } finally {
        lock.releaseLock();
    }
}


// REPLACE this function in your Code.gs
function submitApproverValues(data, user) {
    let emailTasks = [];
    let wasSuccessful = false;
    const addedBy = user.displayName;
    const memoFinderUrl = "https://script.google.com/a/macros/megaworld-lifestyle.com/s/AKfycbymM8ffl8z24fG8zzNcE2_zSxvbR2bcP5BNxthFY9FkPremFC--A-zi3TRQOjUKh1Wq/exec";

    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);

        const required = ["referenceNumber", "approverValues", "memoSubject"];
        const sanitizedData = validateAndSanitizeInputs(data, required);

        if (!user || !isUserPrivileged(user.displayName)) {
            throw new Error("Unauthorized: You do not have permission to add approvers.");
        }

        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName("Conso");
        if (!sheet) throw new Error("Server Error: Sheet 'Conso' not found. Please contact an administrator.");

        Logger.log("Privileged action 'submitApproverValues' initiated by: %s", user.displayName);

        const referenceNumber = sanitizedData.referenceNumber;
        const memoSubject = sanitizedData.memoSubject;
        const normalizedApprovers = (sanitizedData.approverValues || []).filter(val => val && typeof val === "object" && val.value);

        if (normalizedApprovers.length === 0) {
            throw new Error("Invalid Request: No valid approvers provided.");
        }

        const lastRow = sheet.getLastRow();
        const refColumn = sheet.getRange("E2:E" + lastRow).getValues();
        const matchRowIndex = refColumn.findIndex(row => row[0] === referenceNumber);

        if (matchRowIndex === -1) {
            throw new Error(`Data mismatch: Reference number '${referenceNumber}' could not be found. The table may have changed. Please refresh.`);
        }

        const actualRow = matchRowIndex + 2;
        const startColumn = 14; // Column N
        const totalColumns = 67;
        const targetRange = sheet.getRange(actualRow, startColumn, 1, totalColumns);
        const rowData = targetRange.getValues()[0];

        // --- NEW LOGIC: Identify new vs. duplicate approvers ---

        // 1. Create a Set of existing approvers for efficient lookup.
        const existingApproverValues = new Set(rowData.filter(Boolean));

        // 2. Filter the submitted list to get only the ones that are truly new.
        const newApproversToAdd = normalizedApprovers.filter(approver => !existingApproverValues.has(approver.value));
        const duplicateApprovers = normalizedApprovers.filter(approver => existingApproverValues.has(approver.value));

        // Log the findings for easier debugging
        if (duplicateApprovers.length > 0) {
            Logger.log(`Duplicate approvers identified (will be ignored): ${duplicateApprovers.map(a => a.label).join(", ")}`);
        }
        if (newApproversToAdd.length > 0) {
            Logger.log(`New approvers to be added: ${newApproversToAdd.map(a => a.label).join(", ")}`);
        }

          // 3. If there are no new approvers to add, we can stop and inform the user.
        if (newApproversToAdd.length === 0) {
            Logger.log("No new approvers to add. All selected were duplicates.");
            // We release the lock here before returning.
            // CHANGE: Return a custom success 'message' instead of a 'warning'.
            // This tells the frontend to show a green success toast.
            return {
                success: true,
                message: `Success: The selected reviewer(s) were already on the list.`
            };
        }
        
        // --- END NEW LOGIC ---

        // Now, only process the newApproversToAdd
        let insertIndex = 0;
        for (const approver of newApproversToAdd) { // Use the filtered list
            while (insertIndex < rowData.length && rowData[insertIndex]) {
                insertIndex++;
            }
            if (insertIndex >= rowData.length) {
                throw new Error("Process Failed: Not enough empty slots in the sheet to add all selected approvers. Please contact an administrator.");
            }
            rowData[insertIndex] = approver.value;
            // Add the new value to our Set so we don't try to add the same person twice in one request
            existingApproverValues.add(approver.value); 
        }

        targetRange.setValues([rowData]);
        Logger.log("✅ Successfully saved new approvers to the sheet.");
        
        // Prepare email tasks ONLY for the new approvers
        emailTasks = newApproversToAdd.map(approver => ({
            email: approver.email,
            label: approver.label,
            cc: approver.cc || [],
            memoSubject: memoSubject,
            referenceNumber: referenceNumber
        }));
        
        wasSuccessful = true;

    } catch (error) {
        Logger.log("submitApproverValues ERROR inside lock: %s \nStack: %s", error.message, error.stack);
        throw new Error("Submit Approver Failed: " + error.message);
    } finally {
        lock.releaseLock();
        Logger.log("Lock released for 'submitApproverValues'.");
    }
    
    // --- Post-Lock Operations (Emailing) ---
    // This section remains largely the same, but it now only operates on `emailTasks`
    // which was populated exclusively with new approvers.
    if (wasSuccessful && emailTasks.length > 0) {
        Logger.log("Starting email dispatch for %s new approver(s).", emailTasks.length);
        const emailFailures = [];

        emailTasks.forEach((task) => {
            if (task.email && task.label) {
                try {
                    const emailSubject = `For Approval: ${task.memoSubject} (Ref: ${task.referenceNumber})`;
                    const emailBody = `<div style="font-family: Arial, sans-serif; line-height: 1.6;"><p>Hello ${task.label},</p><p>You have been added as a reviewer by <strong>${addedBy}</strong> for the following memo:</p><ul><li><strong>Reference #:</strong> ${task.referenceNumber}</li><li><strong>Subject:</strong> ${task.memoSubject}</li></ul><p>You can view and take action on this document by logging into the <a href="${memoFinderUrl}" target="_blank">MCD Memo Routing System</a>.</p><br/><p>Thank you.<br/>MCD Document Routing System</p></div>`;
                    const mailOptions = { to: task.email, subject: emailSubject, htmlBody: emailBody, name: "MCD Memo Routing Notification" };
                    if (task.cc && Array.isArray(task.cc) && task.cc.length > 0) { mailOptions.cc = task.cc.join(','); }
                    MailApp.sendEmail(mailOptions);
                    let logMessage = `Sent "Added Reviewer" notification to ${task.label} <${task.email}>`;
                    if (mailOptions.cc) { logMessage += ` with CC to: ${mailOptions.cc}`; }
                    Logger.log(logMessage);
                } catch (e) {
                    Logger.log(`CRITICAL WARNING: Failed to send "Added Reviewer" email to ${task.label}. Error: ${e.message}`);
                    emailFailures.push(task.label);
                }
            }
        });

        if (emailFailures.length > 0) {
            return {
                success: true,
                warning: `Approvers added, but failed to send email notifications to: ${emailFailures.join(", ")}. Please notify them manually.`
            };
        }
    }
    
    return { success: true };
}

function getSheetData(sheetName) {
    const MAX_COLUMNS = 80;
    const START_ROW = 1;
    const START_COLUMN = 1;
    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            throw new Error(`Sheet "${sheetName}" not found.`);
        }
        const lastRow = sheet.getLastRow();
        if (lastRow < 1) {
            Logger.log(`Sheet "${sheetName}" is empty.`);
            return null;
        }
        const range = sheet.getRange(START_ROW, START_COLUMN, lastRow, MAX_COLUMNS);
        const data = range.getValues();
        const timeZone = Session.getScriptTimeZone();
        const cleanedData = data.map((row, rowIndex) =>
            row.map((cell, colIndex) => {
                try {
                    if (cell instanceof Date) {
                        return Utilities.formatDate(cell, timeZone, "MMM d, yyyy hh:mm a");
                    }
                    return String(cell).trim();
                } catch (e) {
                    Logger.log(`Cell processing error at row ${rowIndex + 1}, column ${colIndex + 1}: ${e}`);
                    return "";
                }
            })
        );
        const filteredData = cleanedData
            .filter((row, i) => {
                const statusColA = row[0];
                const statusColF = row[5];
                const colJ = row[9];
                const isApprovedOrDisapproved = (status) => status === "APPROVED" || status === "DISAPPROVED";
                if (isApprovedOrDisapproved(statusColA) || isApprovedOrDisapproved(statusColF)) {
                    return false;
                }
                if (!colJ || colJ === "") {
                    return false;
                }
                return row.some((cell) => cell !== "");
            })
            .map((row) => row.slice(4));
        Logger.log(`Filtered ${filteredData.length} valid row(s) from sheet "${sheetName}".`);
        return filteredData.length > 0 ? filteredData : null;
    } catch (error) {
        Logger.log(`Error in getSheetData("${sheetName}"): ${error.message}`);
        return null;
    }
}


// Function to show all Approved Memos


function getApprovedSheetData(sheetName) {
    // --- These constants and the initial setup are identical to getSheetData ---
    const MAX_COLUMNS = 80;
    const START_ROW = 1;
    const START_COLUMN = 1;
    try {
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID); // Assumes SPREADSHEET_ID is a global constant
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            throw new Error(`Sheet "${sheetName}" not found.`);
        }
        const lastRow = sheet.getLastRow();
        if (lastRow < 1) {
            Logger.log(`Sheet "${sheetName}" has no approved items because it is empty.`);
            return null;
        }
        const range = sheet.getRange(START_ROW, START_COLUMN, lastRow, MAX_COLUMNS);
        const data = range.getValues();
        const timeZone = Session.getScriptTimeZone();

        // --- Data cleaning is identical to ensure consistent formatting ---
        const cleanedData = data.map((row, rowIndex) =>
            row.map((cell, colIndex) => {
                try {
                    if (cell instanceof Date) {
                        return Utilities.formatDate(cell, timeZone, "MMM d, yyyy hh:mm a");
                    }
                    return String(cell).trim();
                } catch (e) {
                    Logger.log(`Cell processing error at row ${rowIndex + 1}, column ${colIndex + 1}: ${e}`);
                    return "";
                }
            })
        );
        
        // --- This is the key difference: The filtering logic is changed ---
        const filteredData = cleanedData
            .filter((row, i) => {
                // We are still interested in the same columns for our logic
                const statusColA = row[0]; // From your logic, this is a status column
                const statusColF = row[5]; // This is another status column
                const colJ = row[9];       // This is a required column

                // CORE LOGIC CHANGE: Instead of excluding approved rows, we now *require* them.
                const isApproved = (statusColA === "APPROVED" || statusColF === "APPROVED");
                
                // If the row is not marked as "APPROVED", we exclude it.
                if (!isApproved) {
                    return false;
                }
                
                // We keep the other essential validation checks from your original function.
                if (!colJ || colJ === "") {
                    return false; // A row isn't valid without data in column J
                }
                
                // This ensures we don't return completely blank rows that might somehow pass other checks.
                return row.some((cell) => cell !== "");
            })
            // IMPORTANT: This slicing must be IDENTICAL to the original function so the
            // front-end receives data with the expected structure.
            .map((row) => row.slice(4));

        Logger.log(`Filtered ${filteredData.length} APPROVED row(s) from sheet "${sheetName}".`);
        return filteredData.length > 0 ? filteredData : null;
    } catch (error) {
        Logger.log(`Error in getApprovedSheetData("${sheetName}"): ${error.message}`);
        return null;
    }
}

function updateRowStatusWithValidation(rowData, status, reason, approvalReason, sheetName) {
    try {
        if (!rowData || !Array.isArray(rowData) || !rowData[0]) {
            throw new Error("Row data with a reference number is required.");
        }
        if (!status || String(status).trim() === "") {
            throw new Error("Status (e.g., 'APPROVED') is required.");
        }
        if (!sheetName || String(sheetName).trim() === "") {
            throw new Error("Sheet name is required.");
        }
        const remarks = status === "DISAPPROVED" ? reason : approvalReason;
        if (!remarks || String(remarks).trim() === "") {
            throw new Error("Remarks are required for this action.");
        }
        const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
            throw new Error(`Sheet "${sheetName}" not found.`);
        }
        const data = sheet.getDataRange().getValues();
        const referenceNumberToMatch = String(rowData[0]).trim();
        Logger.log(`Searching for reference: "${referenceNumberToMatch}" in sheet "${sheetName}"`);
        const timestamp = new Date();
        const formattedTimestamp = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "MMM d, yyyy HH:mm");
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            const referenceNumber = String(row[9]).trim();
            if (referenceNumber === referenceNumberToMatch) {
                const rowToUpdate = i + 1;
                let diffDays = "";
                const colBValue = row[1];
                if (colBValue && colBValue instanceof Date) {
                    const diffTime = timestamp.getTime() - colBValue.getTime();
                    diffDays = Math.floor(diffTime / (1000 * 60 * 60 * 24));
                }
                const updateRange = sheet.getRange(rowToUpdate, 1, 1, 4);
                const valuesToWrite = [[status, remarks.trim(), formattedTimestamp, diffDays]];
                updateRange.setValues(valuesToWrite);
                Logger.log(`Match found and row ${rowToUpdate} updated successfully.`);
                return true;
            }
        }
        Logger.log(`No match found for reference number: "${referenceNumberToMatch}"`);
        throw new Error("No matching row found to update.");
    } catch (error) {
        Logger.log("Error in updateRowStatusWithValidation: " + error.message);
        throw new Error("Failed to update row: " + error.message);
    }
}
function authenticateUser(username, password) {
    if (!username || !password) {
        throw new Error("Username or password is missing.");
    }
    const users = {
        GMC: { password: "pass1", displayName: "Graham Coates", sheet: "GMC" },
        GRACE: { password: "grace123", displayName: "Grace Guarico", sheet: "GMC" },
        RCS: { password: "pass2", displayName: "Rosalyn Segura", sheet: "RCS" },
        ALY: { password: "pass2", displayName: "Allyzza Tolosa", sheet: "RCS" },
        JE: { password: "je123", displayName: "Jeremy Rodriguez", sheet: "JE" },
        NOE: { password: "noe123", displayName: "Noe Versoza", sheet: "NOE" },
        MGL: { password: "pass3", displayName: "Michael Lao", sheet: "MGL" },
        MLP: { password: "pass3a", displayName: "Louise Piamonte", sheet: "MGL" },
        KMV: { password: "pass4", displayName: "Tinay Villanueva", sheet: "KMV" },
        ZELFA: { password: "zelfa123", displayName: "Zelfa Valderrama", sheet: "KMV" },
        LDA: { password: "pass5", displayName: "Lorence Aurelio", sheet: "LDA" },
        FA: { password: "pass16", displayName: "Frances Armian", sheet: "LDA" },
        JAYE: { password: "pass6", displayName: "Jaye Trich Pizarro", sheet: "JAYE" },
        JOY: { password: "pass7", displayName: "Jocelyn Melitante", sheet: "JOY" },
        ALEX: { password: "pass8", displayName: "Alex Flores", sheet: "ALEX" },
        AU: { password: "pass9", displayName: "Aurora Palostero", sheet: "AU" },
        MAR: { password: "pass10", displayName: "Mariano Caleja", sheet: "MAR" },
        BRIX: { password: "pass11", displayName: "Brix Valdenarro", sheet: "BRIX" },
        EBA: { password: "pass12", displayName: "Ernesto Andrade", sheet: "EBA" },
        DSM: { password: "pass13", displayName: "Dustin Sta. Maria", sheet: "DSM" },
        JAC: { password: "pass14", displayName: "Jenel Ann Cruzgarcia", sheet: "JAC" },
        DP: { password: "pass15", displayName: "Doreen Penilla", sheet: "DP" },
        DM: { password: "pass17", displayName: "Denisse Malong", sheet: "DM" },
        KL: { password: "pass18", displayName: "Kevin Lin", sheet: "KL" },
        JM: { password: "pass19", displayName: "Juvi Masilungan", sheet: "JM" },
        JG: { password: "pass20", displayName: "Jhoanalyn Gatdula", sheet: "JG" },
        MA: { password: "pass21", displayName: "Mary Arceo", sheet: "MA" },
        JLC: { password: "luther2024", displayName: "Jeron Luther Castro", sheet: "GMC" },
        ANNALEE: { password: "pass22", displayName: "Annalee Pine ", sheet: "ANNALEE" },
        ARRA: { password: "pass23", displayName: "Arralen Batallones", sheet: "ARRA" },
        DULCE: { password: "pass24", displayName: "Ma Dulce Cuenca", sheet: "DULCE" },
        AIZELLE: { password: "pass25", displayName: "Aizelle Anne Yalong", sheet: "AIZELLE" },
        VIA: { password: "pass26", displayName: "Henedina Viado", sheet: "VIA" },
        OAUIE: { password: "pass27", displayName: "Oauie Banagan", sheet: "OAUIE" },
        MICH: { password: "pass28", displayName: "Michelle Ong", sheet: "MICH" },
        FATIMA: { password: "pass29", displayName: "Ma Fatima Bausin", sheet: "FATIMA" },
        MELANIE: { password: "pass30", displayName: "Melanie Lingon", sheet: "MELANIE" },
        STEPH: { password: "pass31", displayName: "Stephen Sumilang", sheet: "STEPH" },
        PAT: { password: "pass32", displayName: "Patricia Mari Quierez", sheet: "PAT" },
        ROSA: { password: "pass33", displayName: "Rosa Cecilia Salvador", sheet: "ROSA" },
        JANICE: { password: "pass34", displayName: "Janice Cadog", sheet: "JANICE" },
        CAMSY: { password: "camsy123", displayName: "Camsy Elvina", sheet: "CAMSY" },
        LORRAINE: { password: "lorraine123", displayName: "Lorraine Cicat", sheet: "CAMSY" },
        MIKEE: { password: "mikee123", displayName: "Mikee Vivo", sheet: "MIKEE" },
        FERMIE: { password: "fermie123", displayName: "Fermila Pacampara", sheet: "MIKEE" },
        VANESSA: { password: "vanessa123", displayName: "Vanessa Vicente", sheet: "VANESSA" },
        TYRON: { password: "tyron123", displayName: "Tyron Tan", sheet: "TYRON" },
        RODA: { password: "roda123", displayName: "Rodalyn Landoy", sheet: "TYRON" },
        GLORIE: { password: "glorie123", displayName: "Glorie Alynna Racelis", sheet: "MGL" },
    };
    const user = users[username];
    if (!user || user.password !== password) {
        throw new Error(`Unauthorized Account.`);
    }
    Logger.log(`Authenticated: ${user.displayName}`);
    return { displayName: user.displayName, sheet: user.sheet };
}



function validateAndSanitizeInputs(data, requiredFields = []) {
    if (!data || typeof data !== "object") {
        throw new Error("Invalid data payload: must be an object.");
    }
    const sanitizedData = {};
    for (const field of requiredFields) {
        if (!data[field] || String(data[field]).trim() === "") {
            const fieldName = field.replace(/([A-Z])/g, " $1").replace(/^./, (str) => str.toUpperCase());
            throw new Error(`Validation failed: ${fieldName} is a required field and cannot be empty.`);
        }
    }
    for (const key in data) {
        if (Object.prototype.hasOwnProperty.call(data, key)) {
            const value = data[key];
            if (typeof value === "string") {
                sanitizedData[key] = value.trim();
            } else if (Array.isArray(value)) {
                sanitizedData[key] = value.map((item) => (typeof item === "string" ? item.trim() : item));
            } else {
                sanitizedData[key] = value;
            }
        }
    }
    return sanitizedData;
}
function ping() {
    Logger.log("User activity ping received at: " + new Date());
    return "backend OK";
}




