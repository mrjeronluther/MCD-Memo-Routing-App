# MCD Memo Routing Web App

## Introduction

A web application built with Google Apps Script to submit, route, and approve MCD memos/documents using Google Sheets as the backend database. The application is intended for internal use within the MCD organization — for memo submitters, approvers, and administrators who manage workflows, tracking, and audit logs.

---

## Installation Instructions

1. Prerequisites
   - Google account with access to Google Apps Script and Google Drive
   - Permissions to create and edit Google Sheets in your Google Workspace
   - Optional: clasp (command-line tool) if you prefer local development

2. Create the Google Sheets database
   - Create a new Google Sheet to act as the database.
   - Add the following sheets (tabs): `Memos`, `Approvals`, `Users` (optional), `Status Log` (optional).

3. Create a Google Apps Script project
   - Go to https://script.google.com and create a new project (or use clasp locally).
   - Name it e.g., "MCD Memo Routing App".

4. Add project files
   - Copy the `.gs`, `.html`, and `.json` files from this repository into your Apps Script project.
   - If using clasp, clone/pull the repo and push files via clasp.

5. Configure the project
   - Update constants in the `.gs` files (SHEET_ID, sheet names, admin emails, etc.).
   - Confirm manifest (appsscript.json) runtime is V8 and required scopes are present.

6. Authorize and test
   - Run a setup or helper function once in the Apps Script editor to accept required scopes.
   - Verify the script can read/write the target Google Sheet.

7. Deploy as Web App
   - In Apps Script editor: Deploy → New deployment → Web app.
   - Set "Execute as" and "Who has access" according to your org policy.
   - Deploy and copy the web app URL for users.

---

## Usage Examples

1) Serve the landing page (Apps Script doGet)
```javascript
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Landing Page')
    .setTitle('MCD Memo Routing')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

2) Minimal addMemo example (append a new memo)

function addMemo(memoData) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(MEMO_SHEET_NAME);
  const id = Utilities.getUuid();
  const row = [
    id,
    memoData.title || '',
    memoData.content || '',
    memoData.submitter || '',
    memoData.department || '',
    new Date(),
    memoData.priority || 'Medium',
    'Pending',
    memoData.routingPath || ''
  ];
  sheet.appendRow(row);
  return id;
}

3. Minimal routeMemo example (set next approver & status)


function routeMemo(memoId, nextApproverEmail) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(MEMO_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === memoId) { // assuming ID is in column A
      sheet.getRange(i + 1, /*col for Status*/ 8).setValue('Pending ' + nextApproverEmail);
      sheet.getRange(i + 1, /*col for Current Approver*/ 9).setValue(nextApproverEmail);
      return true;
    }
  }
  return false;
}


Notes:

The above snippets are simplified examples. Confirm column indices and exact function names in your repository before integrating them verbatim.
Update SHEET_ID, MEMO_SHEET_NAME, and other constants before running.
For MCD Internal Use Only
