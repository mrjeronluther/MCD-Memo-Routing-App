MCD Document Routing System
Introduction

The MCD Document Routing System is a bespoke, Google Apps Script-powered workflow automation tool designed specifically for the MCD team. Its primary purpose is to digitize and streamline the traditional "inter-office envelope" process.

The system replaces manual signature chasing and physical paper shuffling with a secure, web-based platform that provides:

Submission Portal: A centralized form for uploading memos and assigning a routing path (CGMs, Group Heads, and Dept/Div Heads).

Approval Portal: A secure dashboard for managers to review documents, provide remarks, and approve/disapprove items.

Document Finder: A real-time tracking tool for requestors to check the current status and audit trail of any submission.

Automated Audit Trail: Every action is timestamped and logged, generating a final "Approval Page" PDF upon completion.

Installation Instructions

This project is built using Google Apps Script (GAS), Vue.js 3, and Bootstrap 5. To deploy this system, follow these steps:

Google Sheet Setup:

Create a Google Spreadsheet to act as your database.

Create a sheet named Conso for the master log.

Create a second spreadsheet to act as the Reference Log and create a sheet named referencenumber and DefaultEmail.

Google Drive Setup:

Create a dedicated Google Drive folder where all uploaded memos will be stored. Note the Folder ID.

Script Deployment:

Open script.google.com and create three separate projects (Submission, Approver, and Finder).

Copy the corresponding .html and .js code provided into each project.

In the Submission Page project, go to Project Settings and add the following Script Properties:

DRIVE_FOLDER_ID: Your Drive Folder ID.

SPREADSHEET_ID: Your Master Spreadsheet ID.

REF_SHEET_ID: Your Reference/Email Spreadsheet ID.

Permissions & Deployment:

Click Deploy > New Deployment.

Select Web App.

Set Execute as: Me (The Admin).

Set Who has access: Anyone within [Your Domain].

Authorize the script to access Drive, Sheets, and Gmail.

Usage Examples

Below is a "copy-paste ready" code snippet demonstrating how the frontend interacts with the backend logic. This pattern is used to submit data from the browser to the Google Sheet while handling the asynchronous response.

code
JavaScript
download
content_copy
expand_less
/**
 * Example: Triggering a Document Submission
 * This snippet demonstrates how the Vue.js frontend sends the 
 * sanitized payload to the processFormData Google Script function.
 */

const submitToMCDSystem = (payload) => {
  console.log("Initiating submission for:", payload.additionalMemoInfo);

  // Calling the Google Apps Script Backend
  google.script.run
    .withSuccessHandler((result) => {
      // result contains the generated Reference Number
      alert(`Success! Document logged as: ${result.refNumber}`);
      console.log("Confirmation details:", result);
    })
    .withFailureHandler((error) => {
      console.error("Submission failed:", error.message);
      alert("Error: " + error.message);
    })
    .processFormData(payload); 
    // 'processFormData' is the .gs function that handles File Upload & Sheet Logging
};

// Example Payload Structure
const samplePayload = {
  name: "John Doe",
  email: "jdoe@megaworld-lifestyle.com",
  memoType: "Jane Smith",
  dateType: "2025-06-15",
  additionalMemoInfo: "Request for Budget Realignment",
  selectedCGM: ["CAMSY"],
  fileData: "BASE64_ENCODED_STRING_HERE",
  fileMimeType: "application/pdf",
  fileName: "Budget_Memo.pdf"
};

For MCD Internal Use Only
