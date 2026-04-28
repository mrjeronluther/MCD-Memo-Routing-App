# MCD Document Routing System

## Introduction
The **MCD Document Routing System** is an automated, web-based workflow solution built on the Google Workspace ecosystem. It is designed to replace traditional paper-based memo routing with a high-visibility, digital process.

### Core Features:
*   **Submission Portal:** Allows users to upload PDFs/Images and define complex routing paths (CGMs, Group Heads, Department Heads, and Division Heads).
*   **Approval Portal:** Provides a secure dashboard for managers to approve, disapprove, or dynamically add/remove reviewers mid-process.
*   **Document Finder:** A real-time tracking tool for requestors to view the exact status and historical audit trail of their documents.
*   **Automated Notifications:** Triggers email alerts to approvers and sends a final "Approval Page" PDF to the requestor upon completion.

---

## Installation Instructions

### 1. Spreadsheet Setup
You require two primary Google Sheets:
*   **Master Database:** Create a sheet named `Conso`. This will store all transaction data, file URLs, and status logs.
*   **Reference Log:** Create a second spreadsheet with two sheets:
    *   `referencenumber`: To track and prevent duplicate Ref IDs.
    *   `DefaultEmail`: To map property selections to specific notification emails.

### 2. Google Drive Setup
*   Create a folder in Google Drive to host all uploaded documents.
*   Copy the **Folder ID** from the URL (the string of characters after `/folders/`).

### 3. Script Deployment
This system consists of four distinct Project Files. For each project (Submission, Approver, Finder, Landing Page):
1.  Open [Google Apps Script](https://script.google.com).
2.  Create a **New Project**.
3.  Copy the corresponding `.html` and `.js` code from the repository into the project editor.
4.  **Crucial:** In the **Submission Project**, go to `Project Settings` > `Script Properties` and add:
    *   `DRIVE_FOLDER_ID`
    *   `SPREADSHEET_ID` (Master Database)
    *   `REF_SHEET_ID` (Reference Log)

### 4. Web App Deployment
1.  Click **Deploy** > **New Deployment**.
2.  Select **Web App**.
3.  Execute As: **Me** (Admin account).
4.  Who has access: **Anyone within [Your Organization]**.
5.  Authorize all requested permissions (Drive, Gmail, and Sheets).

---

## Usage Examples

### Frontend to Backend Communication
The following snippet demonstrates the standard pattern used within this app to send data from the Vue.js frontend to the Google Apps Script backend. Use this pattern for adding new features or custom validations.

```javascript
function handleSubmission(payload) {
    console.log("Preparing to route memo...");

    // This calls the .gs function 'processFormData'
    google.script.run
        .withSuccessHandler((response) => {
            // response contains the generated REF Number
            console.log("Success! Reference Number:", response.refNumber);
            alert("Submission Successful. Your Reference is: " + response.refNumber);
            
            // Logic to clear form or redirect user
            window.top.location.href = "YOUR_LANDING_PAGE_URL";
        })
        .withFailureHandler((error) => {
            console.error("Critical Error:", error.message);
            alert("System Error: " + error.message);
        })
        .processFormData(payload);
}

// Example Payload Object
const myMemo = {
    name: "Jeron Luther",
    email: "jluther@megaworld-lifestyle.com",
    memoSubject: "Equipment Requisition",
    selectedCGM: ["CAMSY", "MIKEE"],
    fileData: "BASE64_ENCODED_STRING", // Decoded on server via Utilities.base64Decode
    fileMimeType: "application/pdf"
};
```

---

## Tech Stack
*   **Backend:** Google Apps Script (V8 Runtime)
*   **Frontend:** Vue.js 3 (Composition API)
*   **Styling:** Bootstrap 5.3 & Bootstrap Icons
*   **Database:** Google Sheets API
*   **Storage:** Google Drive API

---

> **Warning**  
> **- For MCD Internal Use Only**  
> This application contains proprietary workflow logic and credential structures intended for internal organization use. Unauthorized distribution is prohibited.
