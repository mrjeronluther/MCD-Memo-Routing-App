# MCD-Memo-Routing-Web-App

Of course. Here is a comprehensive documentation for the "MCD Document Routing" application built on Google Apps Script.

Documentation: MCD Document Routing System
1. Overview

The MCD Document Routing System is a comprehensive, custom-built web application powered by Google Apps Script designed to digitize and automate the internal memo and document approval process. It replaces manual routing with a streamlined, traceable, and efficient digital workflow.

The system consists of several interconnected web app modules that allow users to submit documents, approvers to review and action them, and everyone to track the status of submissions in real-time. All data is securely stored and managed within your Google Workspace environment (Google Sheets and Google Drive).

2. Key Features

Centralized Document Submission: A user-friendly form for uploading documents (PDFs, images) and providing all necessary metadata.

Dynamic Multi-level Approval Routing: Senders can select the specific Commercial General Managers (CGMs), Department Heads, and Division Heads required for approval on a per-document basis.

Secure, Role-Based Approval Portal: Approvers log in to a dedicated interface where they see only the documents pending their review.

Privileged Administrative Actions: Authorized users can add or remove reviewers from a memo's workflow on the fly.

Real-time Status Tracking: A dedicated "Document Finder" allows the original submitter to track the latest status of their memo, including who has approved it and whose approval is still pending.

Automated Email Notifications: The system automatically sends email notifications upon submission, when new reviewers are added, and when a memo is fully approved or disapproved.

Automated PDF Generation: Once a memo's lifecycle is complete (Approved or Disapproved), the system automatically generates a summary PDF page containing all approver signatures, timestamps, and remarks.

Optimized Backend Processing: Nightly automated scripts efficiently process statuses for thousands of entries, ensuring the master tracking sheet is always up-to-date without manual intervention.

3. Technology Stack

This application is built entirely within the Google Workspace ecosystem, making it a powerful and cost-effective solution.

Backend Logic: Google Apps Script (GAS) - The server-side code that powers the entire application, handles data processing, and integrates with other Google services.

Frontend Framework: Vue.js 3 (Composition API) - A modern JavaScript framework used to create a reactive, dynamic, and responsive user interface for the web apps.

Styling: Bootstrap 5 - A popular CSS framework used for clean, modern, and mobile-friendly styling of the application's interface.

Database: Google Sheets - Used as the primary database to store all memo metadata, tracking information, status updates, and user credentials.

File Storage: Google Drive - Used to securely store all uploaded memo documents and the system-generated PDF approval/disapproval pages.

Notifications: Gmail (via MailApp Service) - Used to send all automated email notifications to users and approvers.

4. System Components

The application is modular, with each part handling a specific function.

Component	Files	Purpose
Landing Page	LANDING PAGE.HTML, LANDING PAGE.GS	The main entry point for all users. It provides a simple, clean interface with cards linking to the other three modules of the application.
Submission Web App	SUBMISSION PAGE.HTML, SUBMISSION PAGE.GS	A detailed form for users to submit new documents. It captures requestor details, memo subject, coverage, the specific approval path, and handles the file upload.
Approval Web App	APPROVER PAGE.HTML, APPROVER PAGE.GS	A secure, login-protected portal for managers. Approvers can view pending documents, approve or disapprove them with remarks, and (if authorized) modify the reviewer list.
Document Finder App	FINDER.HTML, FINDER.GS	A secure, login-protected tool for the document submitter. They can search by reference number to see the current approval status, who has approved it, and who is still pending.
Google Sheets (Backend)	N/A (Asset)	Acts as the application's database. It includes a Conso sheet (master log), referencenumber sheet, and individual sheets for each approver (e.g., GMC, RCS) where their specific queues are managed.
Backend Helper Scripts	Google Apps Script Helper code (in any .gs file)	A collection of functions that are not part of the web apps but run on triggers or from a custom menu in the Sheet. They handle PDF generation, status checks, and sending final notification emails.
5. Setup and Deployment Guide

Follow these steps to set up and deploy the MCD Document Routing System.

Step 1: Create Google Workspace Assets

Google Sheet:

Create a new Google Sheet. This will be your main database.

Rename the default sheet to Conso. This will be your master tracking log.

Create new sheets and name them for each individual approver code (e.g., GMC, RCS, MGL, CAMSY, JAYE, etc.).

Create another new Google Sheet (or use a separate one for security). In this sheet, create a tab named referencenumber. This will be used to log and ensure unique reference numbers.

Important: Note the ID of both spreadsheets. The ID is the long string of characters in the URL (e.g., https://docs.google.com/spreadsheets/d/[SPREADSHEET_ID]/edit).

Google Drive:

Create a new folder in Google Drive. This folder will store all uploaded memos and generated PDF reports.

Note the ID of this folder. The ID is the last part of the folder's URL (e.g., https://drive.google.com/drive/folders/[FOLDER_ID]).

Step 2: Create the Google Apps Script Project

Open the main Google Sheet you created in Step 1.

Go to Extensions -> Apps Script.

A new Apps Script project editor will open, linked to your spreadsheet.

Step 3: Create and Populate Script Files

In the Apps Script editor, create the following files and copy-paste the corresponding code into each one.

LANDING PAGE.gs and LANDING PAGE.HTML

SUBMISSION PAGE.gs and SUBMISSION PAGE.HTML (Rename the HTML file from index to SUBMISSION PAGE.HTML in the editor)

APPROVER PAGE.gs and APPROVER PAGE.HTML (Rename the HTML file from index to APPROVER PAGE.HTML)

FINDER.gs and FINDER.HTML (Rename the HTML file from index to FINDER.HTML)

BackendHelper.gs (Create this new .gs file and copy the entire Google Apps Script Helper block into it).

Note: In each .gs file, ensure the doGet() function references the correct HTML filename (e.g., HtmlService.createHtmlOutputFromFile('SUBMISSION PAGE')).

Step 4: Configure Script Properties

This is a critical step for the Submission App to work.

In the Apps Script editor, click the Project Settings (gear icon) on the left sidebar.

Scroll down to the Script Properties section and click Add script property.

Add the following three properties with the IDs you noted in Step 1:

Property Name: DRIVE_FOLDER_ID

Value: [Paste the Google Drive Folder ID here]

Property Name: SPREADSHEET_ID

Value: [Paste the main Google Sheet ID here]

Property Name: REF_SHEET_ID

Value: [Paste the Google Sheet ID containing the referencenumber tab here]

Step 5: Deploy the Web Apps

You must deploy each of the four web apps separately.

In the Apps Script editor, click the blue Deploy button and select New deployment.

Click the gear icon next to "Select type" and choose Web app.

Configure the deployment:

Description: e.g., "MCD Submission Form v1"

Execute as: Me (Your Google Account)

Who has access: Anyone within [Your Organization] (or "Anyone" if it's for external users).

Click Deploy.

Authorize the script permissions when prompted.

Copy the Web app URL. This is the live URL for your app.

Repeat this process for all four modules (Landing, Submission, Approval, Finder). You will get four unique URLs.

Step 6: Update the Landing Page Links

Open the LANDING PAGE.HTML file in the Apps Script editor.

Find the memoActions array in the <script> tag.

Replace the placeholder link values with the new deployment URLs you got in the previous step.

      Generated javascript
      const memoActions = ref([
          {
              title: 'Submit a Document',
              description: '...',
              link: 'PASTE_YOUR_SUBMISSION_APP_URL_HERE', // <-- Update this
              button: 'Open Form',
          },
          {
              title: 'Approve a Document',
              description: '...',
              link: 'PASTE_YOUR_APPROVAL_APP_URL_HERE', // <-- Update this
              button: 'Approve Now',
          },
          {
              title: 'Document Finder',
              description: '...',
              link: 'PASTE_YOUR_FINDER_APP_URL_HERE', // <-- Update this
              button: 'Search Document',
          },
      ]);


Redeploy the Landing Page to apply the changes (Deploy -> Manage Deployments -> Edit -> New Version).

Step 7: Configure Backend Triggers

In the Apps Script editor, click the Triggers (clock icon) on the left sidebar.

Click + Add Trigger.

Set up the following triggers to automate the nightly processing:

Trigger 1 (Final Status):

Choose which function to run: runAutomatedProcess

Choose which deployment should run: Head

Select event source: Time-driven

Select type of time based trigger: Day timer

Select time of day: Midnight to 1am

Trigger 2 (Who Disapproved):

Function: runAutomatedProcessone | Type: Day timer | Time: 2am to 3am

Trigger 3 (Disapproval Reason):

Function: runAutomatedProcesstwo | Type: Day timer | Time: 4am to 5am

Your application is now fully set up and operational! You can share the link to the Landing Page with your users.
