# MCD Memo Routing Web App

A comprehensive web application built with Google Apps Script for managing and routing MCD (document) memos through an approval workflow system.

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Project Structure](#project-structure)
- [Technologies Used](#technologies-used)
- [Components](#components)
- [Getting Started](#getting-started)
- [Usage Guide](#usage-guide)
- [File Descriptions](#file-descriptions)
- [Key Functions](#key-functions)
- [Database Schema](#database-schema)
- [Workflow](#workflow)
- [Contributing](#contributing)
- [License](#license)

## 🎯 Overview

The MCD Memo Routing Web App is a sophisticated document management and routing system designed to streamline the approval process for memos and documents. Built using Google Apps Script, HTML5, and Google Sheets integration, this application provides a user-friendly interface for submitting, tracking, and approving memos through a defined workflow.

## ✨ Features

- **Multi-Page Application**: Separate interfaces for landing, submission, finder, and approval pages
- **Document Submission**: Easy-to-use form for submitting new memos
- **Document Search & Retrieval**: Find and view previously submitted documents
- **Approval Workflow**: Streamlined approval process with status tracking
- **Real-time Status Updates**: Track memo routing status in real-time
- **Database Integration**: Seamless integration with Google Sheets as a database backend
- **User-Friendly Interface**: Responsive HTML/CSS frontend for optimal user experience
- **Metadata Management**: Track document dates, approvers, and routing history

## 📁 Project Structure

```
MCD-Memo-Routing-App/
├── Landing Page.gs                          # Main entry point & page routing
├── Landing Page.html                        # Landing page interface
├── Submission Page.gs                       # Document submission backend
├── Submission Page.html                     # Submission form interface
├── Submission Page.json                     # Submission page configuration
├── Approver Page.gs                         # Approval workflow backend
├── Approver Page.html                       # Approver interface
├── Finder Page.gs                          # Document search backend
├── Finder Page.html                        # Document finder interface
├── Memo Routing(Database) Functions         # Core database operations
├── Google Apps Script Helper                # Utility functions
├── Routing Status Formula                   # Status calculation logic
├── appsscript.json                         # Google Apps Script configuration
├── Documentation_ MCD Document Routing System (1).pdf
└── README.md                                # This file
```

## 🛠️ Technologies Used

- **Google Apps Script**: Backend logic and Google Sheets integration
- **HTML5**: User interface markup
- **CSS3**: Styling and responsive design
- **JavaScript (GAS)**: Frontend interactivity and form handling
- **Google Sheets**: Database backend for storing memos and routing data
- **JSON**: Configuration and data format

## 🔧 Components

### Landing Page
The entry point of the application. Provides navigation to different sections:
- **File**: `Landing Page.gs` (244 bytes)
- **File**: `Landing Page.html` (6,502 bytes)
- **Purpose**: Main application interface and navigation hub

### Submission Page
Allows users to submit new memos:
- **Backend**: `Submission Page.gs` (23,583 bytes)
- **Frontend**: `Submission Page.html` (25,643 bytes)
- **Config**: `Submission Page.json` (485 bytes)
- **Features**:
  - Form validation
  - Document metadata capture
  - Database storage
  - Confirmation feedback

### Finder Page
Search and retrieve previously submitted documents:
- **Backend**: `Finder Page.gs` (13,465 bytes)
- **Frontend**: `Finder Page.html` (22,822 bytes)
- **Features**:
  - Advanced search functionality
  - Filter by date, approver, status
  - Document preview
  - Document details display

### Approver Page
Interface for reviewing and approving memos:
- **Backend**: `Approver Page.gs` (32,899 bytes)
- **Frontend**: `Approver Page.html` (73,089 bytes)
- **Features**:
  - Document review queue
  - Approval/rejection workflow
  - Comments and feedback
  - Signature capture
  - Status updates

## 🚀 Getting Started

### Prerequisites

1. Google Account with access to Google Apps Script
2. Google Sheets for database storage
3. Proper permissions to create and manage Google Apps Script projects

### Installation

1. **Create a New Google Apps Script Project**
   - Go to [script.google.com](https://script.google.com)
   - Create a new project
   - Name it "MCD Memo Routing App"

2. **Add Files to Your Project**
   - Copy `Landing Page.gs` and `Landing Page.html` to your project
   - Add all other `.gs` (Google Apps Script) files
   - Add all `.html` (HTML page files)
   - Add `appsscript.json` configuration file

3. **Create Supporting Google Sheet**
   - Create a new Google Sheet to serve as your database
   - Create sheets for: Memos, Approvals, Users, Status Log
   - Update database references in the code with your sheet ID

4. **Configure App Settings**
   - Update `appsscript.json` with correct manifest settings
   - Configure OAuth scopes if needed
   - Set up any required triggers

5. **Deploy as Web App**
   - In Google Apps Script editor, click "Deploy" → "New Deployment"
   - Select type: "Web App"
   - Execute as: Your account
   - Who has access: Appropriate user group
   - Copy the deployment URL

### Environment Setup

```javascript
// Update the following in your scripts:
const SHEET_ID = "YOUR_GOOGLE_SHEET_ID";
const SPREADSHEET = SpreadsheetApp.openById(SHEET_ID);
const MEMO_SHEET = SPREADSHEET.getSheetByName("Memos");
```

## 📖 Usage Guide

### For Document Submitters

1. **Access the Application**
   - Open the web app URL in your browser
   - You'll land on the Landing Page

2. **Submit a Memo**
   - Click "Submit New Memo"
   - Fill in the required fields:
     - Memo Title
     - Content/Body
     - Department
     - Routing Approvers
     - Priority Level
   - Click "Submit"
   - Receive confirmation with tracking number

3. **Track Your Memo**
   - Go to "Finder"
   - Search by tracking number or date range
   - View current routing status

### For Approvers

1. **Access Approval Queue**
   - Open the app as an approver
   - Navigate to "Approver Dashboard"
   - View pending memos requiring your approval

2. **Review & Approve**
   - Click on a memo to review
   - Read the content and metadata
   - Add comments if needed
   - Click "Approve" or "Reject"
   - Memo automatically routes to next approver or completes

3. **View Approval History**
   - Check completed approvals
   - View audit trail of all actions
   - Download approval records if needed

## 📄 File Descriptions

### Backend Files (.gs)

| File | Size | Purpose |
|------|------|---------|
| `Landing Page.gs` | 244 B | Main entry point; initializes HTML interface |
| `Submission Page.gs` | 23.5 KB | Handles memo submission form logic |
| `Approver Page.gs` | 32.9 KB | Manages approval workflow operations |
| `Finder Page.gs` | 13.5 KB | Implements search and filter functionality |
| `Memo Routing(Database) Functions` | 79.7 KB | Core database operations (CRUD) |
| `Google Apps Script Helper` | 57.7 KB | Utility functions for common operations |

### Frontend Files (.html)

| File | Size | Purpose |
|------|------|---------|
| `Landing Page.html` | 6.5 KB | Main navigation and landing interface |
| `Submission Page.html` | 25.6 KB | Memo submission form with validation |
| `Finder Page.html` | 22.8 KB | Document search and filter interface |
| `Approver Page.html` | 73.1 KB | Approval workflow user interface |

### Configuration Files

| File | Size | Purpose |
|------|------|---------|
| `appsscript.json` | 423 B | Google Apps Script manifest and configuration |
| `Submission Page.json` | 485 B | Form field configuration |

### Documentation

| File | Size | Purpose |
|------|------|---------|
| `Routing Status Formula` | 1.96 KB | Status calculation logic |
| `Documentation_ MCD Document Routing System (1).pdf` | 101.4 KB | Complete system documentation |

## 🔑 Key Functions

### Landing Page Functions

```javascript
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('MCD Memo Routing')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
```

### Database Operations

The `Memo Routing(Database) Functions` file includes:

- `addMemo(memoData)` - Create new memo record
- `getMemo(memoId)` - Retrieve specific memo
- `updateMemo(memoId, updates)` - Update existing memo
- `deleteMemo(memoId)` - Delete memo record
- `getAllMemos()` - Retrieve all memos
- `getMemosByStatus(status)` - Filter memos by status
- `getMemosByDateRange(startDate, endDate)` - Filter by date range
- `routeMemo(memoId, nextApprover)` - Route to next approver

### Helper Functions

The `Google Apps Script Helper` file includes:

- `getSheet(sheetName)` - Get reference to sheet
- `appendRow(sheet, data)` - Add new row to sheet
- `updateRow(sheet, rowIndex, data)` - Update existing row
- `findRows(sheet, criteria)` - Search rows by criteria
- `sendEmail(recipient, subject, body)` - Send notifications
- `formatDate(date)` - Format dates consistently

### Status Functions

The `Routing Status Formula` calculates:

- Current approval status
- Next approver in workflow
- Days in each approval stage
- Overall routing progress

## 📊 Database Schema

### Memos Sheet

| Column | Type | Description |
|--------|------|-------------|
| ID | String | Unique memo identifier |
| Title | String | Memo title |
| Content | String | Memo body content |
| Submitter | String | User who submitted memo |
| Department | String | Department/Division |
| Date Submitted | Date | Submission timestamp |
| Priority | String | High/Medium/Low |
| Status | String | Current workflow status |
| Current Approver | String | Next approver in queue |
| Routing Path | String | Comma-separated approver list |

### Approvals Sheet

| Column | Type | Description |
|--------|------|-------------|
| Memo ID | String | Reference to memo |
| Approver | String | Approver name |
| Action | String | Approved/Rejected/Pending |
| Date | Date | Action timestamp |
| Comments | String | Approval comments |
| Signature | String | Digital signature |

## 🔄 Workflow

```
1. SUBMISSION
   User submits memo with details
   └─> Form validation
   └─> Database storage
   └─> Confirmation email

2. ROUTING
   Memo assigned to first approver
   └─> Status: "Pending Approver 1"
   └─> Notification sent to approver

3. APPROVAL
   Approver reviews memo
   └─> Decision: Approve/Reject/Request Changes
   └─> Comments and signature added
   └─> Status updated

4. NEXT STAGE
   If approved, route to next approver
   └─> Repeat approval process
   └─> If rejected, return to submitter
   └─> If final approval, mark complete

5. COMPLETION
   Final approval granted
   └─> Status: "Approved"
   └─> Notification sent to submitter
   └─> Archived in database
```

## 🔒 Security Considerations

- **Authentication**: Integrated with Google account system
- **Authorization**: Role-based access (Submitters, Approvers, Admins)
- **Data Protection**: Google Sheets encryption in transit and at rest
- **Audit Trail**: All actions logged with timestamps and user info
- **Email Notifications**: Secure communication throughout workflow

## 🐛 Troubleshooting

### Common Issues

**Issue**: App won't load
- **Solution**: Verify deployment URL is correct; redeploy if needed
- Check browser console for errors (F12)

**Issue**: Form submission fails
- **Solution**: Verify Google Sheet ID is correct in code
- Check that user has edit permissions to the sheet

**Issue**: Status not updating
- **Solution**: Review Routing Status Formula for errors
- Verify approver email addresses are correct

**Issue**: Emails not sending
- **Solution**: Confirm Google Apps Script has Gmail permission
- Check that recipient email addresses are valid

## 📝 Configuration

### Update App Settings

Edit these constants in your `.gs` files:

```javascript
// Database
const SHEET_ID = "YOUR_SHEET_ID";
const MEMO_SHEET_NAME = "Memos";
const APPROVAL_SHEET_NAME = "Approvals";

// Email
const ADMIN_EMAIL = "admin@company.com";
const EMAIL_SUBJECT_PREFIX = "[MCD Routing]";

// Workflow
const APPROVAL_ROLES = ["Manager", "Director", "VP"];
const AUTO_SEND_EMAIL = true;
```

## 🤝 Contributing

To contribute improvements:

1. Create a new branch for your feature
2. Make your changes
3. Test thoroughly with sample data
4. Document any new functions
5. Submit a pull request with description

## 📚 Related Documentation

- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Google Sheets API Guide](https://developers.google.com/sheets/api)
- [HTML Service Documentation](https://developers.google.com/apps-script/guides/html)
- See also: `Documentation_ MCD Document Routing System (1).pdf` in repository

## 📞 Support

For issues or questions:

1. Check the included PDF documentation
2. Review inline code comments
3. Consult Google Apps Script debugging tools
4. Contact your system administrator

## 📈 Future Enhancements

- [ ] Add bulk memo processing
- [ ] Implement advanced analytics dashboard
- [ ] Add document attachment support
- [ ] Create mobile-optimized interface
- [ ] Implement task reminders
- [ ] Add workflow template system
- [ ] Create approval timeline visualization
- [ ] Add integration with Google Drive for document storage

## 📄 License

This project is provided as-is for internal use within MCD organization.

## 🗂️ Additional Files

This repository also contains:

- **Documentation_ MCD Document Routing System (1).pdf** (101.4 KB)
  - Complete system documentation
  - Setup instructions
  - Detailed workflow diagrams
  - Screenshots and user guides

## 📊 Repository Statistics

- **Repository**: mrjeronluther/MCD-Memo-Routing-App
- **Created**: July 14, 2025
- **Language**: HTML/JavaScript (Google Apps Script)
- **Visibility**: Private
- **Open Issues**: 4
- **Size**: 428 KB
- **Last Updated**: February 9, 2026

---

**Last Updated**: February 9, 2026  
**Version**: 1.0  
**Author**: mrjeronluther

For the latest updates and complete documentation, visit the [GitHub repository](https://github.com/mrjeronluther/MCD-Memo-Routing-App).
