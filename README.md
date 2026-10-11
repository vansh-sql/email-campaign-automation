# Vns.WorkSpace — Email Campaign Automation

A web-based email campaign management tool built with Google Apps Script, Google Sheets, and HTML/CSS/JavaScript. It provides a centralized interface for managing email campaigns, working with message templates, handling attachments, and monitoring campaign activity.

## Live Demo

**Application:** [Open Vns.WorkSpace](https://script.google.com/macros/s/AKfycbxy1SJvL3Wj5OGXEwc-1hxdFUiXV2GcgfM2wLwA5lVkEJeBsy1yJVNNq10aAOb_L_zW/exec)

### Demo Login Credentials

- **User ID:** `TEMP-10`
- **Password:** `123456789@`

> **Note:** The demo depends on the deployed application's availability and configuration. Please avoid entering sensitive or real customer information. If email sending is enabled, use only authorized test recipients.

## Overview

Managing email campaigns manually can become time-consuming when recipient lists, message content, attachments, and campaign records are handled separately.

Vns.WorkSpace brings these tasks together in a single web interface. It uses Google Apps Script to connect the application with Google services and Google Sheets to support data management and campaign tracking.

The project focuses on making email campaign workflows easier to manage, reducing repetitive manual work, and improving visibility into campaign activity.

## The Problem

Manual email campaign workflows can involve several repetitive tasks:

- Managing recipient lists and selecting recipients for a campaign.
- Preparing and reusing email content.
- Handling attachments for outgoing messages.
- Keeping track of sent and failed messages.
- Reviewing previous campaign activity.
- Checking campaign performance across different records.

When these tasks are handled separately, it can be harder to maintain consistent records and monitor the overall workflow.

## Project Objectives

The main objectives of Vns.WorkSpace are to:

- Provide a centralized interface for email campaign management.
- Simplify recipient selection and email composition.
- Support reusable message templates.
- Make attachment handling part of the campaign workflow.
- Maintain campaign activity records for review.
- Display key campaign statistics in a dashboard.
- Reduce repetitive work through Google Apps Script automation.

## Key Features

### 1. Email Campaign Management
Create and manage email campaigns through a web-based interface instead of relying entirely on manual processes.

### 2. Recipient Selection
Work with available recipient data and select recipients for campaign delivery.

### 3. Email Templates
Use reusable email content to make recurring communication easier to prepare and maintain.

### 4. Attachment Support
Support the attachment workflow for emails, helping users include relevant files with campaign messages.

### 5. Campaign History
Review recorded campaign activity and message-related information to help track previous operations.

### 6. Dashboard and KPIs
View campaign-related statistics in one place, including metrics for campaigns, sent messages, failed messages, and logged messages where available.

### 7. Sender Display Name
Configure the sender display name for outgoing emails where supported by the application and Google account permissions.

### 8. Google Sheets Integration
Use Google Sheets as the application's data source for relevant records and campaign information.

### 9. Web-Based Interface
Access the application's features through an HTML/CSS/JavaScript interface powered by Google Apps Script.

## How It Works

The application brings together the user interface, Google Apps Script, Google Sheets, and Gmail services.

1. The user opens the web application and signs in.
2. Available recipient data is used to prepare a campaign.
3. The user selects recipients and prepares the email content.
4. Templates and attachments can be used as supported by the workflow.
5. Google Apps Script processes the relevant operations using the configured Google services.
6. Campaign and message activity can be reviewed through the available records and dashboard.

The exact workflow depends on the deployed application's configuration, permissions, and connected spreadsheet data.

## Technology Stack

| Technology | Purpose |
|---|---|
| Google Apps Script | Application logic and Google services integration |
| Google Sheets | Data storage and campaign records |
| Gmail services | Email delivery |
| HTML | Web interface structure |
| CSS | Layout and styling |
| JavaScript | Client-side interactions and interface behavior |
| Google Drive services | File access and attachment handling where configured |

## Business Value

Vns.WorkSpace is designed to make recurring email operations easier to organize and monitor.

- **Less repetitive work:** Brings common campaign tasks into one interface.
- **Centralized workflow:** Keeps recipient selection, message preparation, and campaign tracking together.
- **Reusable communication:** Templates can reduce the effort needed to prepare recurring emails.
- **Improved visibility:** Dashboard metrics and history help users review campaign activity.
- **Simpler record review:** Google Sheets provides a familiar place to manage relevant data.
- **Potential time savings:** Automating repetitive steps can reduce manual effort, depending on campaign size and workflow.

Actual time savings depend on usage, campaign volume, and the application's configuration.

## Setup and Configuration

This project uses Google Apps Script and Google Sheets. To configure your own instance, you will need access to a Google account and the application's source code.

### Prerequisites

- A Google account.
- Access to Google Apps Script.
- A Google Spreadsheet configured for the application's data.
- Gmail access and the required permissions for sending email.
- The project source files.

### Setup Steps

1. **Create or open a Google Spreadsheet.**  
   Prepare the spreadsheet structure required by the application. Use the sheet names and columns expected by the source code.

2. **Open Google Apps Script.**  
   Create a new Apps Script project or open an existing project.

3. **Add the project files.**  
   Copy the relevant server-side code and HTML files into the Apps Script project, preserving the expected filenames and references.

4. **Configure spreadsheet access.**  
   Ensure the script references the correct spreadsheet and that the required sheets and columns exist.

5. **Review permissions.**  
   Check the Google services used by the application and grant only the permissions needed for its intended functionality.

6. **Test the workflow.**  
   Start with a small, controlled test using authorized test recipients. Verify email content, attachments, and campaign records before using the application for real campaigns.

7. **Deploy the web application.**  
   In Apps Script, create a web app deployment and configure access according to your use case and security requirements.

8. **Review the deployment settings.**  
   Confirm who can access the application, which account executes the script, and whether the selected settings are appropriate for email sending and data access.

> **Configuration note:** Exact spreadsheet IDs, sheet names, column names, and deployment settings must match the implementation. Refer to the source code and project setup notes for the required configuration.

## Testing Checklist

Use this checklist when validating a deployment. Check an item only after confirming the behavior.

- [ ] Application opens successfully.
- [ ] Login works with the intended credentials.
- [ ] Recipient data loads correctly.
- [ ] Recipient selection behaves as expected.
- [ ] Email subject and body are handled correctly.
- [ ] Templates can be used as intended.
- [ ] Attachments are included and delivered correctly.
- [ ] Sender display name behaves as configured.
- [ ] Test emails reach the intended test recipients.
- [ ] Sent and failed message records are handled correctly.
- [ ] Campaign history displays the expected information.
- [ ] Dashboard metrics reflect the available records.
- [ ] Unauthorized access and unintended email sending are appropriately restricted.

## Security and Operational Notes

- Do not commit real passwords, private tokens, or sensitive customer data to a public repository.
- Demo credentials included in a public README should be used only for a restricted demo account.
- Restrict email sending to authorized users and appropriate recipients.
- Verify Google Apps Script deployment access and execution settings before sharing the application.
- Test with non-sensitive data and controlled recipients before sending campaign emails.
- Review Google account permissions and applicable sending limits before operational use.
- Avoid using the demo for confidential or production data.

## Limitations

- The application depends on Google Apps Script, Google Sheets, Gmail services, and their permissions.
- Available functionality depends on the configured spreadsheet structure and deployment settings.
- Email delivery can be affected by account permissions, service limits, recipient systems, and spam filtering.
- Dashboard statistics are only as reliable as the records and calculations maintained by the application.
- The live demo's availability and behavior depend on the current deployment configuration.

## Future Improvements

Potential improvements for future versions include:

- More detailed campaign reports and analytics.
- Additional validation for recipient lists and email content.
- Stronger access controls and safer demo-account management.
- More detailed error reporting and delivery diagnostics.
- Improved attachment validation and handling.
- Campaign scheduling and additional automation options.
- Expanded automated tests for important workflows.

These are possible future enhancements, not claims about existing functionality.

## What This Project Demonstrates

Vns.WorkSpace brings together web interface development, Google Apps Script automation, spreadsheet-based data management, and email workflow integration.

It demonstrates how Google Workspace services can be used to organize repetitive business tasks in a centralized application, with a focus on usability, tracking, and operational visibility.

## Project Summary

**Vns.WorkSpace — Email Campaign Automation** is a Google Apps Script-based application designed to simplify email campaign workflows by combining recipient management, reusable content, attachment handling, and campaign tracking in one interface.

The goal is to reduce repetitive manual effort and make campaign activity easier to manage and review.

---

**Project:** Vns.WorkSpace  
**Category:** Email Automation · Workflow Management · Google Workspace  
**Built with:** Google Apps Script, Google Sheets, HTML, CSS, and JavaScript
