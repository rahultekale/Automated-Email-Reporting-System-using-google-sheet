# Automated Email Reporting System

## Overview
This project implements an automated email reporting system using Google Sheets and Google Apps Script. It allows users to send a formatted weekly sales report via email from Google Sheets.

## Features
- Custom menu in Google Sheets
- Confirmation dialog before sending email
- Automated email with HTML formatting

## Setup Instructions

1. **Create a Google Sheet**:
   - Name the sheet `ReportData` and enter data in range `A1:C10`.

2. **Set Up Google Apps Script**:
   - Open Google Sheets and go to **Extensions > Apps Script**.
   - Replace the default content with the provided `code.gs` script.
   - Create an HTML file named `email.html` and paste the provided HTML content.

3. **Authorize and Use**:
   - Authorize the script and use the custom menu to send the email report.

## Deployment
Follow the steps in this repository to set up and use the project.

