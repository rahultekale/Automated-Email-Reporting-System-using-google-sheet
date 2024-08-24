// This function adds a custom menu to the Google Sheets UI
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Email Actions')  // Create a new menu in the spreadsheet
    .addItem('Send Weekly Report', 'confirmAndSendEmail')  // Add an item to the menu
    .addToUi();  // Add the menu to the spreadsheet UI
}

// This function shows a confirmation dialog and sends an email if confirmed
function confirmAndSendEmail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Send Email', 'Are you sure you want to send the weekly report?', ui.ButtonSet.YES_NO);
  
  if (response == ui.Button.YES) {
    sendWeeklyReport();  // Call the function to send the email
    ui.alert('Email sent successfully!');
  } else {
    ui.alert('Email sending canceled.');
  }
}

// This function sends the weekly report via email
function sendWeeklyReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ReportData");
  var dataRange = sheet.getRange("A1:C10").getValues();  // Fetch the data correctly
  var report = generateHtmlReport(dataRange);  // Pass the data to generateHtmlReport function
  
  var recipients = "saavisam123@gmail.com, rahultekale62@gmail.com";  // Replace with actual email addresses
  var subject = "Weekly Sales Report";
  
  // Send the email with the generated HTML report
  GmailApp.sendEmail(recipients, subject, "", {
    htmlBody: report
  });
}

// This function generates the HTML content for the email
function generateHtmlReport(data) {
  var template = HtmlService.createTemplateFromFile('email');  // Load the HTML template
  template.data = data;  // Assign the data to be used in the HTML template
  return template.evaluate().getContent();  // Generate and return the HTML content
}
