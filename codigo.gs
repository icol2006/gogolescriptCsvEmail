//Import data.csv file
function importCSVFromGoogleDrive() {
  var file = DriveApp.getFoldersByName("CSVImport").next().getFilesByName("data.csv").next(); // reports_folder_id = id of folder where csv reports are saved
  //var file = fSource.getFilesByName("data.csv").next();
  var csvData = Utilities.parseCsv(file.getBlob().getDataAsString());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Automatic Import"));
  var sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);
  
}

//Create menu
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('User scripts')
  .addItem('Import Csv', 'menuimportcsv')
  .addItem('Send Email', 'doGet')
  .addToUi();
}

//Action menu import csv file
function menuimportcsv() {
  importCSVFromGoogleDrive()
}

//Action menu send email 
function menusendemail() {
  sendemail();
}

//Send email
function sendemail(form) {
  
  var spreadsheet   = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheetId = spreadsheet.getId();  
  var file          = Drive.Files.get(spreadsheetId);  
  var url           = file.exportLinks[MimeType.MICROSOFT_EXCEL];
  var token         = ScriptApp.getOAuthToken();
  
  
  var response      = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  token
    }
  });
  
  var fileName = (spreadsheet.getName()) + '.xlsx';
  var blobs   = [response.getBlob().setName(fileName)];
  
  var listemails = form.emailtosend.split(",");
  var message = form.message;
  var subject = form.subject;

  for each (var email in listemails)
  {
    MailApp.sendEmail(email, subject, message, {attachments: blobs});
  }
  
  validate();
    
}

function validate() {
  Browser.msgBox('Email was sent!', Browser.Buttons.OK);  // See: http://www.mousewhisperer.co.uk/drivebunny/message-dialogs-in-apps-script/
}

///////////////
/**
* Displays an HTML-service dialog in Google Sheets that contains client-side
* JavaScript code for the Google Picker API.
*/
function showPicker() {
  var html = HtmlService.createHtmlOutputFromFile('dialog.html')
  .setWidth(600)
  .setHeight(425)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a file');
}


function doGet() {
  var html= HtmlService.createHtmlOutputFromFile('index')
  .setWidth(600)
  .setHeight(425)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Information');
}

