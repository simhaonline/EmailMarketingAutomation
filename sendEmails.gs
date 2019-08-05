var EMAIL_SENT = 'EMAIL_SENT';

function sendEmails() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2; 
  var numRows = 2; 

  var dataRange = sheet.getRange(startRow, 1, numRows, 4);

  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var fileName = row[0] + ' presentation.pdf';
    var file = DriveApp.getFilesByName(fileName).next()
    var emailAddress = row[1]; 
    var message = row[2]; 
    var emailSent = row[3]; 
    if (emailSent != EMAIL_SENT) { 
      var subject = 'Sending emails from a Spreadsheet';
      MailApp.sendEmail(emailAddress, subject, message, {attachments: file});
      sheet.getRange(startRow + i, 4).setValue(EMAIL_SENT);

      SpreadsheetApp.flush();
    }
  }
}
