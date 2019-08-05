function copySlides() {
  var dataSpreadsheetId = 'YOUR_SPREADSHEET_ID'
  var templatePresentationId = 'YOUR_SLIDE_ID'
  var dataRangeNotation = 'YOUR_SHEET_NAME!YOUR_RANGE';
  var values = SpreadsheetApp.openById(dataSpreadsheetId).getRange(dataRangeNotation).getValues();
  
  for (var i = 0; i < values.length; ++i) {
    var row = values[i];
    var customerName = row[0];
    
    var copyTitle = customerName + ' presentation';
    var copyFile = {
      title: copyTitle,
      parents: [{id: 'root'}]
    };
    copyFile = Drive.Files.copy(copyFile, templatePresentationId);
    var presentationCopyId = copyFile.id;
    
    requests = [{
      replaceAllText: {
        containsText: {
          text: '{{YOUR_TAG}}',
          matchCase: true
        },
        replaceText: customerName
      }
    }];
    
    var result = Slides.Presentations.batchUpdate({
      requests: requests
    }, presentationCopyId);
    var numReplacements = 0;
    result.replies.forEach(function(reply) {
      numReplacements += reply.replaceAllText.occurrencesChanged;
    });
    console.log('Created presentation for %s with ID: %s', customerName, presentationCopyId);
    console.log('Replaced %s text instances', numReplacements);
    var blob = DriveApp.getFileById(presentationCopyId).getBlob();
    DriveApp.createFile(blob);
  }
}
