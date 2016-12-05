// inserts a custom menu when the spreadsheet opens.
function onOpen() {
  var menu = [{name: 'Send CPD Certificates', functionName: 'sendCertificates_'}];
  SpreadsheetApp.getActive().addMenu('Certificates', menu);
}

// gets info from the sheet about the event and the attendees
// this information ends up in a 2D array 'values'
function sendCertificates_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var range = sheet.getDataRange();
  var values = range.getValues();

  createDocuments(values)

}

function createDocuments(values) {
  var ui = SpreadsheetApp.getUi();
  var result = ui.alert(
    'This will generate CPD certificates for all attendees and save them in a folder in the same directory as this spreadsheet, and then email all the attendees with the PDF version as an attachment. Continue?',
      ui.ButtonSet.YES_NO);

    // Process the user's response.
    if (result == ui.Button.YES) {
      // User clicked "Yes".
      ui.alert('Generating & emailing CPD certificates');
    } else {
      // User clicked "No" or X in the title bar.
      ui.alert('Aborted');
      return;
  }

  // **** specify here the ID of the template document to be used
  var templateDocId = '1RlQc4RySPpZELd3sX7SF0eDnGfLOdFaBlxDzeLbpowA';

  // **** specify here the ID of the folder to put new CPD Certificates subfolders
  var currentFolder = DriveApp.getFolderById('0B1bx-pvb5NfjfjZKYmdBNUxsOVJMYUxKVFlXWHpqUzgtaUJEdkdWWExwSHFtbW1pNms2VFU');

  // opens the CPD Template document by its ID in your Google Drive (returns a Document object)
  var templateDocument = DocumentApp.openById(templateDocId);

  // gets the File as well (returns a File object)
  var templateFile = DriveApp.getFileById(templateDocId);

  // get eventName, eventDate, cpdPoints, and the emailSubject and emailMessageBody for the event and set up variables
  var eventName = values[1][6].toString();
  var eventDate = values[1][7].toString().split(/\d\d:\d\d:\d\d/)[0]; //splits out just the day/month/year from the event DateTime object, & stringifies
  var cpdPoints = values[1][8].toString();
  var emailSubject = values[1][9].toString();
  var emailMessageBody = values[1][10].toString();

  // create a new folder to contain the new CPD certificates, name the folder by the event and date
  var newFolder = currentFolder
    .createFolder('CPD Certificates - ' + eventName + " - " + eventDate);


  // iterate through the list of values in the spreadsheet, creating a new certificate file for each, and mailmerging the information
  for (var i = 1; i < values.length; i++) {
    var rowValues = values[i];
    // make copy of template file, name it DHI_CPD_Certificate, Title, FirstName, Lastname
    var newDocumentName = 'digitalhealth.net CPD Certificate - ' + values[i][0] + " " + values[i][1] + " " + values[i][2];
    var workingFile = templateFile.makeCopy(newDocumentName + ".doc", newFolder);
    var workingDocument = DocumentApp.openById(workingFile.getId()); //gets the document object
    var workingDocumentBody = workingDocument.getBody();             // gets the body text of the document object

    // does all the mail merge stuff, stripping newline characters which are superfluous

    workingDocumentBody.replaceText('{personTitle}', rowValues[0].replace(/(\r\n|\n|\r)/gm,""));
    workingDocumentBody.replaceText('{personFirstName}', rowValues[1].replace(/(\r\n|\n|\r)/gm,""));
    workingDocumentBody.replaceText('{personLastName}', rowValues[2].replace(/(\r\n|\n|\r)/gm,""));
    workingDocumentBody.replaceText('{personEmail}', rowValues[3].replace(/(\r\n|\n|\r)/gm,""));
    workingDocumentBody.replaceText('{personRole}', rowValues[4].replace(/(\r\n|\n|\r)/gm,""));
    workingDocumentBody.replaceText('{personOrganisation}', rowValues[5].replace(/(\r\n|\n|\r)/gm,""));
    workingDocumentBody.replaceText('{eventName}', eventName.replace(/(\r\n|\n|\r)/gm,""));
    workingDocumentBody.replaceText('{eventDate}', eventDate.replace(/(\r\n|\n|\r)/gm,""));
    workingDocumentBody.replaceText('{cpdPoints}', cpdPoints.replace(/(\r\n|\n|\r)/gm,""));
    workingDocument.saveAndClose();

    // generate a PDF version of the file
    var pdfFile = DriveApp.createFile(workingFile.getAs("application/pdf"));
    newFolder.addFile(pdfFile); // move it into the right folder
    DriveApp.removeFile(pdfFile); //delete it from the Drive root folder

    // create attachment payload for email
    var pdfBlob = DriveApp.getFileById(pdfFile.getId()).getBlob().getBytes()
    var attachment = {fileName: newDocumentName + ".pdf", content: pdfBlob, mimeType:'application/pdf'};

    // Send the freshly constructed email
    var emailTo = rowValues[3].replace(/(\r\n|\n|\r)/gm,"");
    var subject = emailSubject.replace(/(\r\n|\n|\r)/gm,"");
    var message = emailMessageBody.replace(/(\r\n|\n|\r)/gm,"");
    MailApp.sendEmail(emailTo, subject, message, {attachments:[attachment]});

  }
} // createDocuments()
