function startForm() {
  var form = HtmlService.createHtmlOutputFromFile('FormHTML').setWidth(700).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(form, ' ');
}

function finishForm(trainer, name, fin, charID, forumID) {
  var loading = HtmlService.createHtmlOutputFromFile('LoadingHTML').setWidth(700).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(loading, ' ');
  createDocument(trainer, name, fin, charID, forumID);
}

function createDocument(trainer, name, fin, charID, forumID) {
  var template = "REMOVED";
  var documentId = DriveApp.getFileById(template).makeCopy().getId();
  var docPerms = DriveApp.getFileById(documentId);
  var file = DriveApp.getFileById(documentId);
  DriveApp.getFileById(documentId).setName('Officer Development Plan | '+name+' | '+fin+' | '+charID);
  docPerms.setOwner('REMOVED');
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  editDocument(trainer, name, fin, charID, forumID, documentId);
}

function editDocument(trainer, name, fin, charID, forumID, documentId) {
  var date = Utilities.formatDate(new Date(), 'GMT', 'dd-MM-yy');
  var newss = SpreadsheetApp.openById(documentId);
  var sheet = newss.getSheetByName('Officer Development Plan');
  sheet.getRange("B6").setValue(name);
  sheet.getRange("H6").setValue(fin);
  sheet.getRange("E6").setValue(date);
  sheet.getRange("I6").setValue(charID);
  sheet.getRange("L6").setValue(forumID);
  sheet.getRange("J13").setValue('TRUE');
  sheet.getRange("K13").setValue(date);
  sheet.getRange("M13").setValue(trainer);
  completeDocument(documentId);
}

function completeDocument(documentId) {
  var complete = HtmlService.createTemplateFromFile('CompleteHTML');
  complete.data = 'https://docs.google.com/spreadsheets/d/'+documentId
  var htmlOutput = complete.evaluate().setWidth(700).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, ' ');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
